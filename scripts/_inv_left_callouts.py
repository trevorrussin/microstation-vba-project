"""Inventory left-side callouts below the table in fea_tsp."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch


def text_of(el):
    try:
        if el.IsTextElement:
            return el.AsTextElement.Text
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            return "\n".join(
                tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)
            )
    except Exception:
        pass
    return None


def walk(el, out, depth=0):
    try:
        t = text_of(el)
        if t and t.strip():
            r = el.Range
            out.append(
                {
                    "id": el.ID.Low,
                    "type": int(el.Type),
                    "text": t,
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                    "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                    "lvl": el.Level.Name if el.Level else "?",
                    "is_tn": bool(el.IsTextNodeElement),
                    "is_te": bool(el.IsTextElement),
                    "depth": depth,
                }
            )
        if int(el.Type) in (2, 6, 4, 7) and depth < 4:
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk(sub.Current, out, depth + 1)
            except Exception:
                pass
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    mr = app.ActiveModelReference
    print("file", app.ActiveDesignFile.FullName)
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts = []
    while ee.MoveNext():
        walk(ee.Current, texts)

    # Left side: lower X. Table was around left; callouts below table.
    # Sort by X to find left cluster.
    texts.sort(key=lambda x: x["cx"])
    print(f"total texts={len(texts)}")
    print("--- leftmost 40 by cx ---")
    for t in texts[:40]:
        head = t["text"].replace("\n", " | ")[:70]
        print(
            f"  id={t['id']} cx={t['cx']:.0f} cy={t['cy']:.0f} "
            f"[{t['lx']:.0f}..{t['hx']:.0f},{t['ly']:.0f}..{t['hy']:.0f}] "
            f"lvl={t['lvl']} tn={t['is_tn']} {head!r}"
        )

    # Callout-like: multi-line or long text, TCPP / ny
    callouts = [
        t
        for t in texts
        if (
            t["is_tn"]
            or "TRAFFIC" in t["text"].upper()
            or "CONDUIT" in t["text"].upper()
            or "CABLE" in t["text"].upper()
            or "PULL" in t["text"].upper()
            or "MAST" in t["text"].upper()
            or "SIDEWALK" in t["text"].upper()
            or "APS" in t["text"].upper()
            or "POLE" in t["text"].upper()
        )
        and t["cx"] < 1029450  # left of plan center-ish
    ]
    print(f"\n--- left callout-ish ({len(callouts)}) ---")
    for t in sorted(callouts, key=lambda x: -x["cy"]):
        head = t["text"].replace("\n", " | ")[:80]
        print(
            f"  id={t['id']} cy={t['cy']:.1f} cx={t['cx']:.1f} "
            f"box=[{t['lx']:.1f},{t['ly']:.1f}]-[{t['hx']:.1f},{t['hy']:.1f}] "
            f"{head!r}"
        )


if __name__ == "__main__":
    main()
