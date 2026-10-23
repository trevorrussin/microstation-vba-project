"""Compute union bbox of left-below-table callout texts; preview."""
from __future__ import annotations

import re
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
            return "\n".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))
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
                    "text": t,
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                    "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                    "lvl": el.Level.Name if el.Level else "?",
                    "is_tn": bool(el.IsTextNodeElement),
                }
            )
        if int(el.Type) in (2, 6, 4, 7) and depth < 5:
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
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts = []
    while ee.MoveNext():
        walk(ee.Current, texts)

    # Left callout column below table:
    # right edge ~1029432, Y from ~239620 to ~239775 (under mast-arm block)
    # Exclude mid-plan left callouts (cy > 239850)
    stack = [
        t
        for t in texts
        if t["hx"] <= 1029440
        and t["lx"] >= 1029280
        and 239615 <= t["cy"] <= 239780
        and len(t["text"].strip()) > 3
        # skip tiny qty chips / single letters
        and not re.fullmatch(r"\d+", t["text"].strip())
    ]

    # Prefer callout body TextNodes + their TE overlays; drop pure table qty if any
    print(f"candidates={len(stack)}")
    for t in sorted(stack, key=lambda x: -x["hy"]):
        head = t["text"].replace("\n", " | ")[:70]
        print(
            f"  id={t['id']} cy={t['cy']:.1f} "
            f"[{t['lx']:.1f}..{t['hx']:.1f}, {t['ly']:.1f}..{t['hy']:.1f}] "
            f"tn={t['is_tn']} {head!r}"
        )

    if not stack:
        return
    lx = min(t["lx"] for t in stack)
    ly = min(t["ly"] for t in stack)
    hx = max(t["hx"] for t in stack)
    hy = max(t["hy"] for t in stack)
    print(f"\nUNION [{lx:.2f},{ly:.2f}]-[{hx:.2f},{hy:.2f}] w={hx-lx:.1f} h={hy-ly:.1f}")

    # sample existing shapes near this area for template symbology
    ee = mr.Scan(o)
    shapes = []
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsShapeElement:
                continue
            r = el.Range
            w = float(r.High.X) - float(r.Low.X)
            h = float(r.High.Y) - float(r.Low.Y)
            cx = (float(r.Low.X) + float(r.High.X)) / 2
            if cx > 1029450 or w < 50:
                continue
            shapes.append(
                (
                    el.ID.Low,
                    w,
                    h,
                    el.Level.Name if el.Level else "?",
                    el.Color,
                    el.LineWeight,
                    float(r.Low.X),
                    float(r.Low.Y),
                )
            )
        except Exception:
            pass
    print(f"nearby large shapes: {len(shapes)}")
    for s in shapes[:15]:
        print(f"  shape id={s[0]} {s[1]:.0f}x{s[2]:.0f} lvl={s[3]} col={s[4]} wt={s[5]}")


if __name__ == "__main__":
    main()
