"""Broader text dump near small ellipses in active fea_tsp."""
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


def walk(el, texts, circles, depth=0):
    try:
        t = int(el.Type)
        s = text_of(el)
        if s is not None and s.strip():
            r = el.Range
            texts.append(
                {
                    "id": el.ID.Low,
                    "text": s,
                    "type": t,
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                    "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                    "depth": depth,
                }
            )
        try:
            if el.IsEllipseElement:
                r = el.Range
                w = float(r.High.X) - float(r.Low.X)
                h = float(r.High.Y) - float(r.Low.Y)
                circles.append(
                    {
                        "id": el.ID.Low,
                        "w": w,
                        "h": h,
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "level": el.Level.Name if el.Level else "?",
                        "color": el.Color,
                        "weight": el.LineWeight,
                    }
                )
        except Exception:
            pass
        if t in (2, 6, 7, 4) or depth < 6:
            try:
                if hasattr(el, "AsComplexElement"):
                    sub = el.AsComplexElement.GetSubElements()
                    while sub.MoveNext():
                        walk(sub.Current, texts, circles, depth + 1)
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
    texts, circles = [], []
    while ee.MoveNext():
        walk(ee.Current, texts, circles)

    print(f"texts={len(texts)} ellipses={len(circles)}")

    # PB-ish
    pbish = [t for t in texts if re.search(r"PB", t["text"], re.I)]
    print(f"texts with PB: {len(pbish)}")
    for t in pbish[:50]:
        print(f"  id={t['id']} type={t['type']} {t['text']!r}")

    # hash-ish short labels
    hashish = [
        t
        for t in texts
        if "#" in t["text"] or re.search(r"\bP\s*B\b", t["text"], re.I)
    ]
    print(f"texts with # or P B: {len(hashish)}")
    for t in hashish[:80]:
        print(f"  id={t['id']} {t['text']!r} @({t['cx']:.1f},{t['cy']:.1f})")

    # small nearly-circular ellipses
    small = [
        c
        for c in circles
        if max(c["w"], c["h"]) < 30
        and min(c["w"], c["h"]) > 0.5
        and abs(c["w"] - c["h"]) / max(c["w"], c["h"]) < 0.25
    ]
    print(f"small round ellipses: {len(small)}")
    for c in sorted(small, key=lambda x: x["w"])[:40]:
        # nearest text
        nearest = None
        nd = 1e9
        for t in texts:
            d = abs(t["cx"] - c["cx"]) + abs(t["cy"] - c["cy"])
            if d < nd:
                nd = d
                nearest = t
        nt = nearest["text"][:40] if nearest else "?"
        print(
            f"  ell id={c['id']} w={c['w']:.2f} h={c['h']:.2f} "
            f"lvl={c['level']} col={c['color']} wt={c['weight']} "
            f"near_d={nd:.2f} txt={nt!r}"
        )


if __name__ == "__main__":
    main()
