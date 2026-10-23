"""Inventory PB# texts and nearby circle/ellipse markers in active fea_tsp."""
from __future__ import annotations

import math
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


def walk(el, texts, circles, arcs):
    try:
        t = int(el.Type)
        if el.IsTextElement or el.IsTextNodeElement:
            s = text_of(el)
            if s and re.search(r"PB\s*#", s, re.I):
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
                        "el": el,
                    }
                )
        try:
            if el.IsEllipseElement:
                r = el.Range
                circles.append(
                    {
                        "id": el.ID.Low,
                        "type": t,
                        "kind": "ellipse",
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "el": el,
                    }
                )
        except Exception:
            pass
        try:
            if el.IsArcElement:
                r = el.Range
                w = float(r.High.X) - float(r.Low.X)
                h = float(r.High.Y) - float(r.Low.Y)
                if abs(w - h) < max(w, h) * 0.35 and max(w, h) < 80:
                    arcs.append(
                        {
                            "id": el.ID.Low,
                            "type": t,
                            "kind": "arc",
                            "lx": float(r.Low.X),
                            "ly": float(r.Low.Y),
                            "hx": float(r.High.X),
                            "hy": float(r.High.Y),
                            "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                            "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                            "el": el,
                        }
                    )
        except Exception:
            pass
        if t in (2, 6, 7):
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk(sub.Current, texts, circles, arcs)
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
    texts, circles, arcs = [], [], []
    n = 0
    while ee.MoveNext():
        n += 1
        walk(ee.Current, texts, circles, arcs)

    print(
        f"scanned roots={n} PB={len(texts)} ellipses={len(circles)} round_arcs={len(arcs)}"
    )
    for t in texts:
        print(
            f"  PB id={t['id']} {t['text']!r} "
            f"cx={t['cx']:.2f} cy={t['cy']:.2f} "
            f"box=[{t['lx']:.2f},{t['ly']:.2f}]-[{t['hx']:.2f},{t['hy']:.2f}]"
        )

    print("--- matches ---")
    for t in texts:
        tw = t["hx"] - t["lx"]
        th = t["hy"] - t["ly"]
        cands = []
        for c in circles + arcs:
            dx = abs(c["cx"] - t["cx"])
            dy = abs(c["cy"] - t["cy"])
            cw = c["hx"] - c["lx"]
            ch = c["hy"] - c["ly"]
            if (
                dx < cw * 0.6
                and dy < ch * 0.6
                and cw < max(tw, th) * 10
                and ch < max(tw, th) * 10
                and max(cw, ch) > min(tw, th) * 0.5
            ):
                cands.append((dx + dy, c, cw, ch))
        cands.sort(key=lambda x: x[0])
        print(f"PB {t['id']} {t['text'][:40]!r} nearby={len(cands)}")
        for dist, c, cw, ch in cands[:5]:
            print(
                f"    {c['kind']} id={c['id']} cw={cw:.2f} ch={ch:.2f} "
                f"dist={dist:.2f} cx={c['cx']:.2f} cy={c['cy']:.2f}"
            )


if __name__ == "__main__":
    main()
