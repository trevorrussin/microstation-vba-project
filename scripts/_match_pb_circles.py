"""Match each PB* text to enclosing circle(s) in fea_tsp."""
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


def walk(el, texts, circles, parents=None):
    parents = parents or []
    try:
        s = text_of(el)
        if s is not None and re.match(r"^PB\d+$", s.strip(), re.I):
            r = el.Range
            texts.append(
                {
                    "id": el.ID.Low,
                    "text": s.strip(),
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                    "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                    "parent_ids": [p.ID.Low for p in parents],
                    "el": el,
                }
            )
        try:
            if el.IsEllipseElement:
                r = el.Range
                circles.append(
                    {
                        "id": el.ID.Low,
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "w": float(r.High.X) - float(r.Low.X),
                        "h": float(r.High.Y) - float(r.Low.Y),
                        "level": el.Level.Name if el.Level else "?",
                        "color": el.Color,
                        "weight": el.LineWeight,
                        "parent_ids": [p.ID.Low for p in parents],
                        "el": el,
                    }
                )
        except Exception:
            pass
        t = int(el.Type)
        if t in (2, 6, 7, 4):
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk(sub.Current, texts, circles, parents + [el])
            except Exception:
                pass
    except Exception:
        pass


def contains(c, tx, ty, pad=0.0):
    return (
        c["lx"] - pad <= tx <= c["hx"] + pad
        and c["ly"] - pad <= ty <= c["hy"] + pad
    )


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

    print(f"PB texts={len(texts)} ellipses={len(circles)}")
    for t in sorted(texts, key=lambda x: x["text"]):
        tw = t["hx"] - t["lx"]
        th = t["hy"] - t["ly"]
        print(
            f"\n{t['text']} id={t['id']} "
            f"box=[{t['lx']:.2f},{t['ly']:.2f}]-[{t['hx']:.2f},{t['hy']:.2f}] "
            f"w={tw:.2f} h={th:.2f} parents={t['parent_ids']}"
        )
        # ellipses that contain text center, preferably tight
        cands = []
        for c in circles:
            if not contains(c, t["cx"], t["cy"], pad=0.5):
                continue
            # prefer round
            aspect = abs(c["w"] - c["h"]) / max(c["w"], c["h"])
            # size vs text
            margin_x = min(t["lx"] - c["lx"], c["hx"] - t["hx"])
            margin_y = min(t["ly"] - c["ly"], c["hy"] - t["hy"])
            cands.append((c["w"] * c["h"], aspect, margin_x, margin_y, c))
        cands.sort(key=lambda x: x[0])
        print(f"  enclosing ellipses: {len(cands)}")
        for area, aspect, mx, my, c in cands[:8]:
            print(
                f"    id={c['id']} w={c['w']:.2f} h={c['h']:.2f} "
                f"aspect={aspect:.3f} mrg=({mx:.2f},{my:.2f}) "
                f"lvl={c['level']} parents={c['parent_ids']}"
            )


if __name__ == "__main__":
    main()
