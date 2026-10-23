"""Verify PB# labels now have rectangles, not circles."""
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
            return el.AsTextElement.Text.strip()
    except Exception:
        pass
    return None


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    mr = app.ActiveModelReference
    print("file", app.ActiveDesignFile.FullName)
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts, ellipses, shapes = [], [], []
    while ee.MoveNext():
        el = ee.Current
        try:
            t = text_of(el)
            if t and re.match(r"^PB\d+$", t, re.I):
                r = el.Range
                texts.append(
                    {
                        "id": el.ID.Low,
                        "text": t,
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                    }
                )
            if el.IsEllipseElement:
                r = el.Range
                ellipses.append(
                    {
                        "id": el.ID.Low,
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "w": float(r.High.X) - float(r.Low.X),
                        "h": float(r.High.Y) - float(r.Low.Y),
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                    }
                )
            if el.IsShapeElement:
                r = el.Range
                w = float(r.High.X) - float(r.Low.X)
                h = float(r.High.Y) - float(r.Low.Y)
                if max(w, h) < 20:
                    shapes.append(
                        {
                            "id": el.ID.Low,
                            "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                            "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                            "w": w,
                            "h": h,
                            "lx": float(r.Low.X),
                            "ly": float(r.Low.Y),
                            "hx": float(r.High.X),
                            "hy": float(r.High.Y),
                        }
                    )
        except Exception:
            # also walk complexes dropped? texts should be top-level now
            pass

    print(f"PB texts={len(texts)}")
    for t in sorted(texts, key=lambda x: x["text"]):
        enc_ell = [
            e
            for e in ellipses
            if e["lx"] <= t["cx"] <= e["hx"] and e["ly"] <= t["cy"] <= e["hy"]
            and e["w"] < 15
        ]
        enc_shp = [
            s
            for s in shapes
            if s["lx"] - 0.05 <= t["lx"]
            and s["ly"] - 0.05 <= t["ly"]
            and s["hx"] + 0.05 >= t["hx"]
            and s["hy"] + 0.05 >= t["hy"]
            and s["w"] < 15
        ]
        print(
            f"  {t['text']} id={t['id']}: "
            f"enclosing_circles={len(enc_ell)} enclosing_rects={len(enc_shp)}"
        )
        for s in enc_shp:
            print(f"    rect id={s['id']} {s['w']:.2f}x{s['h']:.2f}")
        for e in enc_ell:
            print(f"    CIRCLE STILL id={e['id']} {e['w']:.2f}x{e['h']:.2f}")


if __name__ == "__main__":
    main()
