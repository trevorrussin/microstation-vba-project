"""Find wrap rectangles near left callouts; confirm shape 9693."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    print("file", app.ActiveDesignFile.FullName)
    print("active model", app.ActiveModelReference.Name)
    print("models:", [m.Name for m in app.ActiveDesignFile.Models])

    for m in app.ActiveDesignFile.Models:
        mr = m
        o = Dispatch("MicroStationDGN.ElementScanCriteria")
        o.ExcludeNonGraphical()
        ee = mr.Scan(o)
        hits = []
        while ee.MoveNext():
            el = ee.Current
            try:
                if el.ID.Low == 9693:
                    r = el.Range
                    print(
                        f"FOUND 9693 in {m.Name}: "
                        f"[{float(r.Low.X):.2f},{float(r.Low.Y):.2f}]-"
                        f"[{float(r.High.X):.2f},{float(r.High.Y):.2f}] "
                        f"lvl={el.Level.Name} col={el.Color} wt={el.LineWeight}"
                    )
                if not el.IsShapeElement:
                    continue
                r = el.Range
                lx, ly = float(r.Low.X), float(r.Low.Y)
                hx, hy = float(r.High.Y), float(r.High.Y)
                # typo fix below
            except Exception:
                continue
            try:
                r = el.Range
                lx, ly = float(r.Low.X), float(r.Low.Y)
                hx, hy = float(r.High.X), float(r.High.Y)
                w, h = hx - lx, hy - ly
                # left callout column wraps
                if lx > 1029450 or hx < 1029200:
                    continue
                if w < 80 or h < 40:
                    continue
                if hy < 239600 or ly > 239800:
                    continue
                hits.append(
                    (
                        el.ID.Low,
                        w,
                        h,
                        lx,
                        ly,
                        hx,
                        hy,
                        el.Level.Name if el.Level else "?",
                        el.Color,
                        el.LineWeight,
                    )
                )
            except Exception:
                continue
        if hits:
            print(f"\n=== {m.Name} large left shapes ===")
            for h in sorted(hits, key=lambda x: -x[2]):
                print(
                    f"  id={h[0]} {h[1]:.0f}x{h[2]:.0f} "
                    f"[{h[3]:.1f},{h[4]:.1f}]-[{h[5]:.1f},{h[6]:.1f}] "
                    f"lvl={h[7]} col={h[8]} wt={h[9]}"
                )


if __name__ == "__main__":
    main()
