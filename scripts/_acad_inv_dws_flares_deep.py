"""Deep find DWS and flare entities — may be hatches/blocks not HEX solids."""
from __future__ import annotations

import pythoncom
from win32com.client import Dispatch


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    print("=== PROP-D-ROAD-RAMPS ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != "PROP-D-ROAD-RAMPS":
                continue
            mn, mx = e.GetBoundingBox()
            extra = ""
            if e.ObjectName == "AcDbHatch":
                extra = f" pat={e.PatternName}"
            if e.ObjectName == "AcDbBlockReference":
                extra = f" block={e.Name}"
            print(
                f"  {e.Handle} {e.ObjectName} col={e.Color}{extra} "
                f"[{float(mn[0]):.1f},{float(mn[1]):.1f}]-"
                f"[{float(mx[0]):.1f},{float(mx[1]):.1f}]"
            )
        except Exception:
            continue

    # All hatches near location 01 (X ~1028450-1028600, Y ~187400-187520)
    print("\n=== HATCHES in loc01 bbox ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbHatch":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187400 < cy < 187530):
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            if max(w, h) > 40:
                continue
            print(
                f"  {e.Handle} lyr={e.Layer!r} pat={e.PatternName} "
                f"col={e.Color} area={getattr(e,'Area',None)} "
                f"cx={cx:.1f} cy={cy:.1f} {w:.1f}x{h:.1f}"
            )
        except Exception:
            continue

    # Yellow-ish entities (color 2 or truecolor yellow) small
    print("\n=== color-2 entities in loc01 ===")
    n = 0
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            col = e.Color
            tc = None
            try:
                tc = (e.TrueColor.Red, e.TrueColor.Green, e.TrueColor.Blue)
            except Exception:
                pass
            is_yellow = col == 2 or (tc and tc[0] > 200 and tc[1] > 200 and tc[2] < 80)
            if not is_yellow and col != 256:
                continue
            # also check bylayer yellow via layer color
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187400 < cy < 187530):
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            if max(w, h) > 15 or max(w, h) < 0.3:
                continue
            if col == 256:
                try:
                    lc = doc.Layers.Item(e.Layer).Color
                except Exception:
                    lc = -1
                if lc != 2:
                    continue
            print(
                f"  {e.Handle} {e.ObjectName} lyr={e.Layer!r} col={col} tc={tc} "
                f"cx={cx:.1f} cy={cy:.1f} {w:.1f}x{h:.1f}"
            )
            n += 1
            if n > 60:
                print("  ... truncated")
                break
        except Exception:
            continue

    # blocks in loc01
    print("\n=== BLOCKS in loc01 (small) ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbBlockReference":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187400 < cy < 187530):
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            if max(w, h) > 30:
                continue
            print(
                f"  {e.Handle} block={e.Name!r} lyr={e.Layer!r} "
                f"cx={cx:.1f} cy={cy:.1f} {w:.1f}x{h:.1f}"
            )
        except Exception:
            continue


if __name__ == "__main__":
    main()
