"""Inspect elements near PB1; cancel pending keyin."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch
from view_capture import capture_window


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    app.CadInputQueue.SendKeyin("NULL")
    time.sleep(0.2)

    mr = app.ActiveModelReference
    tx, ty = 1029468.95, 239790.13
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    print("Near PB1 (within 15):")
    while ee.MoveNext():
        el = ee.Current
        try:
            r = el.Range
            cx = (float(r.Low.X) + float(r.High.X)) / 2
            cy = (float(r.Low.Y) + float(r.High.Y)) / 2
            if abs(cx - tx) > 15 or abs(cy - ty) > 15:
                continue
            kind = "?"
            extra = ""
            if el.IsTextElement:
                kind = "text"
                extra = repr(el.AsTextElement.Text)
            elif el.IsEllipseElement:
                kind = "ellipse"
                extra = f"w={float(r.High.X) - float(r.Low.X):.2f}"
            elif el.IsShapeElement:
                kind = "shape"
                extra = (
                    f"w={float(r.High.X) - float(r.Low.X):.2f} "
                    f"h={float(r.High.Y) - float(r.Low.Y):.2f}"
                )
            elif el.IsCellElement:
                kind = "cell"
                extra = el.AsCellElement.Name
            print(
                f"  id={el.ID.Low} type={el.Type} {kind} {extra} "
                f"lvl={el.Level.Name if el.Level else None} col={el.Color} "
                f"cx={cx:.1f} cy={cy:.1f}"
            )
        except Exception:
            pass

    # COM view pan/zoom on PB2
    vw = app.ActiveDesignFile.Views(1)
    cx, cy = 1029512.9, 239820.2
    half_w, half_h = 12.0, 10.0
    try:
        p1 = app.Point3dFromXYZ(cx - half_w, cy - half_h, 0)
        p2 = app.Point3dFromXYZ(cx + half_w, cy + half_h, 0)
        if hasattr(vw, "SetExtents"):
            vw.SetExtents(p1, p2)
        elif hasattr(vw, "Origin") and hasattr(vw, "Delta"):
            vw.Origin = app.Point3dFromXYZ(cx - half_w, cy - half_h, 0)
            vw.Delta = app.Point3dFromXYZ(half_w * 2, half_h * 2, 0)
        vw.Rewrite()
        print("view rewritten")
    except Exception as e:
        print("view set fail", e)

    time.sleep(0.4)
    out = r"c:\repos\microstation-vba-project\scripts\_bts_fmt\pb2_rect_close2.png"
    print(capture_window("OpenRoads", out_path=out))


if __name__ == "__main__":
    main()
