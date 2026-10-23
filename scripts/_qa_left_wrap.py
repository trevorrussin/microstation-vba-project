"""Capture view of left callout wrap rectangle."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from view_capture import capture_window


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    app.CadInputQueue.SendKeyin("NULL")
    time.sleep(0.2)

    # Frame the wrap box with margin
    lx, ly, hx, hy = 1029313.70, 239621.17, 1029434.72, 239775.75
    m = 15.0
    # Use fit view keyin that doesn't wait — MOVE VIEW / WINDOW CENTER
    cx = (lx + hx) / 2
    cy = (ly + hy) / 2
    w = (hx - lx) + 2 * m
    h = (hy - ly) + 2 * m
    vw = app.ActiveDesignFile.Views(1)
    try:
        # CONNECT View: set origin + extents via properties if available
        for name in dir(vw):
            if name.lower() in ("origin", "extents", "delta", "center", "area"):
                print("vw has", name)
        if hasattr(vw, "Origin") and hasattr(vw, "Extents"):
            vw.Origin = app.Point3dFromXYZ(cx - w / 2, cy - h / 2, 0)
            vw.Extents = app.Point3dFromXYZ(w, h, 0)
            print("set Origin+Extents")
        elif hasattr(vw, "SetOrigin") and hasattr(vw, "SetExtents"):
            vw.SetOrigin(app.Point3dFromXYZ(cx - w / 2, cy - h / 2, 0))
            vw.SetExtents(app.Point3dFromXYZ(w, h, 0))
            print("SetOrigin+SetExtents")
    except Exception as e:
        print("view fail", e)

    try:
        app.RedrawAll(True)
    except Exception:
        pass
    time.sleep(0.4)

    out = r"c:\repos\microstation-vba-project\scripts\_bts_fmt\left_callouts_wrap.png"
    print(capture_window("OpenRoads", out_path=out))


if __name__ == "__main__":
    main()
