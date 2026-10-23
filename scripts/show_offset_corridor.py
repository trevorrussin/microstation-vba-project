"""Zoom to offset QA corridor and make it visible in View 1."""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))

import pythoncom
import win32com.client
from view_capture import navigate_view
from win32com.client import Dispatch

CORRIDOR_ELEMENT_ID = 2722043
ZOOM_X = 1_006_333.0
ZOOM_Y = 247_128.0
ZOOM_WIDTH_FT = 400.0


def main() -> None:
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("MicroStationDGN.Application")
        mr = app.ActiveDesignFile.DefaultModelReference
        print(f"file={app.ActiveDesignFile.FullName}")

        o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
        o_scan.ExcludeNonGraphical()
        ee = mr.Scan(o_scan)
        found = False
        while ee.MoveNext():
            el = ee.Current
            if int(el.ID.Low) != CORRIDOR_ELEMENT_ID:
                continue
            found = True
            el.Color = 2
            el.LineWeight = 1
            closed = el.AsClosedElement
            if closed.HasPattern():
                try:
                    pat = closed.GetPattern()
                    pat.Color = 3
                except Exception:
                    pass
            el.Rewrite()
            rng = el.Range
            print(
                f"shape id={CORRIDOR_ELEMENT_ID} level={el.Level.Name} "
                f"range X=[{float(rng.Low.X):.1f},{float(rng.High.X):.1f}] "
                f"Y=[{float(rng.Low.Y):.1f},{float(rng.High.Y):.1f}]"
            )
            break
        if not found:
            print(f"ERROR: shape {CORRIDOR_ELEMENT_ID} not found — run compare script with DRAW_CORRIDOR_HATCH=True")
            return

        try:
            app.CadInputQueue.SendKeyin("VIEW CONTROL PATTERNS ON")
            print("sent: VIEW CONTROL PATTERNS ON")
        except Exception as exc:
            print(f"patterns keyin skipped: {exc}")

        state = navigate_view(ZOOM_X, ZOOM_Y, ZOOM_WIDTH_FT, ZOOM_WIDTH_FT / 2.5, z=0.0)
        print(
            f"view1 center=({state['centerX']:.1f},{state['centerY']:.1f}) "
            f"width={state['width']:.1f} ft"
        )
        print(f"Pan to station 19+50: X={ZOOM_X:.0f} Y={ZOOM_Y:.0f}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
