"""Extend estimate-table column frames to the real last-row bottom.

Padding rebuild grew row bands to y=239540.934 but left tall column shapes
ending at y=239555.734 (cuts through 680.83080011). Unit/Qty lines step.
"""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402
from view_capture import capture_window  # noqa: E402

BOTTOM = 239540.934
HEADER_HI = 239720.642


def _make_band(app, mr, tmpl, x0, x1, y0, y1):
    pts = (
        app.Point3dFromXYZ(x0, y0, 0),
        app.Point3dFromXYZ(x1, y0, 0),
        app.Point3dFromXYZ(x1, y1, 0),
        app.Point3dFromXYZ(x0, y1, 0),
    )
    created = app.CreateShapeElement1(tmpl, pts)
    if isinstance(created, tuple):
        created = created[0]
    mr.AddElement(created)
    created.Rewrite()
    return created


def main() -> None:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = None
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            mr = m
            break
    if mr is None:
        raise SystemExit("XS Stacked Views not found")

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    short_cols = []
    partial = []
    while ee.MoveNext():
        el = ee.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        if "Table Border" not in lvl or int(el.Type) != 6:
            continue
        if float(rng.Low.X) > 1030400 or float(rng.High.X) < 1030150:
            continue
        h = float(rng.High.Y) - float(rng.Low.Y)
        w = float(rng.High.X) - float(rng.Low.X)
        y0 = float(rng.Low.Y)
        y1 = float(rng.High.Y)
        if h < 20:
            continue
        # Partial bottom patch (item+desc only) — replace via full-height cols
        if abs(y0 - BOTTOM) < 0.05 and abs(y1 - 239605.534) < 0.05 and w > 150:
            partial.append(el)
            continue
        if y0 > BOTTOM + 0.1 and y1 > 239700:
            short_cols.append(el)

    print(f"short_cols={len(short_cols)} partial={len(partial)}")
    if not short_cols:
        raise SystemExit("no short column frames found")

    for el in short_cols:
        r = el.Range
        x0, x1 = float(r.Low.X), float(r.High.X)
        y1 = float(r.High.Y)
        _make_band(app, mr, el, x0, x1, BOTTOM, y1)
        mr.RemoveElement(el)
        print(f"extended x=[{x0:.1f},{x1:.1f}] -> y=[{BOTTOM:.3f},{y1:.3f}]")

    for el in partial:
        r = el.Range
        print(
            f"remove partial x=[{float(r.Low.X):.1f},{float(r.High.X):.1f}] "
            f"y=[{float(r.Low.Y):.3f},{float(r.High.Y):.3f}]"
        )
        mr.RemoveElement(el)

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # Verify
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    still_short = []
    while ee2.MoveNext():
        el = ee2.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        if "Table Border" not in lvl or int(el.Type) != 6:
            continue
        h = float(rng.High.Y) - float(rng.Low.Y)
        if h < 20:
            continue
        if float(rng.Low.X) > 1030400 or float(rng.High.X) < 1030150:
            continue
        y0 = float(rng.Low.Y)
        if y0 > BOTTOM + 0.1 and float(rng.High.Y) > 239700:
            still_short.append(
                (
                    el.ID.Low,
                    y0,
                    float(rng.Low.X),
                    float(rng.High.X),
                )
            )
    print(f"still_short={still_short}")

    v = app.ActiveDesignFile.Views(1)
    v.Center = app.Point3dFromXYZ(1030275, 239560, 0)
    v.Extents = app.Point3dFromXYZ(200, 70, 0)
    app.CadInputQueue.SendKeyin("UPDATE VIEW EXTENDED 1")
    time.sleep(2.2)
    p = capture_window(
        "MicroStation",
        out_path=r"c:\repos\microstation-vba-project\scripts\_bts_fmt\table_last3_fixed.png",
    )
    print("capture", p)


if __name__ == "__main__":
    main()
