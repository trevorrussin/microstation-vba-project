"""Pure-COM: recreate workspace hatch + audit tips (no bridge). W04 already fixed."""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import Dispatch

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))
import view_capture

ALIGN_Y = 217040.6
X0 = 1019735.0
YM, YX = 216850.0, 217120.0
XM, XX = 1019700.0, 1022800.0

EXPECTED = [
    (1840.0, "W20-01RA"),
    (2190.0, "W20-05RA"),
    (2540.0, "W04-02R"),
]


def flush(msg: str) -> None:
    print(msg, flush=True)


def main() -> int:
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    mr = app.ActiveModelReference
    flush(f"file={app.ActiveDesignFile.FullName}")

    # --- delete existing TWZWS / orange shapes in WS zone ---
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    ee = mr.Scan(oScan)
    to_del = []
    while ee.MoveNext():
        el = ee.Current
        if not el.IsShapeElement:
            continue
        try:
            rng = el.Range
            eid = int(el.ID.Low)
            lvl = el.Level.Name
            col = int(el.Color)
        except Exception:
            continue
        cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
        cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
        if not (1022200 <= cx <= 1022450 and 217015 <= cy <= 217055):
            continue
        if "TWZWS" in lvl.upper() or col == 6:
            to_del.append(el)
            flush(f"queue delete shape {eid} lvl={lvl} color={col}")
    for el in to_del:
        try:
            mr.RemoveElement(el)
            flush(f"removed {int(el.ID.Low)}")
        except Exception as ex:
            flush(f"remove fail: {ex}")

    # --- place unfilled shape + hatch via COM ---
    y0 = ALIGN_Y - 1.0
    x0 = 1022275.0
    x1 = x0 + 50.0
    y1 = y0 - 12.0
    coords = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    pts = [app.Point3dFromXYZ(x, y, 0) for x, y in coords]

    # msdFillModeNone = 0 (Bentley)
    fill_none = 0
    try:
        fill_none = int(win32com.client.constants.msdFillModeNone)
        flush(f"msdFillModeNone={fill_none}")
    except Exception:
        flush("using fill_none=0")

    raw = app.CreateShapeElement1(None, pts, fill_none)
    shape = raw[0] if isinstance(raw, tuple) else raw
    shape.Color = 6
    shape.LineWeight = 2
    try:
        shape.Level = app.ActiveDesignFile.Levels("TWZWS2_P")
    except Exception as ex:
        flush(f"level set fail {ex}")
    mr.AddElement(shape)
    shape.Rewrite()
    flush(f"placed shape id={int(shape.ID.Low)} FillMode={getattr(shape, 'FillMode', '?')}")

    spacing = 2.0
    ang = 45.0 * math.pi / 180.0
    hatch_raw = app.CreateHatchPattern1(spacing, ang)
    hatch = hatch_raw[0] if isinstance(hatch_raw, tuple) else hatch_raw
    hatch.Color = 6
    hatch.LineWeight = 2
    closed = shape.AsClosedElement
    try:
        mat = app.Matrix3dIdentity()
        if isinstance(mat, tuple):
            mat = mat[0]
        closed.SetPattern(hatch, mat)
    except Exception as ex:
        flush(f"SetPattern with matrix fail: {ex}; retry 1-arg")
        closed.SetPattern(hatch)
    shape.Rewrite()
    try:
        hp = closed.HasPattern
    except Exception:
        hp = "?"
    flush(f"hatch HasPattern={hp} spacing={spacing}")

    # --- audit W04 hidden state ---
    oScan2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan2.ExcludeNonGraphical()
    ee2 = mr.Scan(oScan2)
    while ee2.MoveNext():
        el = ee2.Current
        if not el.IsCellElement:
            continue
        c = el.AsCellElement
        if c.Name != "W04-02R":
            continue
        if abs(float(c.Origin.Y) - 216917.4) > 2:
            continue
        flush(f"W04 id={int(el.ID.Low)} still present")
        ce = c.GetSubElements()
        i = 0
        hid_n = 0
        while ce.MoveNext():
            sub = ce.Current
            try:
                if sub.IsHidden:
                    hid_n += 1
                    flush(f"  sub[{i}] IsHidden=True color={sub.Color} lvl={sub.Level.Name}")
            except Exception:
                pass
            i += 1
        flush(f"  hidden count={hid_n}")

    # --- tip audit ---
    flush("\n=== south perps + signs ===")
    oScan3 = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan3.ExcludeNonGraphical()
    ee3 = mr.Scan(oScan3)
    perps = []
    signs = []
    while ee3.MoveNext():
        el = ee3.Current
        try:
            rng = el.Range
        except Exception:
            continue
        cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
        cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
        if not (XM <= cx <= XX and YM <= cy <= YX + 100):
            continue
        if el.IsLineElement:
            try:
                s = el.AsLineElement.StartPoint
                e = el.AsLineElement.EndPoint
            except Exception:
                continue
            length = ((float(e.X) - float(s.X)) ** 2 + (float(e.Y) - float(s.Y)) ** 2) ** 0.5
            if 75 <= length <= 85:
                tipx = float(s.X) if float(s.Y) < float(e.Y) else float(e.X)
                tipy = min(float(s.Y), float(e.Y))
                perps.append((tipx - X0, tipx, tipy, int(el.ID.Low)))
        if el.IsCellElement:
            name = el.AsCellElement.Name
            if name == "TWZSGN_P":
                continue
            up = name.upper()
            if up.startswith(("W", "G", "R", "NY")):
                signs.append((name, float(el.AsCellElement.Origin.X), float(el.AsCellElement.Origin.Y)))

    perps.sort()
    for sta, tipx, tipy, pid in perps:
        nearest = None
        best = 9999.0
        for name, ox, oy in signs:
            d = ((ox - tipx) ** 2 + (oy - tipy) ** 2) ** 0.5
            if d < best:
                best = d
                nearest = (name, d)
        if nearest and nearest[1] < 120:
            flush(f"  sta≈{sta:7.0f} tip=({tipx:.1f},{tipy:.1f}) SIGN {nearest[0]} d={nearest[1]:.1f}")
        else:
            flush(f"  sta≈{sta:7.0f} tip=({tipx:.1f},{tipy:.1f}) BLANK")

    flush("\n=== 619-311 placeable signs (excl NYW8-33) ===")
    flush(f"  sheet signs: G20-2 | NYW8-33(gap) | W20-1 | W20-5R | W4-2R")
    flush(f"  faces found: {sorted({n for n, _, _ in signs})}")
    for sta, code in EXPECTED:
        tipx = X0 + sta
        hit = [n for n, ox, oy in signs if abs(ox - tipx) < 100]
        flush(f"  expect {code} @ X≈{tipx:.0f}: {hit[0] if hit else 'MISSING'}")
    g20 = [n for n, ox, oy in signs if n.upper().startswith("G20")]
    flush(f"  expect G20-02 downstream: {g20[0] if g20 else 'MISSING'} {[round(ox) for n,ox,oy in signs if n.upper().startswith('G20')]}")

    # zoom + capture
    try:
        ci = app.CadInputQueue
        ci.SendKeyin("WINDOW AREA")
        ci.SendDataPoint(app.Point3dFromXYZ(1019800, 216880, 0), 1)
        ci.SendDataPoint(app.Point3dFromXYZ(1022800, 217080, 0), 1)
        ci.SendReset()
        flush("zoomed")
    except Exception as ex:
        flush(f"zoom {ex}")

    out = ROOT / "Bridge" / "captures" / "review_w04_hatch_fixed.png"
    p = view_capture.capture_microstation(out)
    flush(f"capture {p}")

    # close-up of W04 + workspace
    try:
        ci = app.CadInputQueue
        ci.SendKeyin("WINDOW AREA")
        ci.SendDataPoint(app.Point3dFromXYZ(1022200, 216880, 0), 1)
        ci.SendDataPoint(app.Point3dFromXYZ(1022450, 217060, 0), 1)
        ci.SendReset()
        out2 = ROOT / "Bridge" / "captures" / "review_w04_ws_closeup.png"
        p2 = view_capture.capture_microstation(out2)
        flush(f"capture {p2}")
    except Exception as ex:
        flush(f"closeup {ex}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
