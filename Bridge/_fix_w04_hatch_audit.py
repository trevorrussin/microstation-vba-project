"""Fix W04 yellow overlay + re-place workspace hatch; audit tips vs sheet."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import Dispatch

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))
import bridge_client as bc
import view_capture

ALIGN_Y = 217040.6
X0 = 1019735.0
YM, YX = 216850.0, 217120.0
XM, XX = 1019700.0, 1022800.0

# 619-311 expected placeable signs (NYW8-33 not in library) at order tips
# from south-align run: advance-warning spacing 350 → signs at 1840/2190/2540 + G20 at DS
EXPECTED = [
    (1840.0, "W20-01RA"),
    (2190.0, "W20-05RA"),
    (2540.0, "W04-02R"),
    (400.0, "G20-02"),  # downstream cumulative after WS end — tip X ~1022725
]


def flush(msg: str) -> None:
    print(msg, flush=True)


def hide_w04_yellow(mr) -> int:
    hidden = 0
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    ee = mr.Scan(oScan)
    while ee.MoveNext():
        el = ee.Current
        if not el.IsCellElement:
            continue
        c = el.AsCellElement
        if not str(c.Name).upper().startswith("W04-02"):
            continue
        oy = float(c.Origin.Y)
        if oy < YM or oy > YX:
            continue
        fr = el.Range
        face_max = max(float(fr.High.X) - float(fr.Low.X), float(fr.High.Y) - float(fr.Low.Y))
        flush(f"W04 {c.Name} id={int(el.ID.Low)} origin=({c.Origin.X:.2f},{c.Origin.Y:.2f})")
        ce = c.GetSubElements()
        i = 0
        while ce.MoveNext():
            sub = ce.Current
            try:
                lvl = sub.Level.Name
                col = int(sub.Color)
            except Exception:
                i += 1
                continue
            if lvl == "SF_P" and col == 4:
                sr = sub.Range
                side = max(float(sr.High.X) - float(sr.Low.X), float(sr.High.Y) - float(sr.Low.Y))
                if side < 0.7 * face_max:
                    try:
                        sub.IsHidden = True
                        sub.Rewrite()
                        hidden += 1
                        flush(f"  hid yellow legend sub[{i}] side={side:.2f}")
                    except Exception as ex:
                        flush(f"  hide fail sub[{i}] {ex}")
            elif lvl == "SFB_P":
                try:
                    sub.DisplayPriority = 1000
                    sub.Rewrite()
                except Exception:
                    pass
            i += 1
        try:
            el.Rewrite()
        except Exception:
            pass
    return hidden


def delete_workspace_shapes(mr) -> list[int]:
    deleted = []
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
        except Exception:
            continue
        cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
        cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
        if not (1022200 <= cx <= 1022450 and 217020 <= cy <= 217055):
            continue
        if "TWZWS" in lvl.upper() or int(el.Color) == 6:
            to_del.append(eid)
    for eid in to_del:
        try:
            r = bc.bridge.call(
                "DELETE_ELEMENT",
                elementId=str(eid),
                ownElementOnly="N",
                reason="replace solid-looking workspace with unfilled hatch",
            )
            flush(f"deleted WS {eid}: {r.get('status')}")
            deleted.append(eid)
        except Exception as ex:
            flush(f"del fail {eid}: {ex}")
    return deleted


def place_workspace() -> dict:
    # 50x12 work space south of align at merging taper / protective vehicle zone
    # Match prior agent placement: X~1022275-1022325, Y just north of align (~217041)
    # South side of eastbound align → Y slightly below align
    y0 = ALIGN_Y - 1.0
    x0 = 1022275.0
    x1 = x0 + 50.0
    y1 = y0 - 12.0
    verts = f"{x0},{y0}|{x1},{y0}|{x1},{y1}|{x0},{y1}"
    r = bc.bridge.call("PLACE_WORKSPACE", verticesTSV=verts, reason="visible diagonal hatch unfilled")
    flush(f"PLACE_WORKSPACE {r}")
    return r


def audit_tips(mr) -> None:
    flush("\n=== south perps (~80ft) + nearest sign cell ===")
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    ee = mr.Scan(oScan)
    perps = []
    signs = []
    while ee.MoveNext():
        el = ee.Current
        try:
            rng = el.Range
        except Exception:
            continue
        cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
        cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
        if not (XM <= cx <= XX and YM <= cy <= YX + 80):
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
                sta = tipx - X0
                perps.append((sta, tipx, tipy, int(el.ID.Low)))
        if el.IsCellElement:
            name = el.AsCellElement.Name
            if name == "TWZSGN_P":
                continue
            up = name.upper()
            if up.startswith(("W", "G", "R", "NY")):
                signs.append((name, float(el.AsCellElement.Origin.X), float(el.AsCellElement.Origin.Y), int(el.ID.Low)))

    perps.sort()
    for sta, tipx, tipy, pid in perps:
        nearest = None
        best = 9999.0
        for name, ox, oy, sid in signs:
            d = ((ox - tipx) ** 2 + (oy - tipy) ** 2) ** 0.5
            if d < best:
                best = d
                nearest = (name, sid, d)
        if nearest and nearest[2] < 120:
            flush(f"  sta≈{sta:7.0f} tip=({tipx:.1f},{tipy:.1f}) SIGN {nearest[0]} d={nearest[2]:.1f}")
        else:
            flush(f"  sta≈{sta:7.0f} tip=({tipx:.1f},{tipy:.1f}) BLANK (non-sign / no face)")

    flush("\n=== expected sheet 619-311 placeable signs ===")
    for sta, code in EXPECTED:
        tipx = 1022725.0 if code == "G20-02" else (X0 + sta)
        hit = None
        for name, ox, oy, sid in signs:
            if abs(ox - tipx) < 100:
                hit = (name, ox, oy, sid)
                break
        flush(f"  expect {code} near X={tipx:.0f}: {'FOUND ' + hit[0] if hit else 'MISSING'}")


def check_hatch(mr) -> None:
    flush("\n=== workspace shapes ===")
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    ee = mr.Scan(oScan)
    while ee.MoveNext():
        el = ee.Current
        if not el.IsShapeElement:
            continue
        rng = el.Range
        cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
        cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
        if not (1022200 <= cx <= 1022450 and 217015 <= cy <= 217055):
            continue
        try:
            hp = el.AsClosedElement.HasPattern
        except Exception:
            hp = "?"
        try:
            fm = el.FillMode
        except Exception:
            fm = "?"
        flush(
            f"  SHAPE id={int(el.ID.Low)} color={el.Color} FillMode={fm} HasPattern={hp} "
            f"X=[{rng.Low.X:.1f},{rng.High.X:.1f}] Y=[{rng.Low.Y:.1f},{rng.High.Y:.1f}]"
        )


def zoom_corridor(app) -> None:
    try:
        v = app.ActiveDesignFile.Views(1)
        # Approximate: set origin/extents if available
        # Fallback: CadInputQueue WINDOW AREA
        ci = app.CadInputQueue
        ci.SendKeyin("WINDOW AREA")
        # datapoints for corridor
        pt1 = app.Point3dFromXYZ(1019800, 216900, 0)
        pt2 = app.Point3dFromXYZ(1022800, 217100, 0)
        ci.SendDataPoint(pt1, 1)
        ci.SendDataPoint(pt2, 1)
        ci.SendReset()
        flush("zoomed WINDOW AREA corridor")
    except Exception as ex:
        flush(f"zoom fail {ex}")


def main() -> int:
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    mr = app.ActiveModelReference
    flush(f"file={app.ActiveDesignFile.FullName}")

    n = hide_w04_yellow(mr)
    flush(f"hidden yellow legend parts: {n}")

    delete_workspace_shapes(mr)
    place_workspace()
    check_hatch(mr)
    audit_tips(mr)

    zoom_corridor(app)
    out = ROOT / "Bridge" / "captures" / "review_w04_hatch_fixed.png"
    try:
        p = view_capture.capture_microstation(out)
        flush(f"capture {p}")
    except Exception as ex:
        flush(f"capture fail {ex}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
