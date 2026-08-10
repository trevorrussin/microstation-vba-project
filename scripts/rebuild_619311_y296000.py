"""Build 619-311 at Y=296000 — G20 black hole, AP tip base, diagonal downstream."""
from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import bridge_client
import view_capture
import wztc_ops

Y = 296000.0
UP = [23760.0, Y, 0.0]
DN = [23860.0, Y, 0.0]


def capture(name: str, x: float, y: float, w: float, h: float) -> None:
    view_capture.navigate_view(x, y, w, h, view_num=1)
    time.sleep(0.45)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / f"{name}.png"
    shutil.copy2(src, dest)
    print(f"CAPTURE {dest}", flush=True)


def main() -> int:
    import pythoncom
    import win32com.client
    from win32com.client import Dispatch

    wztc_ops.set_bridge(bridge_client.chat_bridge)
    wztc_ops.reset_plan_session_flags()

    print(f"BUILD 619-311 at Y={Y}", flush=True)
    wztc_ops.build_wztc_order_table(
        speed=55, road_type="Non-Freeway", lane_width=12,
        shoulder_width=">= 8 ft", sheet_num="619-311", area_type="RURAL",
    )
    t0 = time.time()
    result = wztc_ops.run_sheet_build(
        upstream_edge=UP, downstream_edge=DN, outward_sign=-1.0,
        include_visual_qa=True, clear_prior_stations=True,
    )
    print(f"status={result.get('status')} in {time.time()-t0:.1f}s", flush=True)
    sc = wztc_ops.get_geometry_scorecard(sheet_num="619-311")
    print("scorecard", sc.get("passed"), sc.get("failures"), flush=True)

    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    mr = app.ActiveModelReference

    # G20 SF_P hole color
    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    ee = mr.Scan(o_scan)
    g20_hole = None
    ap_ys = []
    sign_ys = []
    dn_cones = []
    while ee.MoveNext():
        el = ee.Current
        try:
            r = el.Range
            cx = 0.5 * (float(r.Low.X) + float(r.High.X))
            cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
        except Exception:
            continue
        if abs(cy - Y) > 200:
            continue
        # G20 face
        try:
            if (el.IsCellElement or el.IsSharedCellElement) and abs(cx - 23990) < 30 and abs(cy - (Y - 113)) < 40:
                ee2 = el.AsCellElement.GetSubElements() if el.IsCellElement else el.GetSubElements()
                while ee2.MoveNext():
                    sub = ee2.Current
                    try:
                        lvl = sub.Level.Name if sub.Level else ""
                    except Exception:
                        lvl = ""
                    if lvl.upper() == "SF_P" and int(sub.Type) == 2:
                        g20_hole = int(sub.Color)
        except Exception:
            pass
        # AP panel cell ~47 wide
        try:
            w = float(r.High.X) - float(r.Low.X)
            h = float(r.High.Y) - float(r.Low.Y)
            if abs(cx - 22465) < 20 and 40 < w < 55 and 15 < h < 30:
                ap_ys.append(round(cy, 1))
        except Exception:
            pass
        # diamond faces ~ further out on approach
        try:
            if el.IsCellElement and 20900 < cx < 22100:
                w = float(r.High.X) - float(r.Low.X)
                if 40 < w < 90:
                    sign_ys.append(round(cy, 1))
        except Exception:
            pass
        try:
            lvl = el.Level.Name if el.Level else ""
        except Exception:
            lvl = ""
        if "TWZCD" in lvl.upper() and cx >= 23850:
            dn_cones.append({"x": round(cx, 1), "y": round(cy, 1)})

    dn_cones = sorted(dn_cones, key=lambda c: c["x"])
    print("G20 SF_P hole color", g20_hole, "(want 240)", flush=True)
    print("AP face Ys", ap_ys, "sample sign Ys", sorted(set(sign_ys))[:6], flush=True)
    print("downstream cones", dn_cones, flush=True)

    mid = 0.5 * (UP[0] + DN[0])
    capture("cursor_296000_overview", mid, Y, 3200, 600)
    capture("cursor_296000_g20", 23990, Y - 90, 320, 300)
    capture("cursor_296000_ap", 22465, Y + 60, 300, 280)
    capture("cursor_296000_downstream", 23885, Y, 450, 280)
    capture("cursor_296000_ap_vs_signs", 21700, Y + 50, 1800, 400)

    ys = [c["y"] for c in dn_cones]
    diagonal = len(ys) >= 2 and abs(ys[0] - ys[-1]) > 5
    ok = (
        result.get("status") == "OK"
        and sc.get("passed")
        and g20_hole == 240
        and diagonal
    )
    summary = {
        "y": Y, "status": result.get("status"), "scorecard": sc.get("passed"),
        "g20HoleColor": g20_hole, "apYs": ap_ys, "dnCones": dn_cones, "ok": ok,
    }
    (ROOT / "Bridge" / "captures" / "inspect_296000.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8")
    print("SUMMARY", json.dumps(summary, indent=2), flush=True)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
