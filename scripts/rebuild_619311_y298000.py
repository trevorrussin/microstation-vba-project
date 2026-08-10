"""Build 619-311 at Y=298000 (1000 ft below engineer-corrected Y=299000).

Verifies: unfilled cones (FillMode 0), G20-02 code+size label,
arrow panel with TWZSGN_P post + stem.
"""
from __future__ import annotations

import json
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import bridge_client
import placement_registry
import view_capture
import wztc_ops

# 1000 ft below the prior Y=298000 attempt (and 2000 below engineer Y=299000).
Y = 297000.0
UP = [23760.0, Y, 0.0]
DN = [23860.0, Y, 0.0]


def capture(name: str, x: float, y: float, w: float, h: float) -> Path:
    view_capture.navigate_view(x, y, w, h, view_num=1)
    time.sleep(0.45)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / f"{name}.png"
    shutil.copy2(src, dest)
    print(f"CAPTURE {dest} ({dest.stat().st_size} bytes)", flush=True)
    return dest


def main() -> int:
    import pythoncom
    import win32com.client
    from win32com.client import Dispatch

    wztc_ops.set_bridge(bridge_client.chat_bridge)
    wztc_ops.reset_plan_session_flags()

    print(f"BUILD 619-311 at Y={Y}", flush=True)
    table = wztc_ops.build_wztc_order_table(
        speed=55,
        road_type="Non-Freeway",
        lane_width=12,
        shoulder_width=">= 8 ft",
        sheet_num="619-311",
        area_type="RURAL",
    )
    print("order table:", table.get("status"), flush=True)

    t0 = time.time()
    result = wztc_ops.run_sheet_build(
        upstream_edge=UP,
        downstream_edge=DN,
        outward_sign=-1.0,
        include_visual_qa=True,
        clear_prior_stations=True,
        force=False,
    )
    print(f"run_sheet_build in {time.time()-t0:.1f}s status={result.get('status')}",
          flush=True)

    sc = wztc_ops.get_geometry_scorecard(sheet_num="619-311")
    print("SCORECARD passed=", sc.get("passed"), "failures=", sc.get("failures"), flush=True)

    placements = placement_registry.resolve_latest_placements(sheet_num="619-311")
    g20 = next((p for p in placements if "G20" in str(p.get("primitiveId"))), None)
    ap = next((p for p in placements if p.get("kind") == "arrowPanel"), None)
    print("G20", {k: g20.get(k) for k in ("primitiveId", "elementIds", "x", "y")} if g20 else None)
    print("AP", {k: ap.get(k) for k in ("primitiveId", "elementIds", "bridgeOp", "x", "y")} if ap else None)

    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    mr = app.ActiveModelReference
    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    ee = mr.Scan(o_scan)

    cone_fills = {}
    g20_texts = []
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
        try:
            lvl = el.Level.Name if el.Level else ""
        except Exception:
            lvl = ""
        if "TWZCD" in lvl.upper():
            fill = -1
            try:
                fill = int(el.AsShapeElement.FillMode)
            except Exception:
                try:
                    fill = int(el.FillMode)
                except Exception:
                    pass
            cone_fills[fill] = cone_fills.get(fill, 0) + 1
        try:
            if el.IsTextElement or el.IsTextNodeElement:
                if abs(cx - 23990) > 120:
                    continue
                if el.IsTextElement:
                    txt = el.AsTextElement.Text
                else:
                    tn = el.AsTextNodeElement
                    txt = " | ".join(
                        str(tn.TextLine(i))
                        for i in range(1, int(tn.TextLinesCount) + 1)
                    )
                g20_texts.append({"cx": round(cx, 2), "cy": round(cy, 2), "text": txt, "lvl": lvl})
        except Exception:
            pass

    print("cone FillMode counts:", cone_fills, flush=True)
    print("G20-area texts:", json.dumps(g20_texts, indent=2), flush=True)

    mid = 0.5 * (UP[0] + DN[0])
    capture("cursor_297000_overview", mid, Y, 3200, 600)
    capture("cursor_297000_ap", 22465, Y + 40, 280, 220)
    capture("cursor_297000_g20", 23990, Y - 90, 280, 280)
    capture("cursor_297000_taper", 22800, Y + 8, 900, 350)

    ok = (
        result.get("status") == "OK"
        and sc.get("passed")
        and (cone_fills.get(0, 0) >= 30 or cone_fills.get(-1, 0) >= 30)
        and cone_fills.get(1, 0) == 0
        and any("G20" in (t.get("text") or "").upper() for t in g20_texts)
    )
    summary = {
        "y": Y,
        "status": result.get("status"),
        "scorecardPassed": sc.get("passed"),
        "coneFills": cone_fills,
        "g20Texts": g20_texts,
        "ap": ap.get("elementIds") if ap else None,
        "ok": ok,
    }
    (ROOT / "Bridge" / "captures" / "inspect_297000.json").write_text(
        json.dumps(summary, indent=2, default=str), encoding="utf-8"
    )
    print("SUMMARY", json.dumps(summary, indent=2, default=str), flush=True)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
