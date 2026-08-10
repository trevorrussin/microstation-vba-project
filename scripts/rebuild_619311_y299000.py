"""Build 619-311 at Y=299000 (1000 ft below Claude's Y=300000 corridor).

Does NOT clear the prior Y=300k/299.9k bands — new corridor only.
Inspects DGN via COM (not screenshots) for: cone junction duplicates,
G20 size label, arrow-panel stem+cell.
"""
from __future__ import annotations

import json
import math
import shutil
import sys
import time
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import bridge_client
import placement_registry
import view_capture
import wztc_ops

# 1000 ft below the original Y=300000 band (also below Y=299900 rebuild).
Y = 299000.0
WORK_LEN = 100.0  # same bay length Claude used (23760..23860)
# Match Claude's X band so comparison is apples-to-apples.
UP = [23760.0, Y, 0.0]
DN = [23860.0, Y, 0.0]


def capture(name: str, x: float, y: float, w: float, h: float) -> Path:
    view_capture.navigate_view(x, y, w, h, view_num=1)
    time.sleep(0.4)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / f"{name}.png"
    shutil.copy2(src, dest)
    print(f"CAPTURE {dest} ({dest.stat().st_size} bytes)", flush=True)
    return dest


def inspect_band(y: float, half_h: float = 80.0) -> dict:
    """COM spatial inspection around the new corridor — not PrintWindow."""
    cx = 0.5 * (UP[0] + DN[0])
    # Wide enough to cover approach signs + downstream
    near = wztc_ops.find_elements_near(
        cx, y, radius=3500.0, type_filter="", force=True)
    rows = near if isinstance(near, list) else (near.get("rows") or near.get("elements") or [])
    in_band = []
    for r in rows:
        try:
            ry = float(r.get("y") or r.get("centerY") or r.get("cy") or 0)
        except (TypeError, ValueError):
            continue
        if abs(ry - y) <= half_h + 200:  # loose for signs along approach
            in_band.append(r)

    # Cone / marker clustering by rounded XY
    pts: dict[tuple[float, float], list] = defaultdict(list)
    texts = []
    cells = []
    for r in in_band:
        t = str(r.get("type") or r.get("elementType") or "").upper()
        try:
            x = float(r.get("x") or r.get("centerX") or r.get("cx") or 0)
            yy = float(r.get("y") or r.get("centerY") or r.get("cy") or 0)
        except (TypeError, ValueError):
            continue
        key = (round(x, 2), round(yy, 2))
        level = str(r.get("level") or "")
        cell = str(r.get("cellName") or r.get("cell") or "")
        if "TWZCD" in level.upper() or t in ("SHAPE", "COMPLEX_SHAPE", "CELL"):
            if abs(yy - y) <= half_h:
                pts[key].append(r)
        if t in ("TEXT", "TEXT_NODE") or "TEXT" in t:
            texts.append(r)
        if cell:
            cells.append(r)

    dups = {k: v for k, v in pts.items() if len(v) > 1}
    g20_labels = [
        r for r in texts
        if any(s in str(r.get("text") or r.get("content") or "").upper()
               for s in ('36"', "36 X", "END ROAD", "G20"))
    ]
    ap_cells = [r for r in cells if "TWZAP" in str(r.get("cellName") or r.get("cell") or "").upper()
                or "ARROW" in str(r.get("cellName") or "").upper()]

    placements = placement_registry.resolve_latest_placements(sheet_num="619-311")
    ap_regs = [p for p in placements if p.get("kind") == "arrowPanel"]
    cone_regs = [p for p in placements if p.get("kind") == "cone"]

    return {
        "inBandApprox": len(in_band),
        "duplicatePointKeys": len(dups),
        "duplicateSamples": [
            {"xy": k, "count": len(v), "ids": [x.get("elementId") for x in v[:4]]}
            for k, v in list(dups.items())[:8]
        ],
        "g20RelatedTexts": [
            {"id": r.get("elementId"), "text": r.get("text") or r.get("content"),
             "x": r.get("x"), "y": r.get("y")}
            for r in g20_labels[:10]
        ],
        "arrowPanelCellsNear": len(ap_cells),
        "registryArrowPanels": [
            {"primitiveId": p.get("primitiveId"), "elementIds": p.get("elementIds"),
             "bridgeOp": p.get("bridgeOp"), "x": p.get("x"), "y": p.get("y")}
            for p in ap_regs[-3:]
        ],
        "registryConeRuns": len(cone_regs),
        "registryConeExtras": [
            {"pid": p.get("primitiveId"), "coneCount": p.get("coneCount"),
             "run": p.get("run")}
            for p in cone_regs[-6:]
        ],
    }


def main() -> int:
    wztc_ops.set_bridge(bridge_client.chat_bridge)
    wztc_ops.reset_plan_session_flags()

    print(f"BUILD 619-311 at Y={Y} (1000 ft below Y=300000)", flush=True)
    print(f"EDGES up={UP} dn={DN}", flush=True)
    # Do NOT clear prior bands — only place a new corridor here.

    table = wztc_ops.build_wztc_order_table(
        speed=55,
        road_type="Non-Freeway",
        lane_width=12,
        shoulder_width=">= 8 ft",
        sheet_num="619-311",
        area_type="RURAL",
    )
    print("order table:", table.get("status"),
          "specDriven=", table.get("specDriven"), flush=True)

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
    print(json.dumps({
        "status": result.get("status"),
        "failedPhase": result.get("failedPhase"),
        "phases": [
            {k: p.get(k) for k in ("phase", "skipped", "error", "status")
             if k in p or True}
            for p in (result.get("phases") or [])
        ],
        "scorecard": (result.get("geometry") or {}).get("scorecard")
            if isinstance(result.get("geometry"), dict) else None,
        "visualQaPassed": result.get("visualQaPassed"),
        "errors": result.get("errors") or (result.get("geometry") or {}).get("errors"),
    }, default=str, indent=2)[:4000], flush=True)

    sc = wztc_ops.get_geometry_scorecard(sheet_num="619-311")
    print("SCORECARD", json.dumps({
        "passed": sc.get("passed"),
        "failures": sc.get("failures"),
        "placed": sc.get("placed"),
    }, default=str, indent=2)[:2000], flush=True)

    insp = inspect_band(Y)
    print("INSPECT", json.dumps(insp, default=str, indent=2)[:3000], flush=True)
    (ROOT / "Bridge" / "captures" / "inspect_299000.json").write_text(
        json.dumps({"y": Y, "edges": {"up": UP, "dn": DN}, "scorecard": sc,
                    "inspect": insp, "runStatus": result.get("status")},
                   default=str, indent=2),
        encoding="utf-8",
    )

    mid = 0.5 * (UP[0] + DN[0])
    capture("cursor_299000_overview", mid, Y, 3200, 600)
    capture("cursor_299000_work", mid, Y, 900, 400)
    capture("cursor_299000_taper", UP[0] + 400, Y, 1100, 450)
    capture("cursor_299000_downstream", DN[0] - 100, Y, 1000, 400)

    ok = result.get("status") == "OK" and sc.get("passed") and insp["duplicatePointKeys"] == 0
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
