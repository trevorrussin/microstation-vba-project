"""Clean horizontal 619-311 rebuild (dy/dx = 0) in a clear model area."""
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

# Clean band well north of prior junk (~Y 217k / 218.5k)
Y = 230000.0
WORK_LEN = 150.0
# Westbound through bay (matches prior travel sense): upstream = east edge
UP = [1020150.0, Y, 0.0]
DN = [1020150.0 - WORK_LEN, Y, 0.0]  # 1020000, 230000


def capture(name: str, x: float, y: float, w: float, h: float) -> Path:
    view_capture.navigate_view(x, y, w, h, view_num=1)
    time.sleep(0.3)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / f"{name}.png"
    shutil.copy2(src, dest)
    print(f"CAPTURE {dest} ({dest.stat().st_size} bytes)", flush=True)
    return dest


def main() -> int:
    wztc_ops.set_bridge(bridge_client.chat_bridge)
    wztc_ops.reset_plan_session_flags()

    print(f"EDGES up={UP} dn={DN} dy={DN[1]-UP[1]}", flush=True)

    print("CLEAR keep_alignments=False ...", flush=True)
    cleared = wztc_ops.clear_plan_elements(keep_alignments=False)
    print("cleared:", json.dumps(cleared, default=str)[:400], flush=True)

    print("BUILD order table 619-311 ...", flush=True)
    table = wztc_ops.build_wztc_order_table(
        speed=45,
        road_type="Non-Freeway",
        lane_width=12,
        shoulder_width="8 ft",
        sheet_num="619-311",
        area_type="RURAL",
    )
    print(
        "order table:",
        table.get("status"),
        "rows",
        len(table.get("rows") or []),
        "walkMax",
        max((float(w["stationFt"]) for w in (table.get("stationWalk") or [])), default=0),
        flush=True,
    )

    print("RUN_SHEET_BUILD ...", flush=True)
    t0 = time.time()
    result = wztc_ops.run_sheet_build(
        upstream_edge=UP,
        downstream_edge=DN,
        outward_sign=-1.0,
        include_visual_qa=True,
        clear_prior_stations=True,
        force=False,
    )
    print(f"run_sheet_build in {time.time()-t0:.1f}s", flush=True)
    print(json.dumps({
        "status": result.get("status"),
        "sheet": result.get("sheet"),
        "phases": result.get("phases"),
        "planCurrentStep": result.get("planCurrentStep"),
        "nextTool": result.get("nextTool"),
        "allComplete": (result.get("planStatus") or {}).get("allComplete"),
    }, indent=2, default=str), flush=True)

    mid_x = (UP[0] + DN[0]) / 2.0
    paths = [
        capture("rebuild_h_overview", mid_x, Y, 5500, 900),
        capture("rebuild_h_work", mid_x, Y, 500, 250),
        capture("rebuild_h_upstream", mid_x + 1200, Y, 2200, 700),
        capture("rebuild_h_downstream", mid_x - 400, Y, 900, 400),
        capture("rebuild_h_pv", DN[0] + 80, Y, 350, 220),
        capture("rebuild_h_ap", mid_x + 600, Y, 500, 300),
    ]
    print("DONE captures:", [p.name for p in paths], flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
