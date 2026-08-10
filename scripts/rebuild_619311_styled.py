"""Style-fixed horizontal 619-311 rebuild (dy=0), north of engineer reference."""
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

# Clean band north of prior agent rebuild (230k) and engineer reference (~214k)
Y = 232000.0
WORK_LEN = 150.0
UP = [1020150.0, Y, 0.0]
DN = [1020150.0 - WORK_LEN, Y, 0.0]


def capture(name: str, x: float, y: float, w: float, h: float) -> Path:
    view_capture.navigate_view(x, y, w, h, view_num=1)
    time.sleep(0.35)
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

    table = wztc_ops.build_wztc_order_table(
        speed=45,
        road_type="Non-Freeway",
        lane_width=12,
        shoulder_width="8 ft",
        sheet_num="619-311",
        area_type="RURAL",
    )
    print("order table:", table.get("status"), "rows", len(table.get("rows") or []), flush=True)

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
        "phases": result.get("phases"),
        "gateFailures": (result.get("geometry") or {}).get("gateFailures")
            if isinstance(result.get("geometry"), dict) else result.get("gateFailures"),
        "placed": (result.get("geometry") or {}).get("placedCount")
            if isinstance(result.get("geometry"), dict) else None,
        "errors": (result.get("geometry") or {}).get("errors")
            if isinstance(result.get("geometry"), dict) else None,
    }, default=str, indent=2)[:2500], flush=True)

    mid = 0.5 * (UP[0] + DN[0])
    capture("style_rebuild_overview", mid, Y, 2800, 500)
    capture("style_rebuild_work", mid, Y, 900, 350)
    capture("style_rebuild_taper", UP[0] + 400, Y, 1000, 400)
    capture("style_rebuild_signs", UP[0] + 1800, Y, 1200, 400)
    capture("style_rebuild_downstream", DN[0] - 80, Y, 900, 350)
    return 0 if result.get("status") == "OK" else 1


if __name__ == "__main__":
    raise SystemExit(main())
