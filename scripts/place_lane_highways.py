"""Place 1-, 4-, and 5-lane highway linework below the existing 3-lane band."""
from __future__ import annotations

import sys
import time
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import bridge_client
import view_capture
import wztc_ops

X0, X1 = 23310.0, 24310.0
DASH, GAP, LANE = 10.0, 30.0, 12.0
PERIOD = DASH + GAP


def place_solid(y: float) -> str:
    r = wztc_ops.place_polyline(
        [[X0, y, 0.0], [X1, y, 0.0]], reason=f"solid edge y={y}")
    eid = str(r.get("elementId") or "")
    wztc_ops.change_element_level(eid, "Default", own_element_only=True, reason="align-like")
    wztc_ops.change_element_symbology(
        eid, color=0, weight=0, own_element_only=True, reason="solid")
    return eid


def place_dashed_row(y: float) -> list[str]:
    ids: list[str] = []
    x = X0
    while x + DASH <= X1 + 1e-9:
        x1 = min(x + DASH, X1)
        r = wztc_ops.place_polyline(
            [[x, y, 0.0], [x1, y, 0.0]], reason=f"10/30 dash y={y}")
        eid = str(r.get("elementId") or "")
        wztc_ops.change_element_level(eid, "Default", own_element_only=True, reason="align-like")
        wztc_ops.change_element_symbology(
            eid, color=0, weight=0, own_element_only=True, reason="dash seg")
        ids.append(eid)
        x += PERIOD
    return ids


def place_highway(n_lanes: int, top_y: float) -> dict:
    """n_lanes travel lanes => 2 solids + (n_lanes-1) dashed separators, 12 ft apart."""
    print(f"=== {n_lanes}-lane top_y={top_y} ===", flush=True)
    solids: list[str] = []
    dashed_rows: list[dict] = []
    for i in range(n_lanes + 1):
        y = top_y - i * LANE
        if i == 0 or i == n_lanes:
            solids.append(place_solid(y))
        else:
            segs = place_dashed_row(y)
            dashed_rows.append({"y": y, "n": len(segs)})
    bot = top_y - n_lanes * LANE
    print(f"  solids={solids} dashed_rows={dashed_rows} bot={bot}", flush=True)
    return {"solids": solids, "dashed_rows": dashed_rows, "bot": bot}


def main() -> int:
    wztc_ops.set_bridge(bridge_client.chat_bridge)
    # Existing 3-lane top at 295000; each iteration 200 ft below previous top.
    base_top = 295000.0
    specs = [
        (1, base_top - 200.0),
        (4, base_top - 400.0),
        (5, base_top - 600.0),
    ]
    for n, ty in specs:
        place_highway(n, ty)

    view_capture.navigate_view(23810, 294580, 1100, 520, view_num=1)
    time.sleep(0.45)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / "lane_highways_1_4_5.png"
    shutil.copy2(src, dest)
    print("CAPTURE", dest, dest.stat().st_size, flush=True)
    print("TOPS", [(n, ty) for n, ty in specs], flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
