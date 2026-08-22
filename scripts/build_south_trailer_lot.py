"""Place south trailer lot by copying upper-lot block 6A741 at aerial cluster positions.

Uses wztc_ops via MCP bridge — run while AutoCAD MCP session is active.
Template block 6A741: insert (14187.174, -4588.244), rotation 270 deg (E-W trailers).
Spacing matches upper lot: 240 ft in X, 120 ft between rows in Y.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops

TEMPLATE_ID = "6A741"
TEMPLATE_X = 14187.174280896772
TEMPLATE_Y = -4588.243723261282

# Aerial cluster positions (CAD coords, derived from satellite layout + existing anchors)
CLUSTERS: dict[str, list[tuple[float, float]]] = {
    # West column along property line — ~13 trailers stacked north-to-south
    "west_column": [(11880, y) for y in range(-5050, -6251, -120)],
    # North-central row just below existing lot cutoff
    "north_row": [(x, -5050) for x in range(12400, 14601, 240)],
    # Center back-to-back double block (2 rows)
    "center_front": [(x, -5320) for x in range(13000, 14401, 240)],
    "center_rear": [(x, -5440) for x in range(13000, 14401, 240)],
    # Southwest block (2 short rows)
    "sw_row1": [(x, -5800) for x in range(11900, 13901, 240)],
    "sw_row2": [(x, -5920) for x in range(11900, 13901, 240)],
    # South bottom row along south curb — long row per aerial
    "south_row": [(x, -6180) for x in range(12000, 16801, 240)],
}


def copy_trailer(tx: float, ty: float, label: str) -> str:
    dx = tx - TEMPLATE_X
    dy = ty - TEMPLATE_Y
    r = wztc_ops.copy_element(
        TEMPLATE_ID, dx, dy, own_element_only=False,
        reason=f"south lot {label} ({tx},{ty})")
    eid = str(r.get("elementId") or r.get("createdElementIds") or "")
    print(f"  {label}: ({tx},{ty}) -> {eid}", flush=True)
    return eid


def main() -> int:
    total = 0
    for cluster_name, positions in CLUSTERS.items():
        print(f"=== {cluster_name} ({len(positions)} trailers) ===", flush=True)
        for i, (tx, ty) in enumerate(positions):
            copy_trailer(tx, ty, f"{cluster_name}[{i}]")
            total += 1
    print(f"Placed {total} trailers total.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
