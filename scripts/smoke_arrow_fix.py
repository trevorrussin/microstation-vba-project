"""Live smoke: fixed arrow angle + lanes_out dedicated math."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")

# Continuous 2+2 primary / 1+1 secondary — all SAS, facing toward box
r1 = wztc_ops.place_orthogonal_intersection(
    24500, 291500,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=240,
    secondary_stub_ft=100,
    primary_lanes=4,
    secondary_lanes=2,
    reason="arrow-fix smoke continuous",
)
print("continuous", r1.get("status"), "arrows", len(r1.get("turnArrows") or r1.get("arrowsPlaced") or []),
      "err", r1.get("errors"), list(r1.keys())[:12])

# 3+3 → 2+2 dedicated left on primary
r2 = wztc_ops.place_orthogonal_intersection(
    25000, 291500,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=280,
    secondary_stub_ft=110,
    primary_lanes=6,
    secondary_lanes=2,
    primary_lanes_out=2,
    reason="arrow-fix smoke dedicated 3to2",
)
print("dedicated", r2.get("status"), list(r1.keys()))
print("r2 keys sample", {k: r2.get(k) for k in (
    "status", "turnArrowCount", "arrowsPlaced", "errors", "note"
) if k in r2 or True})

wztc_ops.adjust_view(center_x=24500, center_y=291500, width=300, height=300)
cap = wztc_ops.capture_view()
p = Path(cap["path"])
dst = OUT / "arrow_fix_continuous.png"
dst.write_bytes(p.read_bytes())
print("saved", dst)

wztc_ops.adjust_view(center_x=25000, center_y=291500, width=320, height=320)
cap = wztc_ops.capture_view()
p = Path(cap["path"])
dst = OUT / "arrow_fix_dedicated.png"
dst.write_bytes(p.read_bytes())
print("saved", dst)

# Zoom west approach of continuous
wztc_ops.adjust_view(center_x=24420, center_y=291500, width=120, height=80)
cap = wztc_ops.capture_view()
p = Path(cap["path"])
dst = OUT / "arrow_fix_west_approach.png"
dst.write_bytes(p.read_bytes())
print("saved", dst)
