"""Rebuild smoke intersections with shared-option + RH-lane arrows."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")

# Fresh empty area
jx, jy = 27000.0, 291200.0

r1 = wztc_ops.place_orthogonal_intersection(
    jx, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=240,
    secondary_stub_ft=100,
    primary_lanes=4,
    secondary_lanes=2,
    reason="arrow QA v3 continuous 2+2 shared",
)
print("cont", r1.get("status"), "arrowCount", r1.get("arrowCount"), r1.get("errors"))
cells = {}
for a in r1.get("arrows") or []:
    cells.setdefault(a.get("arm"), []).append(
        (a.get("cellName"), round(a.get("angleDeg") or 0, 1),
         round(a.get("x") or 0, 1), round(a.get("y") or 0, 1))
    )
for arm, rows in sorted(cells.items()):
    print(" ", arm, rows)

r2 = wztc_ops.place_orthogonal_intersection(
    jx + 500, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=280,
    secondary_stub_ft=110,
    primary_lanes=6,
    secondary_lanes=2,
    reason="arrow QA v3 continuous 3+3 shared",
)
print("6lane", r2.get("status"), "arrowCount", r2.get("arrowCount"))
for a in (r2.get("arrows") or []):
    if a.get("arm") == "primary_neg":
        print(" ", a.get("cellName"), round(a.get("angleDeg") or 0, 1),
              round(a.get("x") or 0, 1), round(a.get("y") or 0, 1))

r3 = wztc_ops.place_orthogonal_intersection(
    jx + 1000, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=280,
    secondary_stub_ft=110,
    primary_lanes=6,
    secondary_lanes=2,
    primary_lanes_out=2,
    reason="arrow QA v3 dedicated 3to2",
)
print("ded", r3.get("status"), "arrowCount", r3.get("arrowCount"))
for a in (r3.get("arrows") or []):
    if a.get("arm") == "primary_neg":
        print(" ", a.get("cellName"), round(a.get("angleDeg") or 0, 1))

for name, cx, cy, w in [
    ("qa_v3_cont", jx, jy, 280),
    ("qa_v3_cont_west", jx - 80, jy, 100),
    ("qa_v3_cont_south", jx, jy - 80, 100),
    ("qa_v3_6lane", jx + 500, jy, 300),
    ("qa_v3_ded", jx + 1000, jy, 300),
    ("qa_v3_ded_west", jx + 920, jy, 120),
]:
    wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=w)
    p = Path(wztc_ops.capture_view()["path"])
    (OUT / f"{name}.png").write_bytes(p.read_bytes())
    print("saved", name)
