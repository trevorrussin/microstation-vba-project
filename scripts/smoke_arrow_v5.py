"""Smoke: overlap LSR, tip fix, ONLY spacing, 3→2 asymmetric median."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")

jx, jy = 32000.0, 291200.0

r1 = wztc_ops.place_orthogonal_intersection(
    jx, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=240,
    secondary_stub_ft=100,
    primary_lanes=4,
    secondary_lanes=2,
    reason="arrow QA v5 shared overlap+tip",
)
print("shared", r1.get("status"), r1.get("arrowCount"))
for a in r1.get("arrows") or []:
    if a.get("arm") in ("primary_neg", "secondary_right"):
        print(
            " ", a.get("arm"), a.get("cellName"),
            round(a.get("angleDeg") or 0, 1),
            round(a.get("x") or 0, 1), round(a.get("y") or 0, 1),
        )

r2 = wztc_ops.place_orthogonal_intersection(
    jx + 500, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=300,
    secondary_stub_ft=120,
    primary_lanes=6,
    secondary_lanes=2,
    primary_lanes_out=2,
    reason="arrow QA v5 3to2 asymmetric",
)
print("drop", r2.get("status"), r2.get("arrowCount"))
for a in r2.get("arrows") or []:
    if a.get("arm") == "primary_neg":
        print(
            " ", a.get("cellName"),
            round(a.get("angleDeg") or 0, 1),
            round(a.get("x") or 0, 1), round(a.get("y") or 0, 1),
        )

for name, cx, cy, w in [
    ("qa_v5_shared", jx, jy, 280),
    ("qa_v5_south", jx, jy - 85, 110),
    ("qa_v5_west", jx - 80, jy, 110),
    ("qa_v5_drop", jx + 500, jy, 320),
    ("qa_v5_drop_west", jx + 400, jy, 140),
]:
    wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=w)
    p = Path(wztc_ops.capture_view()["path"])
    (OUT / f"{name}.png").write_bytes(p.read_bytes())
    print("saved", name)
