"""Clean final smoke: shared options + RH lanes + facing (no +180)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")

jx, jy = 30000.0, 291200.0
r = wztc_ops.place_orthogonal_intersection(
    jx, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=240,
    secondary_stub_ft=100,
    primary_lanes=4,
    secondary_lanes=2,
    reason="arrow QA FINAL shared+RH",
)
print("status", r.get("status"), "arrowCount", r.get("arrowCount"))
for a in r.get("arrows") or []:
    print(
        a.get("arm"), a.get("cellName"),
        round(a.get("angleDeg") or 0, 1),
        round(a.get("x") or 0, 1), round(a.get("y") or 0, 1),
    )

r2 = wztc_ops.place_orthogonal_intersection(
    jx + 450, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=280,
    secondary_stub_ft=110,
    primary_lanes=6,
    secondary_lanes=2,
    reason="arrow QA FINAL 3+3 shared",
)
print("6lane", r2.get("status"), [
    a.get("cellName") for a in (r2.get("arrows") or [])
    if a.get("arm") == "primary_neg"
])

r3 = wztc_ops.place_orthogonal_intersection(
    jx + 900, jy,
    primary_road_type="two_way",
    secondary_road_type="two_way",
    primary_length_ft=280,
    secondary_stub_ft=110,
    primary_lanes=6,
    secondary_lanes=2,
    primary_lanes_out=2,
    reason="arrow QA FINAL 3to2 dedicated",
)
print("ded", r3.get("status"), [
    a.get("cellName") for a in (r3.get("arrows") or [])
    if a.get("arm") == "primary_neg"
])

for name, cx, cy, w in [
    ("qa_final_cont", jx, jy, 280),
    ("qa_final_west", jx - 80, jy, 110),
    ("qa_final_south", jx, jy - 85, 110),
    ("qa_final_6lane", jx + 450, jy, 300),
    ("qa_final_ded_west", jx + 820, jy, 130),
]:
    wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=w)
    p = Path(wztc_ops.capture_view()["path"])
    (OUT / f"{name}.png").write_bytes(p.read_bytes())
    print("saved", name)
