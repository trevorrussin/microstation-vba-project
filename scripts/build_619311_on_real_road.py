"""Build 619-311 on real EB corridor — Urban 55, 100 ft work area.

Clean band: Y_NORTH_OUTER=288200 (same band as prior polish; G20 closed south).
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")

Y_NORTH_OUTER = 288200.0
X1, X2 = 32000.0, 36000.0
LANE = 12.0
YELLOW_GAP = 2.0
SHOULDER = 8.0
Y_LANE_LINE = Y_NORTH_OUTER - (2 * LANE + YELLOW_GAP + LANE)  # 288162
HALF_LEN = LANE + SHOULDER
WA_UP_X = 34500.0
WA_DN_X = 34600.0
YC_VIEW = Y_NORTH_OUTER - (2 * LANE + YELLOW_GAP / 2.0)

print("Y_NORTH_OUTER", Y_NORTH_OUTER, "Y_LANE_LINE", Y_LANE_LINE)
print("clear", wztc_ops.clear_plan_elements(keep_alignments=False).get("deleted"))

ot = wztc_ops.build_wztc_order_table(
    speed=55, road_type="Non-Freeway", lane_width=12,
    shoulder_width=">= 8 ft", sheet_num="619-311", area_type="URBAN",
)
print("order_table", ot.get("status"))

up = [WA_UP_X, Y_LANE_LINE, 0.0]
dn = [WA_DN_X, Y_LANE_LINE, 0.0]
lat = wztc_ops.resolve_sheet_lateral(
    up, dn, closed_side="right", real_road_edge=True,
)
print("lateral", lat.get("outward_sign"), lat.get("half_len"),
      lat.get("closed_outward"), lat.get("note", "")[:80])

result = wztc_ops.run_sheet_build(
    upstream_edge=up, downstream_edge=dn,
    arrow_panel_choice="trailer", include_visual_qa=True, force=True,
)
print("build", result.get("status"))
for p in result.get("phases") or []:
    print(" phase", p.get("phase"), p.get("result") or p.get("note") or "")
print("realRoadNext", result.get("realRoadNext", "")[:100])

# run_sheet_build already deletes guides when real_road_edge; corridor after wipe
road = wztc_ops.place_two_way_highway(
    lanes=4, x1=X1, y1=Y_NORTH_OUTER, x2=X2, y2=Y_NORTH_OUTER,
    lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP, shoulder_width_ft=SHOULDER,
    side="right", reason="619-311 clean band north-outer=288200",
)
print("corridor", road.get("status"))

wztc_ops.adjust_view(center_x=34000, center_y=YC_VIEW, width=4500, height=500)

for name, cx, cy, w, h in [
    ("qa_311_v4_overview", 34000, YC_VIEW, 4200, 500),
    ("qa_311_v4_work", 34550, Y_LANE_LINE, 400, 120),
    ("qa_311_v4_taper", 33600, Y_LANE_LINE, 1000, 150),
    ("qa_311_v4_g20", 34750, Y_LANE_LINE - 10, 500, 120),
    ("qa_311_v4_ap", 33200, Y_LANE_LINE - 10, 300, 120),
]:
    wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=h)
    p = Path(wztc_ops.capture_view()["path"])
    (OUT / f"{name}.png").write_bytes(p.read_bytes())
    print("saved", name)

aps = [
    e for e in wztc_ops.find_elements_near(33205, Y_LANE_LINE - 20, 80, "CELL", force=True)
    if e.get("cellName") == "TWZAP_P"
]
print("AP count", len(aps), [(e["elementId"], e["cy"]) for e in aps])
g20 = [
    e for e in wztc_ops.find_elements_near(34730, Y_LANE_LINE - 20, 100, force=True)
    if e.get("type") == "CELL" or (e.get("cellName") or "").upper().startswith("G20")
]
print("G20 area cells", [(e.get("cellName"), e.get("cy"), e.get("cx")) for e in g20[:10]])
# Expect G20 face/post south of lane line (closed shoulder ~288142 tip)
g20_faces = [e for e in g20 if str(e.get("cellName", "")).upper().startswith("G20")]
if g20_faces:
    cy = float(g20_faces[0]["cy"])
    print("G20 Y", cy, "south_of_lane" if cy < Y_LANE_LINE else "NORTH_WRONG")
sc = wztc_ops.get_geometry_scorecard("619-311")
print("scorecard", sc.get("passed"))
print(f"LOOK HERE: Y~{YC_VIEW:.0f} (north outer {Y_NORTH_OUTER}), X 32000..36000")
