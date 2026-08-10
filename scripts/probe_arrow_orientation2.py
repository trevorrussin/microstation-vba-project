"""Place orientation probes with reference lines; capture close-up."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"

# Reference: horizontal line east, vertical line north, cells at angles
bx, by = 26000.0, 291000.0
# East-pointing guide line
wztc_ops.place_polyline([(bx, by - 20), (bx + 80, by - 20)], reason="probe guide east")
# North-pointing guide
wztc_ops.place_polyline([(bx - 20, by), (bx - 20, by + 80)], reason="probe guide north")

# Place SAS at 0,90,180,270 along a row; place SALS too at same angles below
for i, ang in enumerate([0, 90, 180, 270, -90]):
    x = bx + 10 + i * 35
    r = wztc_ops.place_cell("SAS", x, by + 30, 0, float(ang), library_path=LIB, reason=f"probe SAS {ang}")
    print("SAS", ang, r.get("elementId"))
    r = wztc_ops.place_cell("SALS", x, by + 60, 0, float(ang), library_path=LIB, reason=f"probe SALS {ang}")
    print("SALS", ang, r.get("elementId"))

wztc_ops.adjust_view(center_x=bx + 80, center_y=by + 40, width=200, height=120)
p = Path(wztc_ops.capture_view()["path"])
(OUT / "arrow_angle_probes2.png").write_bytes(p.read_bytes())
print("saved probes2")

# Also: for west approach travel=+X, compare old vs new angle at same point
# old atan2(0,1)=0, new atan2(-1,0)=-90
wx, wy = 24580.0, 291185.0  # near continuous west approach south of yellow?
wztc_ops.place_cell("SAS", wx, wy - 15, 0, 0.0, library_path=LIB, reason="compare angle 0")
wztc_ops.place_cell("SAS", wx, wy + 15, 0, -90.0, library_path=LIB, reason="compare angle -90")
wztc_ops.adjust_view(center_x=wx, center_y=wy, width=80, height=60)
p = Path(wztc_ops.capture_view()["path"])
(OUT / "arrow_angle_compare_west.png").write_bytes(p.read_bytes())
print("saved compare")
