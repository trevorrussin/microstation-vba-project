"""Definitive facing test: SAS only, west of a stop-bar stand-in."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"

# Fake eastbound approach: yellow above, white below, stop bar on right (intersection)
bx, by = 29000.0, 291000.0
# stop bar (intersection to the EAST)
wztc_ops.place_polyline([(bx + 40, by - 20), (bx + 40, by + 5)], reason="face test stop")
# lane centerline-ish
wztc_ops.place_polyline([(bx - 40, by - 6), (bx + 38, by - 6)], reason="face test lane")

# Travel toward intersection = +X (east). Candidate angles:
for i, ang in enumerate([-90, 90, 0, 180, 270]):
    r = wztc_ops.place_cell(
        "SAS", bx - 20, by - 6 - i * 0, 0, float(ang),
        library_path=LIB, reason=f"face SAS {ang}",
    )
    # place vertically separated
print("placing row")
for i, ang in enumerate([-90, 90, 0, 180, 270]):
    y = by - 40 - i * 25
    wztc_ops.place_polyline([(bx + 40, y - 8), (bx + 40, y + 8)], reason="stop")
    wztc_ops.place_cell("SAS", bx - 10, y, 0, float(ang), library_path=LIB, reason=f"face {ang}")
    print(ang)

wztc_ops.adjust_view(center_x=bx, center_y=by - 90, width=140, height=160)
p = Path(wztc_ops.capture_view()["path"])
(OUT / "qa_face_angles.png").write_bytes(p.read_bytes())
print("saved qa_face_angles")
