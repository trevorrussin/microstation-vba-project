"""Capture junction + probe arrow cell orientation via correct adjust_view API."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)

OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
OUT.mkdir(parents=True, exist_ok=True)
LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"


def snap(name: str, cx: float, cy: float, w: float) -> None:
    r = wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=w)
    print("adjust", name, r.get("status") if isinstance(r, dict) else r)
    cap = wztc_ops.capture_view()
    path = Path(cap["path"]) if isinstance(cap, dict) else None
    if path and path.exists():
        dst = OUT / f"{name}.png"
        dst.write_bytes(path.read_bytes())
        print("saved", dst)
    else:
        print("cap fail", cap)


snap("arrow_junc_overview", 23800, 292500, 280)
snap("arrow_side_south", 23800, 292360, 160)
snap("arrow_side_north", 23800, 292640, 160)

# Orientation probes away from junction
base_x, base_y = 24500.0, 292500.0
for i, ang in enumerate([0, 90, 180, 270]):
    r = wztc_ops.place_cell(
        "SAS", base_x + i * 45, base_y, 0, float(ang),
        library_path=LIB,
        reason=f"arrow orientation probe angle={ang}",
    )
    print("place", ang, r.get("status"), r.get("elementId") if isinstance(r, dict) else r)

snap("arrow_angle_probes", base_x + 67.5, base_y, 220)

# Also find any CELL south that might be engineer example (wider)
rows = wztc_ops.find_elements_near(23800, 292300, 600, type_filter="CELL")
print("cells south-ish", len(rows) if isinstance(rows, list) else rows)
if isinstance(rows, list):
    for e in rows[:40]:
        print(e)
