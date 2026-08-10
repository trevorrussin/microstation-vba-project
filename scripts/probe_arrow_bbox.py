"""Use element range to infer SAS tip direction vs ACTIVE ANGLE."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"

bx, by = 29500.0, 290500.0
ids = []
for i, ang in enumerate([0, 90, 180, 270, -90]):
    x = bx + i * 50
    r = wztc_ops.place_cell("SAS", x, by, 0, float(ang), library_path=LIB, reason=f"range probe {ang}")
    eid = str(r.get("elementId") or "")
    ids.append((ang, eid, x, by))
    print("placed", ang, eid)

# Also SALS at 0 and -90
for ang in [0, -90, 90, 180]:
    x = bx + 300
    y = by + (0 if ang == 0 else 40 if ang == -90 else 80 if ang == 90 else 120)
    r = wztc_ops.place_cell("SALS", x, y, 0, float(ang), library_path=LIB, reason=f"SALS range {ang}")
    ids.append((f"SALS{ang}", str(r.get("elementId") or ""), x, y))

for ang, eid, x, y in ids:
    if not eid:
        continue
    rng = wztc_ops.get_elements_range([eid])
    if not isinstance(rng, dict):
        print(ang, eid, rng)
        continue
    cx = (rng["lowX"] + rng["highX"]) / 2
    cy = (rng["lowY"] + rng["highY"]) / 2
    dx = cx - x
    dy = cy - y
    w = rng["highX"] - rng["lowX"]
    h = rng["highY"] - rng["lowY"]
    # Tip bias: where center of bbox sits relative to origin
    print(
        f"ang={ang} origin=({x},{y}) bbox_c=({cx:.2f},{cy:.2f}) "
        f"delta=({dx:.2f},{dy:.2f}) size=({w:.2f}x{h:.2f})"
    )
