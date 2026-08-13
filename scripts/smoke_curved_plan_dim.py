"""Smoke curved plan dim — ArcElement must hug the tip arc."""
from __future__ import annotations

import math
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import view_capture  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)


def main() -> int:
    cx, cy, r = 83000.0, 287000.0, 200.0
    path = []
    for i in range(0, 91, 5):
        a = math.radians(i)
        path.append([cx + r * math.cos(a), cy + r * math.sin(a)])
    ox = cx + (r + 25) * math.cos(math.radians(45))
    oy = cy + (r + 25) * math.sin(math.radians(45))
    hug = ops.place_path_hugging_dimension(
        path, "120", [ox, oy], reason="smoke curved plan arc")
    print("hug", hug)
    eid = (hug.get("createdElementIds") or [None])[0]
    if eid:
        print("range", ops.get_elements_range([str(eid)]))
    view_capture.navigate_view(cx + 90, cy + 110, 480, 420, view_num=1)
    time.sleep(2.5)
    dest = ROOT / "Bridge" / "captures" / "smoke_curved_plan_dim.png"
    shutil.copy2(view_capture.capture_microstation(), dest)
    print("saved", dest.name)
    return 0 if hug.get("dimType") == "CurvedPlanArc" else 1


if __name__ == "__main__":
    raise SystemExit(main())
