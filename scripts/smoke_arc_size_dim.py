"""Smoke Arc Size + verify element range is real (not empty)."""
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
    cx, cy, r = 80000.0, 287000.0, 200.0
    path = []
    for i in range(0, 91, 5):
        a = math.radians(i)
        path.append([cx + r * math.cos(a), cy + r * math.sin(a)])
    ox = cx + (r + 25) * math.cos(math.radians(45))
    oy = cy + (r + 25) * math.sin(math.radians(45))

    direct = ops.place_arc_size_dimension(
        cx, cy, path[0][0], path[0][1], path[-1][0], path[-1][1],
        ox, oy, override_text="120'-0\"", reason="smoke arc-size 3pt")
    print("direct", direct)
    eid = str(direct.get("elementId") or "")
    rng = ops.get_elements_range([eid] if eid else [])
    print("range", rng)
    ok = float(rng.get("width") or 0) > 1.0 and float(rng.get("height") or 0) > 1.0
    print("geometry_ok", ok)

    hug = ops.place_path_hugging_dimension(
        path, "120", [ox, oy], reason="smoke hug arc-size")
    print("hug", hug)

    view_capture.navigate_view(cx + 80, cy + 120, 500, 400, view_num=1)
    time.sleep(0.6)
    dest = ROOT / "Bridge" / "captures" / "smoke_arc_size_dim.png"
    shutil.copy2(view_capture.capture_microstation(), dest)
    print("saved", dest.name)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
