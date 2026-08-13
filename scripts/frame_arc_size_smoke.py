"""Frame ArcSize element by range and capture."""
from __future__ import annotations

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
    # Latest smoke ids from previous run; re-place if needed
    r = ops.place_arc_size_dimension(
        80500, 287000,
        80700, 287000,
        80500, 287200,
        80500 + 225 * 0.7071, 287000 + 225 * 0.7071,
        override_text="120'-0\"",
        reason="smoke arc frame")
    print("placed", r)
    eid = str(r.get("elementId") or "")
    rng = ops.get_elements_range([eid])
    print("range", rng)
    cx = float(rng["centerX"])
    cy = float(rng["centerY"])
    w = max(float(rng["width"]) * 1.8, 300)
    h = max(float(rng["height"]) * 1.8, 300)
    view_capture.navigate_view(cx, cy, w, h, view_num=1)
    time.sleep(0.8)
    dest = ROOT / "Bridge" / "captures" / "smoke_arc_size_dim.png"
    shutil.copy2(view_capture.capture_microstation(), dest)
    print("saved", dest, "at", cx, cy, w, h)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
