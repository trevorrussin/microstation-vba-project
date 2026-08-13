"""Compare Arc Size placement variants — find one that hugs the roadside."""
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
    # Known 90deg roadside arc
    cx, cy, r = 82000.0, 287000.0, 200.0
    a0, a1 = 0.0, math.pi / 2
    x1 = cx + r * math.cos(a0)
    y1 = cy + r * math.sin(a0)
    x2 = cx + r * math.cos(a1)
    y2 = cy + r * math.sin(a1)
    # Dim arc slightly outside
    rd = r + 25.0
    amid = 0.5 * (a0 + a1)
    hx = cx + rd * math.cos(amid)
    hy = cy + rd * math.sin(amid)

    # Also place the tip arc itself so we can see concentricity
    mid = [cx + r * math.cos(amid), cy + r * math.sin(amid), 0]
    print("tipArc", ops.place_arc(x1, y1, mid[0], mid[1], x2, y2, reason="tip arc guide"))

    print("arcSize", ops.place_arc_size_dimension(
        cx, cy, x1, y1, x2, y2, hx, hy,
        override_text="120'-0\"", reason="variant hug outside"))

    rng = ops.get_elements_range(
        str(ops.place_arc_size_dimension(
            cx, cy, x1, y1, x2, y2, hx, hy,
            override_text="120'-0\"", reason="variant2")).get("elementId", "").split()
    )
    # Re-get last from journal is hard; just frame known area
    view_capture.navigate_view(cx + 80, cy + 120, 500, 450, view_num=1)
    time.sleep(2.5)
    dest = ROOT / "Bridge" / "captures" / "smoke_arc_hug.png"
    shutil.copy2(view_capture.capture_microstation(), dest)
    print("saved", dest.name, "expect dim arc near tip arc r=200..225")
    print("height_pt", hx, hy)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
