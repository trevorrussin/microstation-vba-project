"""Inspect Arc Size smoke elements and capture view."""
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
    print(ops.get_elements_range(["164691", "164692"]))
    view_capture.navigate_view(79000, 287100, 600, 500, view_num=1)
    time.sleep(0.8)
    src = Path(view_capture.capture_microstation())
    dest = ROOT / "Bridge" / "captures" / "smoke_arc_size_dim.png"
    shutil.copy2(src, dest)
    print("saved", dest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
