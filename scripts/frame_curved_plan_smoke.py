"""Frame and capture the curved plan dim smoke."""
from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import view_capture  # noqa: E402

view_capture.navigate_view(83200, 287120, 450, 400, view_num=1)
time.sleep(2.5)
dest = ROOT / "Bridge" / "captures" / "smoke_curved_plan_dim.png"
shutil.copy2(view_capture.capture_microstation(), dest)
print("saved", dest)
