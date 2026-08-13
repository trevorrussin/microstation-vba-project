"""Smoke: ny_Plan SizeArrow with OverrideText + short curved chain."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)

print("single", ops.place_dimension(
    76000, 287900, 76100, 287900, 76050, 287885,
    override_text="100'-0\"", reason="override smoke"))

print("chain", ops.place_path_hugging_dimension(
    [[76200, 287900], [76250, 287905], [76300, 287920], [76340, 287950]],
    "120", [76270, 287880], reason="chain smoke"))
