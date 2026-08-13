"""Smoke: sheet override + HIDE sentinel for intermediate segment text."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)


def main() -> int:
    # Three-seg chain near a quiet XY
    path = [
        [77000.0, 287500.0],
        [77040.0, 287500.0],
        [77080.0, 287500.0],
        [77120.0, 287500.0],
    ]
    r = ops.place_path_hugging_dimension(
        path, "120", [77060.0, 287480.0], reason="smoke hide chain")
    print("chain", r)
    # Single with explicit HIDE
    r2 = ops.place_dimension(
        77200, 287500, 77280, 287500, 77240, 287480,
        override_text="HIDE", reason="smoke hide single")
    print("hide", r2)
    r3 = ops.place_dimension(
        77300, 287500, 77380, 287500, 77340, 287480,
        override_text="120'-0\"", reason="smoke show single")
    print("show", r3)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
