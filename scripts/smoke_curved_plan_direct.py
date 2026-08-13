"""Direct PLACE_CURVED_PLAN_DIMENSION call."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)


def main() -> int:
    r = ops.place_curved_plan_dimension(
        83100, 287000,
        83300, 287000,
        83100, 287200,
        83276.78, 287176.78,
        override_text="120'-0\"",
        reason="direct curved plan",
    )
    print(r)
    return 0 if r.get("status") == "OK" else 1


if __name__ == "__main__":
    raise SystemExit(main())
