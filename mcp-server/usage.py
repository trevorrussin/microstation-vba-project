"""Token-usage cost tracking, extracted from chat_driver.py (2026-08-04
split) since it's a self-contained concern: a pricing table plus a small
class that accumulates cost and appends one row per API response to a log
file.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

# $ per million tokens (Anthropic pricing, confirmed current). Cache write is
# priced off the base input rate at a TTL-dependent multiplier (1.25x for a
# 5-minute breakpoint, 2x for 1-hour); chat_driver.py only ever sets ttl="1h"
# (see run_turn), so WRITE_MULT is fixed at 2x rather than reading the TTL
# back out of usage -- the API doesn't report which TTL a
# cache_creation_input_tokens figure was billed at. Cache read is ~0.1x
# input. This is an estimate for in-app visibility, not a reconciliation of
# the actual invoice -- check https://console.anthropic.com/settings/usage
# for the authoritative number.
PRICING = {
    "claude-opus-5": {"input": 5.00, "output": 25.00},
    "claude-sonnet-5": {"input": 3.00, "output": 15.00},
}
CACHE_WRITE_MULT = 2.0   # 1h TTL
CACHE_READ_MULT = 0.1


class UsageTracker:
    """Accumulates token usage across the whole chat_driver.py process
    lifetime and appends one row per API response to usage_file, so cost
    is visible locally without checking the Anthropic Console. Historical
    sessions before this existed aren't recoverable from local data --
    only the Console has that."""

    def __init__(self, usage_file: Path):
        self.usage_file = usage_file
        self.total_cost_usd = 0.0

    def record(self, usage, model: str) -> float:
        rates = PRICING.get(model)
        if rates is None:
            return 0.0  # unknown model string -- don't guess a price

        input_tok = getattr(usage, "input_tokens", 0) or 0
        output_tok = getattr(usage, "output_tokens", 0) or 0
        cache_read = getattr(usage, "cache_read_input_tokens", 0) or 0

        # usage.cache_creation carries the exact 5m/1h split when present --
        # more precise than assuming every write used this file's 1h
        # cache_control TTL. Falls back to the flat field (older SDK
        # responses may not populate cache_creation) at the 1h rate, since
        # 1h is the only TTL this file ever requests.
        cache_creation = getattr(usage, "cache_creation", None)
        if cache_creation is not None:
            cache_write_5m = getattr(cache_creation, "ephemeral_5m_input_tokens", 0) or 0
            cache_write_1h = getattr(cache_creation, "ephemeral_1h_input_tokens", 0) or 0
            cache_write_cost = cache_write_5m * 1.25 + cache_write_1h * CACHE_WRITE_MULT
            cache_write = cache_write_5m + cache_write_1h
        else:
            cache_write = getattr(usage, "cache_creation_input_tokens", 0) or 0
            cache_write_cost = cache_write * CACHE_WRITE_MULT

        cost = (
            input_tok * rates["input"]
            + output_tok * rates["output"]
            + cache_write_cost * rates["input"]
            + cache_read * rates["input"] * CACHE_READ_MULT
        ) / 1_000_000

        self.total_cost_usd += cost

        line = "\t".join([
            datetime.now().isoformat(sep=" ", timespec="seconds"),
            model,
            f"input={input_tok}",
            f"output={output_tok}",
            f"cacheWrite={cache_write}",
            f"cacheRead={cache_read}",
            f"costUsd={cost:.4f}",
            f"runningTotalUsd={self.total_cost_usd:.4f}",
        ])
        with open(self.usage_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")

        return cost
