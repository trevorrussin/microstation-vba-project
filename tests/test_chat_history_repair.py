"""Regression: SDK content blocks must be visible to tool-pairing repair."""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))
import chat_history as ch


def test_repair_drops_consecutive_assistant_tool_uses_as_dicts():
    msgs = [
        {"role": "user", "content": "hi"},
        {"role": "assistant", "content": [
            {"type": "tool_use", "id": "toolu_a", "name": "x", "input": {}},
        ]},
        {"role": "assistant", "content": [
            {"type": "tool_use", "id": "toolu_b", "name": "y", "input": {}},
        ]},
        {"role": "assistant", "content": "done"},
    ]
    ch._repair_tool_pairing(msgs)
    assert all(
        not ch._msg_tool_use_ids(m) or (
            i + 1 < len(msgs)
            and msgs[i + 1].get("role") == "user"
            and ch._msg_tool_use_ids(m).issubset(ch._msg_tool_result_ids(msgs[i + 1]))
        )
        for i, m in enumerate(msgs)
        if isinstance(m, dict)
    )
    assert not any(ch._msg_tool_use_ids(m) for m in msgs)


def test_repair_sees_sdk_style_tool_use_objects():
    """Live bug 2026-08-05: runner appends SDK blocks; repair only saw dicts
    so unanswered tool_use survived save → next turn API 400."""
    tu = SimpleNamespace(type="tool_use", id="toolu_sdk1", name="x", input={})
    msgs = [
        {"role": "user", "content": "hi"},
        {"role": "assistant", "content": [tu]},
        {"role": "assistant", "content": "final without tool_result"},
    ]
    ch._repair_tool_pairing(msgs)
    assert not any(ch._msg_tool_use_ids(m) for m in msgs)


def test_msg_tool_use_ids_reads_sdk_objects():
    tu = SimpleNamespace(type="tool_use", id="toolu_z", name="x", input={})
    ids = ch._msg_tool_use_ids({"role": "assistant", "content": [tu]})
    assert ids == {"toolu_z"}
