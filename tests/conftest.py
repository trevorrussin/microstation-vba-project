"""Pytest path bootstrap — mcp-server modules import as top-level names."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
MCP = ROOT / "mcp-server"
if str(MCP) not in sys.path:
    sys.path.insert(0, str(MCP))


@pytest.fixture(autouse=True)
def _isolate_sheet_plan_path(tmp_path, monkeypatch):
    """Never let a test write Bridge/sheet-plan.json — the live runtime file.

    Found 2026-08-20: any test that locks designer inputs + a corridor makes
    PlanSession.sheet_plan_active() True, and every corridor/lateral/build
    call after that calls _save_sheet_plan(), which writes straight to the
    module-global wztc_ops.SHEET_PLAN_PATH with no test hook — the real
    project file, not a fixture. A single `pytest tests/` run clobbered the
    live 619-311 proof-build's resumability record with fixture data
    (sheetNum=619-311, visual_qa_passed=True from a mocked test), and this
    was already happening before this session's changes — test_corridor_path
    reproduces it against an unmodified checkout. Autouse so no test author
    has to remember to add this; tests that want the real path can still
    read wztc_ops.SHEET_PLAN_PATH before this fixture's monkeypatch applies
    only if they capture it at import time (none currently do).
    """
    try:
        import wztc_ops
    except ImportError:
        return
    monkeypatch.setattr(wztc_ops, "SHEET_PLAN_PATH", tmp_path / "sheet-plan.json")
