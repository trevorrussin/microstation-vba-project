"""Pytest path bootstrap — mcp-server modules import as top-level names."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MCP = ROOT / "mcp-server"
if str(MCP) not in sys.path:
    sys.path.insert(0, str(MCP))
