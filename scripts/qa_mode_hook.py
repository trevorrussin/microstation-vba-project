"""PreToolUse hook (Edit, Write, MultiEdit, NotebookEdit, Bash, PowerShell): keep production work out of the tooling.

While a production QA task is open, refuse edits to the files that define what the agent can do (bridge, VBA modules, the
C# add-in, the key-in registry, workflow contracts, ladders, QA code, tests) and commands that load or build them. A
development task (`qa_gate start --mode development --scratch <file.dgn|none>`) lifts this. See mcp-server/modes.py.

Claude Code (.claude/settings.json "PreToolUse"): stdin JSON, exit 2 + stderr = refused. A broken hook never blocks.
"""
from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))


def assess(tool: str, tool_input: dict) -> str:
    import modes
    if tool in ("Edit", "Write", "MultiEdit", "NotebookEdit"):
        return modes.assess_path(tool_input.get("file_path") or tool_input.get("notebook_path") or "")
    if tool in ("Bash", "PowerShell"):
        return modes.assess_command(tool_input.get("command", ""))
    return ""


def main() -> int:
    try:
        data = json.load(sys.stdin)
        msg = assess(data.get("tool_name", ""), data.get("tool_input") or {})
    except Exception:
        return 0
    if msg:
        print(msg, file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
