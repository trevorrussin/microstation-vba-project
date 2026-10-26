"""Task modes: production engineering vs capability development.

The 2026-10-24 circles run asked for a small display change and ended up rebuilding an add-in and probing OpenRoads
internals on the engineer's real file. Two kinds of work, two sets of permissions:

  production    (default) Change the engineer's drawing with established capabilities only. Refused while such a task is open:
                editing the tooling that defines capabilities (bridge, VBA modules, the C# add-in, the key-in registry, workflow
                contracts, ladders, QA code, tests) and running the scripts that load or build it (hot_reload, ord-addin build).
  development   Build or investigate a capability. Allowed to change the tooling. Must NOT touch a real drawing:
                `qa_gate start --mode development --scratch <file.dgn>` names the scratch DGN that is open (checked against
                the active file at start and again at close), or `--scratch none` for code-only work, in which case the bridge
                refuses every drawing edit (task_guard).

Only an open task triggers the rule. With no task open (ordinary code sessions) nothing here applies. The check is a static look
at the tool call (path, or command text); code that builds a path at run time, or a script that edits these files itself, evades it.
It stops the accidental case.
"""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TASKS = ROOT / "outputs" / "qa" / "tasks"
WINDOW_S = 12 * 3600
MODES = ("production", "development")

# Paths (repo-relative, forward slashes; a trailing / means a directory) that define what the agent CAN do.
PROTECTED = ("mcp-server/", "Modules/", "ord-addin/", "qa/", "tests/", "Data/workflow-contracts/", "Data/capabilities/",
             "Data/diagnostic-ladders/", "Data/command-registry.tsv", "Data/keyin-candidates", "Data/recipe-candidates.tsv",
             "scripts/qa_", "scripts/hot_reload", ".claude/", ".cursor/")
# Commands that load, build or promote tooling into the running application.
LOADERS = re.compile(r"hot_reload\.py|\bbuild\.ps1\b|\bregister\.ps1\b|restart_chat_driver\.py|vba_compile_recovery\.py|"
                     r"promote_registry|registry[_-]promote|harvest_registry", re.I)
_WRITES = re.compile(r"(?:>>?|\bTee-Object\b|\bSet-Content\b|\bAdd-Content\b|\bOut-File\b|\bsed\s+-i\b|\bcp\b|\bmv\b|\bcopy\b|\bmove\b|\bcopy-item\b|\bmove-item\b|\bgit\s+(?:checkout|restore|apply|stash)\b)", re.I)


def _rel(path: str) -> str:
    p = str(path).replace("\\", "/")
    root = str(ROOT).replace("\\", "/").rstrip("/") + "/"
    if p.lower().startswith(root.lower()):
        p = p[len(root):]
    return p.lstrip("./") if p.startswith("./") else p


def is_protected(path: str) -> bool:
    r = _rel(path)
    return any(r.startswith(pre) or r.lower().startswith(pre.lower()) for pre in PROTECTED)


def task_mode(task: dict) -> str:
    return task.get("mode") or "production"


def open_tasks() -> list[dict]:
    out = []
    if not TASKS.is_dir():
        return out
    now = time.time()
    for d in TASKS.iterdir():
        p = d / "task.json"
        if not p.exists():
            continue
        try:
            t = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        touched = max(t.get("created", 0), t.get("last_run", 0) or 0, t.get("reopened", 0) or 0)
        if t.get("status") == "in_progress" and now - touched <= WINDOW_S:
            out.append(t)
    return out


def production_tasks() -> list[dict]:
    return [t for t in open_tasks() if task_mode(t) == "production"]


def may_edit_drawing(task: dict) -> bool:
    """A development task with --scratch none is code-only: it never edits a drawing."""
    return not (task_mode(task) == "development" and str(task.get("scratch", "")).lower() == "none")


def _refusal(what: str, tasks: list[dict]) -> str:
    ids = ", ".join(t["id"] for t in tasks[:3])
    return (f"{what} is capability development, and production task {ids} is open. A production task changes the engineer's drawing "
            "with established capabilities only; it does not rebuild the tooling. Either finish it (`qa_gate close`, or `qa_gate escalate` "
            "and tell the engineer what is unsupported), or the engineer starts a development task on a scratch file: "
            "`python scripts/qa_gate.py start --mode development --scratch <file.dgn|none> --mode-reason \"...\" --type generic --desc \"...\"`.")


def assess_path(path: str, tasks: list[dict] | None = None) -> str:
    if not is_protected(path):
        return ""
    tasks = production_tasks() if tasks is None else tasks
    return _refusal(f"Editing {_rel(path)}", tasks) if tasks else ""


def assess_command(command: str, tasks: list[dict] | None = None) -> str:
    """Refuse a shell command that loads/builds tooling, or writes to a protected path, while a production task is open."""
    tasks = production_tasks() if tasks is None else tasks
    if not tasks:
        return ""
    m = LOADERS.search(command)
    if m:
        return _refusal(f"`{m.group(0)}`", tasks)
    if _WRITES.search(command):
        for tok in re.findall(r"[\w./\\:~-]+", command):
            if ("/" in tok or "\\" in tok or "." in tok) and is_protected(tok):
                return _refusal(f"Writing {_rel(tok)}", tasks)
    return ""
