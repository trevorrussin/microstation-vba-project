"""Loads every QA task with its event log and attempt rows (shared by qa/trajectory.py and its report script)."""
from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))


def collect(tasks_dir: str | None = None) -> list[dict]:
    import task_guard
    from qa import trajectory
    from qa.checks import process
    tdir = tasks_dir or str(task_guard.TASKS)
    rows = []
    if not os.path.isdir(tdir):
        return rows
    for d in sorted(os.listdir(tdir)):
        p = os.path.join(tdir, d, "task.json")
        if not os.path.exists(p):
            continue
        try:
            t = json.load(open(p, encoding="utf-8"))
            ev = process.events_for_task(t)
            rows.append(trajectory.task_metrics(t, ev, task_guard.attempt_rows(t["id"])))
        except Exception:
            continue
    return rows
