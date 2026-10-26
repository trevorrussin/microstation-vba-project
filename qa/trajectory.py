"""Trajectory quality: how the work went, not only whether the final drawing passed.

Built from the files the guards already write (task.json, attempts.jsonl, receipts.jsonl). No new logging.
Per task: time to the first read and first edit, edits ok/failed, the longest failed streak, redundant failures (the same op
failing with the same error again), replans, ladder rungs with a result vs rungs a scripted edit skipped, and the PROC-01
violations. Across tasks: verified-success rate (closed after a passing run), escalation rate, redundant-failure rate,
process-violation rate. Safety and completion are reported together: a gate that blocks everything has no false passes
and no completions.
"""
from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))


def task_metrics(task: dict, events: list[dict], attempts: list[dict]) -> dict:
    from qa.checks import process
    created = task.get("created", 0)
    reads = [e["at"] for e in events if e["kind"] == "read"]
    edits = [a for a in attempts if a["kind"] in ("attempt_ok", "attempt_fail")]
    fails = [a for a in edits if a["kind"] == "attempt_fail"]
    seen, redundant = set(), 0
    for a in fails:
        key = (a.get("op"), (a.get("note") or "")[:80])
        redundant += key in seen
        seen.add(key)
    streak = best = 0
    for a in attempts:
        if a["kind"] == "attempt_fail":
            streak += 1
            best = max(best, streak)
        elif a["kind"] in ("attempt_ok", "replan"):
            streak = 0
    first_edit = min((a["at"] for a in edits), default=None)
    # receipts.jsonl is shared by every task, so a read is attributed to this task only up to its first edit; a task with no
    # logged edit (all tasks before attempts.jsonl existed) has no measurable read timing
    reads = [t for t in reads if first_edit is not None and t <= first_edit]
    first_read = min(reads, default=None)
    rungs = (task.get("ladder") or {}).get("rungs", {})
    viol = process.violations(events)
    return {"id": task["id"], "mode": task.get("mode") or "production", "status": task.get("status"), "last_status": task.get("last_status"),
            "secs_to_first_read": None if first_read is None else round(first_read - created, 1),
            "secs_to_first_edit": None if first_edit is None else round(first_edit - created, 1),
            "read_before_edit": None if first_edit is None else first_read is not None,
            "edits": len(edits), "failed_edits": len(fails), "redundant_failures": redundant, "longest_fail_streak": best,
            "replans": sum(1 for a in attempts if a["kind"] == "replan"),
            "ladder": bool(task.get("ladder")), "rungs_recorded": len(rungs),
            "violations": viol, "verified_success": task.get("status") == "closed" and task.get("last_status") == "PASS"}


def summarize(rows: list[dict]) -> dict:
    n = len(rows) or 1
    edits = sum(r["edits"] for r in rows) or 1
    timed = [r for r in rows if r["read_before_edit"] is not None]
    return {"tasks": len(rows), "tasks_with_logged_edits": len(timed),
            "verified_success_rate": sum(r["verified_success"] for r in rows) / n,
            "escalated_rate": sum(r["status"] == "escalated" for r in rows) / n,
            "read_before_edit_rate": (sum(r["read_before_edit"] for r in timed) / len(timed)) if timed else None,
            "failed_edit_rate": sum(r["failed_edits"] for r in rows) / edits,
            "redundant_failure_rate": sum(r["redundant_failures"] for r in rows) / edits,
            "tasks_with_process_violations": sum(bool(r["violations"]) for r in rows) / n}


def load_all(tasks_dir: str | None = None) -> list[dict]:
    import process_events
    return process_events.collect(tasks_dir)


def markdown(rows: list[dict], summ: dict) -> str:
    L = ["# QA trajectory report", "",
         f"{summ['tasks']} tasks. Verified success {summ['verified_success_rate']:.0%}; escalated {summ['escalated_rate']:.0%}; "
         f"read before first edit {'n/a' if summ['read_before_edit_rate'] is None else format(summ['read_before_edit_rate'], '.0%')} "
         f"(of {summ['tasks_with_logged_edits']} tasks with logged edits; older tasks predate attempts.jsonl); failed edits {summ['failed_edit_rate']:.0%} of edits; "
         f"redundant failures {summ['redundant_failure_rate']:.0%} of edits; tasks with process violations {summ['tasks_with_process_violations']:.0%}.", "",
         "Read with the corpus report (`scripts/qa_corpus_report.py`): false passes show the gate is safe, this shows the work is not wasteful or blocked.", "",
         "| task | mode | status | edits | failed | redundant | streak | replans | rungs | to 1st read s | to 1st edit s | violations |", "|---|---|---|---|---|---|---|---|---|---|---|---|"]
    for r in rows:
        L.append(f"| {r['id'][:46]} | {r['mode']} | {r['status']} | {r['edits']} | {r['failed_edits']} | {r['redundant_failures']} | {r['longest_fail_streak']} | "
                 f"{r['replans']} | {r['rungs_recorded'] if r['ladder'] else '-'} | {r['secs_to_first_read']} | {r['secs_to_first_edit']} | {'; '.join(v[:40] for v in r['violations']) or '-'} |")
    return "\n".join(L) + "\n"
