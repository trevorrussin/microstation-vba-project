"""Tool-boundary enforcement for the QA gate: the bridge refuses drawing edits unless a QA task is open.

The stop hook is a last line of defence that only sees tasks that exist. This guard sits where edits enter
MicroStation (bridge_client.Bridge.call_batch): no open `cad.drafting` task = the edit is refused, so the agent
cannot forget `qa_gate start`. Reads (GET_/LIST_/FIND_/DESCRIBE_/...) are always allowed.

Open = status `in_progress` and touched within WINDOW_S. A task the stop hook gave up on is `blocked`, an old one is
`expired`; neither counts (`qa_gate reopen --task ID` makes it open again, or start a new task).

Failed attempts: an edit that raises or comes back not-OK is logged (outputs/qa/attempts.jsonl). After
MAX_FAILED_ATTEMPTS in a row on the open task the guard refuses the next edit (ReplanRequired) until the agent has
consulted a source since the last failure and records `qa_gate replan --task ID --note "..."`, or escalates. A success
resets the count. This replaces open-ended trial and error with: try, try once more, read, change strategy or hand off.

Escape hatch for maintenance and scratch scripts that are not drawing work: set QA_ALLOW_NO_TASK=1 in that
process (scripts/playbook_runner.py does; its elements are scratch and removed).
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path

TASKS = Path(__file__).resolve().parents[1] / "outputs" / "qa" / "tasks"
WINDOW_S = 12 * 3600
READ_PREFIXES = ("GET_", "LIST_", "FIND_", "DESCRIBE_", "RESOLVE_", "COMPUTE_", "CHECK_", "SEARCH_", "CROSS_VALIDATE")
READ_OPS = {"STATION_TO_POINT", "CLEAR_SELECTION", "SELECT_ELEMENT", "SHOW_SHEET_VIEWER", "HANDOFF"}  # no drawing change
ATTEMPTS = TASKS.parent / "attempts.jsonl"
MAX_FAILED_ATTEMPTS = 2
_cache: dict = {"at": 0.0, "open": None}


class QAGateRequired(RuntimeError):
    """Raised instead of sending a drawing edit when no QA task is open."""


class PlaybookNotRead(QAGateRequired):
    """Raised instead of sending a drawing edit while every open task has a required playbook with no fresh receipt."""


class ReplanRequired(QAGateRequired):
    """Raised instead of sending a drawing edit after MAX_FAILED_ATTEMPTS failed edits in a row with no replan."""


def is_edit(op: str) -> bool:
    return not (op.startswith(READ_PREFIXES) or op in READ_OPS)


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
        if t.get("status") == "in_progress" and t.get("type") == "cad.drafting" and now - touched <= WINDOW_S:
            out.append(t)
    return out


def require_open_task(ops: list[str]) -> None:
    if os.environ.get("QA_ALLOW_NO_TASK") == "1":
        return
    edits = sorted({o for o in ops if is_edit(o)})
    if not edits:
        return
    if time.time() - _cache["at"] > 2.0:  # one directory scan per couple of seconds, not per op
        _cache.update(at=time.time(), open=open_tasks())
    if _cache["open"]:
        import modes
        if not any(modes.may_edit_drawing(t) for t in _cache["open"]):
            raise QAGateRequired(
                f"refused {', '.join(edits[:4])}: the open task is a code-only development task (--scratch none); it never edits a drawing. "
                "Start a task that names a scratch file (`qa_gate start --mode development --scratch <file.dgn>`) or a production task for real work.")
        import ladders
        import retrieval
        gaps = [retrieval.gaps(t) + ladders.gaps(t) for t in _cache["open"]]
        if all(gaps):  # an open task with nothing left to read or disposition lets the edit through
            todo = sorted({g for gl in gaps for g in gl})
            raise PlaybookNotRead(
                f"refused {', '.join(edits[:4])}: research not finished for the open task. Still to do: "
                + "; ".join(todo[:10]) + (f" (+{len(todo) - 10} more)" if len(todo) > 10 else "") +
                ". Read with get_tutorial_playbook / get_standard_section / read_standard_pages / read_tutorial, then "
                "`qa_gate dispose --cite \"<cite>\" --applies yes|no --reason \"...\"`. `yes` needs the read; `no` needs only the reason. "
                "Ladder rungs: check (or ask the engineer), then `qa_gate rung --id R# --status negative|positive|handed_off|na --evidence \"...\"`.")
    if _cache["open"]:
        check_failure_budget(_cache["open"], edits)
    if not _cache["open"]:
        raise QAGateRequired(
            f"refused {', '.join(edits[:4])}: no open QA task. Start one first: "
            "`python scripts/qa_gate.py start --type cad.drafting --desc \"...\" --extent ...` "
            "(an expired or blocked task does not count: `qa_gate reopen --task ID`). "
            "Scratch/maintenance scripts that are not drawing work set QA_ALLOW_NO_TASK=1.")


# ------------------------------------------------------------------ failed-attempt budget
def _log(row: dict) -> None:
    ATTEMPTS.parent.mkdir(parents=True, exist_ok=True)
    with open(ATTEMPTS, "a", encoding="utf-8") as f:
        f.write(json.dumps(row) + "\n")


def attempt_rows(task_id: str | None = None) -> list[dict]:
    rows = []
    if ATTEMPTS.exists():
        for ln in ATTEMPTS.read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(ln)
            except Exception:
                continue
            if task_id is None or r.get("task") == task_id:
                rows.append(r)
    return rows


def record_attempts(ops: list[str], outcome) -> None:
    """Log each edit op of a batch against every open task. `outcome` is the result list, or the exception raised.
    Best effort: logging must never change what the bridge returns."""
    try:
        edits = [(i, o) for i, o in enumerate(ops) if is_edit(o)]
        if not edits:
            return
        tasks = open_tasks()
        for i, op in edits:
            if isinstance(outcome, BaseException):
                ok, note = False, str(outcome)[:200]
            else:
                r = outcome[i] if i < len(outcome) else {}
                ok = str(r.get("status", "")).upper() == "OK"
                note = "" if ok else str(r.get("note") or r.get("status", ""))[:200]
            for t in tasks:
                _log({"task": t["id"], "kind": "attempt_ok" if ok else "attempt_fail", "op": op, "note": note, "at": time.time()})
    except Exception:
        pass


def record_replan(task_id: str, note: str) -> None:
    _log({"task": task_id, "kind": "replan", "note": note, "at": time.time()})


def failed_streak(task_id: str) -> int:
    """Failed edits since the last success or replan."""
    n = 0
    for r in attempt_rows(task_id):
        if r["kind"] == "attempt_fail":
            n += 1
        elif r["kind"] in ("attempt_ok", "replan"):
            n = 0
    return n


def last_failure_at(task_id: str) -> float:
    return max([r["at"] for r in attempt_rows(task_id) if r["kind"] == "attempt_fail"] or [0.0])


def check_failure_budget(tasks: list[dict], edits: list[str]) -> None:
    if tasks and all(failed_streak(t["id"]) >= MAX_FAILED_ATTEMPTS for t in tasks):
        t = tasks[0]
        raise ReplanRequired(
            f"refused {', '.join(edits[:4])}: {failed_streak(t['id'])} edits in a row failed on task {t['id']}. Stop trying variants. "
            "Read what governs the operation (get_tutorial_playbook / search_tutorials / read_tutorial / search_standards), then "
            f"`python scripts/qa_gate.py replan --task {t['id']} --note \"what failed, what the source says, the different approach\"`, "
            "or `qa_gate escalate --reason ...` and hand the step to the engineer (the playbook marks GUI steps who: engineer).")
