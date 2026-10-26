"""PROC-01: did the work follow the process, judged from the task's event log (not from what the agent says).

Events (ctx["events"], oldest first), each {"kind": ..., "at": float}:
  task_open | read | search | attempt_ok | attempt_fail | replan | escalate | verify_pass | outcome (value=pass|blocked|fail)

Rules (each a concrete way the 2026-10-24 "turn off the circles" run went wrong):
  P1  no edit before the task is open
  P2  no edit before a playbook/standard/tutorial was read (research first)
  P3  no more than MAX_FAILED_ATTEMPTS failed edits in a row without a replan or an escalation
  P4  a replan needs a read or search after the last failure (changing strategy needs new information)
  P5  outcome `blocked` needs an escalation with a reason; outcome `pass` needs a verifying run
  P6  a task with a diagnostic ladder (Data/diagnostic-ladders): no edit while a cheaper rung has no recorded result
      (a `positive` rung ends the diagnosis; events: ladder_open{gating:[ids]}, rung{rung,status})
"""
from __future__ import annotations

import json
import os
import sys

MAX_FAILED_ATTEMPTS = 2
EDITS = ("attempt_ok", "attempt_fail")
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def violations(events: list[dict]) -> list[str]:
    out, opened, read, streak = [], False, False, 0
    read_since_fail, escalated, verified = True, False, False
    gating, done, positive = [], set(), False
    for e in events:
        k = e.get("kind")
        if k == "task_open":
            opened = True
        elif k == "ladder_open":
            gating = list(e.get("gating", []))
        elif k == "rung":
            done.add(e.get("rung"))
            positive = positive or e.get("status") == "positive"
        elif k in ("read", "search"):
            read = read or k == "read"
            read_since_fail = True
        elif k in EDITS:
            if not opened:
                out.append("P1 edit before any QA task was open")
            elif not read:
                out.append("P2 edit before any playbook, standard or tutorial was read")
            if gating and not positive and not set(gating) <= done:
                out.append("P6 scripted edit before the cheaper diagnostic rungs had a result: " + ", ".join(i for i in gating if i not in done))
            if k == "attempt_fail":
                streak += 1
                read_since_fail = False
                if streak > MAX_FAILED_ATTEMPTS:
                    out.append(f"P3 more than {MAX_FAILED_ATTEMPTS} failed edits in a row with no replan or escalation")
            else:
                streak = 0
        elif k == "replan":
            if not read_since_fail:
                out.append("P4 replan with nothing read since the last failure")
            streak = 0
        elif k == "escalate":
            escalated = bool(str(e.get("reason", "")).strip())
            streak = 0
        elif k == "verify_pass":
            verified = True
        elif k == "outcome":
            if e.get("value") == "blocked" and not escalated:
                out.append("P5 outcome blocked without an escalation reason")
            if e.get("value") == "pass" and not verified:
                out.append("P5 outcome pass without a verifying run")
    return list(dict.fromkeys(out))


def discipline(task: dict, ctx: dict) -> dict:
    ev = ctx.get("events")
    if ev is None:
        return {"status": "NA", "observed": "no event log for this task", "expected": "", "evidence": [], "fix": ""}
    v = violations(ev)
    exp = "research first, bounded retries, honest outcome"
    if v:
        return {"status": "FAIL", "observed": f"{len(v)} process violation(s)", "expected": exp, "evidence": v,
                "fix": "Read the governing playbook, stop after two failed edits, replan from the source or hand the step to the engineer."}
    return {"status": "PASS", "observed": f"{len(ev)} events follow the process", "expected": exp, "evidence": [], "fix": ""}


def events_for_task(task: dict) -> list[dict]:
    """Event log of a live task from the files the guards already write (attempts.jsonl, receipts.jsonl, task.json)."""
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import task_guard
    ev = [{"kind": "task_open", "at": task.get("created", 0)}]
    rec = os.path.join(ROOT, "outputs", "qa", "receipts.jsonl")
    if os.path.exists(rec):
        for ln in open(rec, encoding="utf-8"):
            try:
                r = json.loads(ln)
            except Exception:
                continue
            if r.get("at", 0) >= task.get("created", 0) - 3600:  # research done just before `start` counts
                ev.append({"kind": "search" if r.get("kind") == "search" else "read", "at": r["at"]})
    for r in task_guard.attempt_rows(task["id"]):
        ev.append({"kind": r["kind"], "at": r["at"]})
    import ladders
    ev += ladders.events(task)
    return sorted(ev, key=lambda e: e["at"])
