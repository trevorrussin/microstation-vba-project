"""Experience memory: what the agent actually did on past tasks, including what failed and what stayed unresolved.

Knowledge RAG (standards, playbooks, tutorials) says what is correct. This file says what happened when it was tried here.
Rows are appended to Data/experience/experience.jsonl (tracked in git, so Cursor and Claude Code share it).

Rules that keep it honest:
  - `status` is resolved | unresolved | handed_off. An unresolved investigation is never stored as a solution.
  - Each finding carries a level: observed (read back or tool error text) | reported (taken from a log or the engineer, not re-tested)
    | inferred. Findings are about one context (file, model, version); they are hints, not rules.
  - `qa_gate escalate` and `qa_gate close` write a row automatically from the task; `qa_gate experience add` writes one by hand.
  - `qa_gate start` prints the closest rows for the task description.
"""
from __future__ import annotations

import json
import math
import os
import re
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH = os.path.join(ROOT, "Data", "experience", "experience.jsonl")
STATUSES = ("resolved", "unresolved", "handed_off")
LEVELS = ("observed", "reported", "inferred")
_STOP = set("the and for with that this from have has not are was were but all any into onto over under when then than "
            "they them their will can could should would about after before while which what where there here".split())


def _tokens(text: str) -> set[str]:
    return {w for w in re.findall(r"[a-z0-9_]{4,}", (text or "").lower()) if w not in _STOP}


def rows() -> list[dict]:
    out = []
    if os.path.exists(PATH):
        for ln in open(PATH, encoding="utf-8"):
            try:
                out.append(json.loads(ln))
            except Exception:
                continue
    return out


def add(symptom: str, status: str, tried: list[dict], findings: list[dict], context: str = "", task: str = "", source: str = "manual") -> dict:
    if status not in STATUSES:
        raise ValueError(f"status must be one of {', '.join(STATUSES)}")
    if len((symptom or "").strip()) < 15:
        raise ValueError("symptom must describe the problem in a sentence (>= 15 characters)")
    for f in findings:
        if f.get("level") not in LEVELS:
            raise ValueError(f"finding level must be one of {', '.join(LEVELS)}: {f}")
    row = {"id": max([r.get("id", 0) for r in rows()] or [0]) + 1, "at": time.time(), "date": time.strftime("%Y-%m-%d"),
           "symptom": symptom.strip(), "context": context.strip(), "status": status, "tried": tried, "findings": findings,
           "task": task, "source": source}
    os.makedirs(os.path.dirname(PATH), exist_ok=True)
    with open(PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")
    return row


def search(text: str, k: int = 3, min_shared: int = 2) -> list[dict]:
    q = _tokens(text)
    if not q:
        return []
    scored = []
    for r in rows():
        d = _tokens(r["symptom"] + " " + r.get("context", "") + " " + " ".join(t.get("what", "") for t in r.get("tried", [])))
        shared = q & d
        if len(shared) >= min_shared:
            scored.append((len(shared) / math.sqrt(len(d) or 1), r))
    return [r for _, r in sorted(scored, key=lambda x: -x[0])[:k]]


def format_row(r: dict) -> str:
    lines = [f"  #{r['id']} {r['date']} [{r['status'].upper()}] {r['symptom']}" + (f"  (context: {r['context']})" if r.get("context") else "")]
    for t in r.get("tried", []):
        lines.append(f"      tried: {t.get('what', '')} -> {t.get('result', '')}")
    for f in r.get("findings", []):
        lines.append(f"      {f['level']}: {f['text']}")
    return "\n".join(lines)


def for_task(desc: str, k: int = 3) -> str:
    hits = search(desc, k)
    if not hits:
        return ""
    return ("PAST EXPERIENCE on similar problems (hints for this context only; an UNRESOLVED row is not a solution; "
            "do not repeat an approach that failed without a reason it would differ):\n" + "\n".join(format_row(r) for r in hits))


def from_task(task: dict) -> dict | None:
    """Row for a task that ended: escalated -> unresolved/handed_off, closed -> resolved. Built from the task record only:
    ladder rung results become `tried`, observed/inferred claims become findings, the escalation text is `reported`."""
    status = task.get("status")
    if status not in ("escalated", "closed"):
        return None
    tried = [{"what": f"ladder {rid}: {s['evidence'][:120]}", "result": s["status"]} for rid, s in ((task.get("ladder") or {}).get("rungs", {})).items()]
    findings = [{"level": "observed" if c["level"] in ("observed", "verified") else "inferred", "text": c["text"]}
                for c in task.get("claims", []) if c["level"] in ("observed", "inferred", "verified")]
    if status == "escalated" and task.get("escalation"):
        findings.append({"level": "reported", "text": task["escalation"]})
    handed = any(s["status"] == "handed_off" for s in (task.get("ladder") or {}).get("rungs", {}).values())
    st = "resolved" if status == "closed" else ("handed_off" if handed else "unresolved")
    return {"symptom": task.get("desc", ""), "status": st, "tried": tried, "findings": findings,
            "context": "active model " + str(task.get("model") or "unknown"), "task": task.get("id", "")}
