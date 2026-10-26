"""Diagnostic ladders: for a known symptom, the checks in the order to run them (cheap and read-only first).

Data lives in Data/diagnostic-ladders/*.yaml. A task whose description matches a ladder carries `task["ladder"]`;
each rung gets a recorded result (`qa_gate rung`). Rungs whose risk is not `edit` gate the `edit` rung: the bridge guard
(task_guard) and PROC-01 refuse scripted edits while an earlier rung has no result. The rung result is the agent's
statement, so it needs evidence text; the gate cannot prove the check was real, but it makes skipping the order visible.

Rung statuses:
  negative    checked, this is not the cause
  positive    checked, this is the cause (rungs after it are not needed)
  handed_off  a GUI step (who: engineer) was handed to the engineer with the exact instruction; counts as dispositioned
  na          does not apply here (reason in the evidence)
"""
from __future__ import annotations

import glob
import os
import re
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LADDERS = os.path.join(ROOT, "Data", "diagnostic-ladders")
STATUSES = ("negative", "positive", "handed_off", "na")
GATED_RISK = "edit"  # a rung with this risk is held back until the others have a result


def _yaml(path):
    import yaml
    return yaml.safe_load(open(path, encoding="utf-8"))["ladder"]


def load_all() -> list[dict]:
    out = []
    for p in sorted(glob.glob(os.path.join(LADDERS, "*.yaml"))):
        try:
            out.append(_yaml(p))
        except Exception:
            continue
    return out


def get(ladder_id: str) -> dict | None:
    return next((l for l in load_all() if l["id"] == ladder_id), None)


def match(text: str) -> dict | None:
    """Best ladder whose keywords hit the text at least `min_hits` times (distinct keywords)."""
    low = (text or "").lower()
    best, best_hits = None, 0
    for l in load_all():
        hits = sum(1 for k in l.get("keywords", []) if str(k).lower() in low)
        if hits >= int(l.get("min_hits", 2)) and hits > best_hits:
            best, best_hits = l, hits
    return best


def attach(task: dict, ladder: dict) -> None:
    task["ladder"] = {"id": ladder["id"], "rungs": {}}


def rungs_of(task: dict) -> list[dict]:
    l = get((task.get("ladder") or {}).get("id", ""))
    return list(l["rungs"]) if l else []


def state(task: dict) -> dict:
    return (task.get("ladder") or {}).get("rungs", {})


def gating_rungs(task: dict) -> list[dict]:
    return [r for r in rungs_of(task) if r.get("risk") != GATED_RISK]


def gaps(task: dict) -> list[str]:
    """What must have a recorded result before a scripted edit: every non-edit rung up to the first positive one."""
    if not task.get("ladder"):
        return []
    st, out = state(task), []
    for r in gating_rungs(task):
        s = st.get(r["id"])
        if s and s["status"] == "positive":
            break
        if not s:
            who = "ask the engineer" if r.get("who") == "engineer" else "check"
            out.append(f"ladder {r['id']} ({who}): {r['title']}")
    return out


def record(task: dict, rung_id: str, status: str, evidence: str) -> dict:
    if not task.get("ladder"):
        raise ValueError("this task has no diagnostic ladder")
    ids = [r["id"] for r in rungs_of(task)]
    if rung_id not in ids:
        raise ValueError(f"no rung {rung_id}; rungs: {', '.join(ids)}")
    if status not in STATUSES:
        raise ValueError(f"status must be one of {', '.join(STATUSES)}")
    if len((evidence or "").strip()) < 15:
        raise ValueError("--evidence must say what was checked and what came back (>= 15 characters)")
    rung = next(r for r in rungs_of(task) if r["id"] == rung_id)
    if status == "handed_off" and rung.get("who") != "engineer":
        raise ValueError(f"{rung_id} is an agent step (who: agent): check it, do not hand it off")
    row = {"status": status, "evidence": evidence.strip(), "at": time.time()}
    task["ladder"]["rungs"][rung_id] = row
    return row


def summary(ladder: dict) -> str:
    lines = [f"DIAGNOSTIC LADDER {ladder['id']}: {ladder['title']}",
             "  Work the rungs in order. Record each result: qa_gate rung --id R# --status negative|positive|handed_off|na --evidence \"...\"",
             "  Scripted edits (the last rung) are refused until every earlier rung has a result."]
    for r in ladder["rungs"]:
        lines.append(f"  {r['id']} [{r.get('who')}, {r.get('risk')}] {r['title']}")
        how = re.sub(r"\s+", " ", str(r.get("how", "")).strip())
        if how:
            lines.append(f"      {how[:300]}")
    return "\n".join(lines)


def events(task: dict) -> list[dict]:
    """Ladder events for PROC-01: one `ladder_open` (with the count of gating rungs) and one `rung` per recorded result."""
    if not task.get("ladder"):
        return []
    ev = [{"kind": "ladder_open", "at": task.get("created", 0),
           "gating": [r["id"] for r in gating_rungs(task)]}]
    for rid, s in state(task).items():
        ev.append({"kind": "rung", "at": s["at"], "rung": rid, "status": s["status"]})
    return ev
