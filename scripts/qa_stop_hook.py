"""Stop hook: do not let the agent end its turn while an open QA task is not PASS.

Claude Code (.claude/settings.json "Stop"): reads hook JSON on stdin; prints
  {"decision": "block", "reason": "..."}  -> Claude keeps working (fix + re-run qa_gate)
Cursor (.cursor/hooks.json "stop", run with --cursor): prints {"followup_message": "..."}.

Only tasks touched in the last 12 h are enforced (old abandoned tasks only warn).
Loop guard: after 6 blocks without a new `qa_gate run`, the hook stops blocking and tells the
engineer the task is stuck (so a broken environment cannot trap the session).
"""
from __future__ import annotations

import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TASKS = os.path.join(ROOT, "outputs", "qa", "tasks")
WINDOW_S = 12 * 3600
MAX_BLOCKS = 6


def main() -> int:
    cursor = "--cursor" in sys.argv
    try:
        json.load(sys.stdin)
    except Exception:
        pass
    if not os.path.isdir(TASKS):
        return 0
    now = time.time()
    msgs, stuck = [], []
    for d in sorted(os.listdir(TASKS)):
        p = os.path.join(TASKS, d, "task.json")
        if not os.path.exists(p):
            continue
        try:
            t = json.load(open(p, encoding="utf-8"))
        except Exception:
            continue
        if t.get("status") != "in_progress":
            continue
        touched = max(t.get("created", 0), t.get("last_run", 0) or 0)
        if now - touched > WINDOW_S:
            continue
        if t.get("last_status") == "PASS":
            msgs.append(f"QA task {t['id']} PASSED its last run but is still open: run "
                        f"`python scripts/qa_gate.py close --task {t['id']}` and report the evidence.")
            continue
        last_run = t.get("last_run") or 0
        if t.get("_hook_last_run") != last_run:
            t["_hook_last_run"], t["_hook_blocks"] = last_run, 0
        t["_hook_blocks"] = t.get("_hook_blocks", 0) + 1
        json.dump(t, open(p, "w", encoding="utf-8"), indent=1)
        if t["_hook_blocks"] > MAX_BLOCKS:
            stuck.append(t["id"])
            continue
        latest = os.path.join(TASKS, d, "latest.json")
        detail = ""
        if os.path.exists(latest):
            rec = json.load(open(latest, encoding="utf-8"))
            bad = [f"{r['id']} {r['status']}: {r['observed']}" for r in rec.get("results", [])
                   if r.get("blocking", True) and r["status"] in ("FAIL", "UNCERTAIN")]
            detail = " Open rows: " + " | ".join(bad[:6])
        state = t.get("last_status") or "never run"
        msgs.append(f"QA task {t['id']} ({t['type']}) is {state}.{detail} Fix the problems, re-run "
                    f"`python scripts/qa_gate.py run --task {t['id']}`, and only finish when it PASSES "
                    f"(then `qa_gate close`). If it cannot be fixed from evidence, run `qa_gate escalate --task "
                    f"{t['id']} --reason ...` and tell the engineer exactly what is blocked.")
    if stuck:
        note = ("QA gate stuck (no new run after repeated stop attempts): " + ", ".join(stuck) +
                ". Tell the engineer these tasks are NOT verified.")
        print(json.dumps({"systemMessage": note} if not cursor else {"followup_message": note}))
        return 0
    if msgs:
        reason = "\n".join(msgs)
        print(json.dumps({"followup_message": reason} if cursor else {"decision": "block", "reason": reason}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
