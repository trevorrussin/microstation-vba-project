"""Stop hook: do not let the agent end its turn while an open QA task is not PASS.

Claude Code (.claude/settings.json "Stop"): reads hook JSON on stdin; prints
  {"decision": "block", "reason": "..."}  -> Claude keeps working (fix + re-run qa_gate)
Cursor (.cursor/hooks.json "stop", run with --cursor): prints {"followup_message": "..."}.

Only tasks touched in the last 12 h are enforced. An older task becomes `expired`, and a task the hook gave up on
after 6 blocks without a new `qa_gate run` becomes `blocked` (so a broken environment cannot trap the session).
Both are written to task.json, reported to the engineer as NOT VERIFIED, no longer count as open for the tool-boundary
guard (mcp-server/task_guard.py), and are resumed with `qa_gate reopen --task ID`.
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
    msgs, stuck, notes = [], [], []
    try:
        sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
        import retrieval
        rm = retrieval.stop_message()
        if rm:
            msgs.append(rm)
    except Exception:
        pass
    if not os.path.isdir(TASKS):
        if msgs:
            print(json.dumps({"followup_message": msgs[0]} if cursor else {"decision": "block", "reason": msgs[0]}))
        return 0
    now = time.time()
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
        touched = max(t.get("created", 0), t.get("last_run", 0) or 0, t.get("reopened", 0) or 0)
        if now - touched > WINDOW_S:
            t["status"], t["state_reason"] = "expired", f"no activity for {(now - touched) / 3600:.0f} h"
            json.dump(t, open(p, "w", encoding="utf-8"), indent=1)
            notes.append(f"QA task {t['id']} EXPIRED ({t['state_reason']}): its work is NOT verified. "
                         f"`qa_gate reopen --task {t['id']}` to continue it, or start a new task.")
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
            t["status"], t["state_reason"] = "blocked", f"{MAX_BLOCKS} stop attempts without a new qa_gate run"
            json.dump(t, open(p, "w", encoding="utf-8"), indent=1)
            stuck.append(t["id"])
            continue
        latest = os.path.join(TASKS, d, "latest.json")
        detail = ""
        if os.path.exists(latest):
            rec = json.load(open(latest, encoding="utf-8"))
            bad = [f"{r['id']} {r['status']}: {r['observed']}" for r in rec.get("results", [])
                   if r.get("blocking", True) and r["status"] in ("FAIL", "UNCERTAIN")]
            detail = " Open rows: " + " | ".join(bad[:6])
            need = rec.get("needs_signoff") or []
            if need:
                detail += (" ENGINEER SIGN-OFF needed for: " + ", ".join(n["id"] for n in need) + ". The agent cannot sign off: tell the engineer to review "
                           "and run `python scripts/qa_gate.py signoff --task " + t["id"] + "` in their own terminal.")
        state = t.get("last_status") or "never run"
        msgs.append(f"QA task {t['id']} ({t['type']}) is {state}.{detail} Fix the problems, re-run "
                    f"`python scripts/qa_gate.py run --task {t['id']}`, and only finish when it PASSES "
                    f"(then `qa_gate close`). If it cannot be fixed from evidence, run `qa_gate escalate --task "
                    f"{t['id']} --reason ...` and tell the engineer exactly what is blocked.")
    if stuck:
        note = ("QA gate BLOCKED (no new run after repeated stop attempts): " + ", ".join(stuck) +
                ". Tell the engineer these tasks are NOT verified; they no longer count as open (`qa_gate reopen --task ID`).")
        note = " ".join([note] + notes)
        print(json.dumps({"systemMessage": note} if not cursor else {"followup_message": note}))
        return 0
    if notes and not msgs:
        print(json.dumps({"systemMessage": " ".join(notes)} if not cursor else {"followup_message": " ".join(notes)}))
        return 0
    if msgs:
        reason = "\n".join(msgs + notes)
        print(json.dumps({"followup_message": reason} if cursor else {"decision": "block", "reason": reason}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
