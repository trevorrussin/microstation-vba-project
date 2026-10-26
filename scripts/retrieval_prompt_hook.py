"""UserPromptSubmit hook: if the prompt matches tutorial playbooks that have no fresh read receipt, tell the agent to
read them before answering and arm the Stop hook (mcp-server/retrieval.py) to hold the turn until it has.

Claude Code (.claude/settings.json "UserPromptSubmit"): reads {"prompt": ...} on stdin, prints additionalContext.
"""
from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))


def ladder_context(prompt: str) -> str:
    """Diagnostic ladder + past experience for a symptom the engineer's own words describe (pasted material ignored)."""
    import experience
    import ladders
    import retrieval
    own = retrieval.prompt_text_for_matching(prompt)
    lad = ladders.match(own)
    if not lad:
        return ""
    txt = ladders.summary(lad)
    past = experience.for_task(own)
    return txt + (chr(10) + past if past else "")


def main() -> int:
    try:
        prompt = json.load(sys.stdin).get("prompt", "")
        import retrieval
        lad_ctx = ladder_context(prompt)
        matched = retrieval.hook_matches(prompt)
        unread = [p for p in matched if not retrieval.has_fresh_receipt(p["topic"])]
        topics = [p["topic"] for p in unread]
        if retrieval.already_told(topics):  # same demand, already given and given up on: do not repeat it
            unread = []
        else:
            retrieval.set_pending(topics, need_search=bool(unread))  # nothing unread -> nothing to wait for
    except Exception:
        return 0
    if lad_ctx and not unread:
        print(json.dumps({"hookSpecificOutput": {"hookEventName": "UserPromptSubmit", "additionalContext": lad_ctx}}))
    if unread:
        names = ", ".join(p["topic"] for p in unread) or "(all already read)"
        ctx = (f"RETRIEVE BEFORE ANSWERING: this request matches playbooks: {', '.join(p['topic'] for p in matched)}. "
               f"Unread: {names}. Call get_tutorial_playbook for each unread one (and any it points to). Then search_standards "
               "for the governing requirement and search_tutorials / read_tutorial for the written procedure the playbook "
               "cites (the NYSDOT PDF, decks, ORD manual). Establish the problem, the steps and the standards from those, "
               "cite step ids and sections, and only then answer or edit. If retrieval fails, say so." + (chr(10) * 2 + lad_ctx if lad_ctx else ""))
        print(json.dumps({"hookSpecificOutput": {"hookEventName": "UserPromptSubmit", "additionalContext": ctx}}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
