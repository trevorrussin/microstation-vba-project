"""Stop hook: block a final message whose claims the turn's tool calls do not back.

Claude Code (.claude/settings.json "Stop"): reads hook JSON on stdin (transcript_path, stop_hook_active); prints
  {"decision": "block", "reason": "..."}  -> Claude keeps working (run the check, or say it was not run)
Cursor (.cursor/hooks.json "stop", run with --cursor): reads the same JSON (transcript_path, loop_count); prints
  {"followup_message": "..."}  -> Cursor sends it as the next message. Cursor transcripts keep no tool results,
  so only the "no tool call" and hedging checks run there.
Blocks at most once per turn (stop_hook_active / loop_count).

Checks (deterministic, on the final assistant message of the turn; quoted text in backticks or blockquotes is ignored):
  1. "I checked / I verified / confirmed by" with NO tool call in the turn.
  2. "I think / I believe / probably" about state (file, model, drawing, build, test ...).
  3. (Claude Code) "N passed" / "tests pass" with no tool result in the turn containing "passed".
  4. (Claude Code) "<CHECK-ID> PASS" (a QA gate check id) with no tool result in the turn containing that id followed by PASS.
It compares what the message says with what the tools returned; it does not judge whether the evidence is the right evidence.
"""
from __future__ import annotations

import json
import re
import sys

CHECK_CLAIM = re.compile(r"\b(i checked|i have checked|i verified|i confirmed|i know\.\s*i checked|verified by|confirmed by)\b", re.I)
HEDGE = re.compile(r"\b(i think|i believe|i assume|probably|presumably)\b", re.I)
TESTS_CLAIM = re.compile(r"\b(\d+\s+(?:tests?\s+)?passed|tests?\s+pass(?:es|ed)?|all\s+tests\s+pass)\b", re.I)
CHECK_PASS_CLAIM = re.compile(r"\b([A-Z]{2,5}(?:-[A-Z0-9]{1,5})?-\d{2})\b[^.\n]{0,40}?\bPASS(?:ES|ED)?\b")


def _strip_quotes(text: str) -> str:
    text = re.sub(r"```.*?```", " ", text, flags=re.S)
    text = re.sub(r"`[^`]*`", " ", text)
    return "\n".join(l for l in text.splitlines() if not l.lstrip().startswith(">"))


def _result_text(block: dict) -> str:
    c = block.get("content")
    if isinstance(c, str):
        return c
    if isinstance(c, list):
        return " ".join(str(b.get("text", "")) for b in c if isinstance(b, dict))
    return ""


def last_turn(lines: list[str]) -> tuple[str, int, str]:
    """-> (text of the final assistant message, number of tool_use blocks, joined tool_result text) for the latest turn.

    Claude Code lines: {"type": "assistant", "message": {"role": "assistant", "content": [...]}}.
    Cursor lines:      {"role": "assistant", "message": {"content": [...]}} (no tool results)."""
    entries = []
    for ln in lines:
        try:
            entries.append(json.loads(ln))
        except Exception:
            continue
    tools, text, results = 0, "", []
    for e in reversed(entries):
        msg = e.get("message") or {}
        content = msg.get("content")
        role = msg.get("role") or e.get("role") or e.get("type")
        if role == "user":
            blocks = content if isinstance(content, list) else [{"type": "text", "text": str(content or "")}]
            if any(b.get("type") == "tool_result" for b in blocks if isinstance(b, dict)):
                results += [_result_text(b) for b in blocks if isinstance(b, dict) and b.get("type") == "tool_result"]
                continue                                   # a tool result, same turn
            break                                          # the human message that started the turn
        if role == "assistant" and isinstance(content, list):
            for b in content:
                if b.get("type") == "tool_use":
                    tools += 1
                elif b.get("type") == "text" and not text:
                    text = b.get("text", "")
    return text, tools, "\n".join(results)


def verdict(text: str, tools: int, results: str | None = None) -> str:
    """-> "" when fine, else the reason to block. `results` None = tool results unavailable (Cursor): checks 3 and 4 are skipped."""
    body = _strip_quotes(text)
    if tools == 0 and CHECK_CLAIM.search(body):
        return ("Your message says a check was run, but this turn made no tool call. Run the check now, or say plainly "
                "that it was not run. Do not claim evidence you did not collect this turn.")
    if HEDGE.search(body) and re.search(r"\b(file|model|drawing|view|sheet|level|element|compile|test|build|exists?)\b", body, re.I):
        return ("Your message hedges (\"I think\" / \"probably\") about state. Check it, then state `I know. I checked: X, "
                "result: Y`, or label it as an inference or unknown.")
    if results is not None:
        if TESTS_CLAIM.search(body) and not re.search(r"\bpassed\b", results, re.I):
            return ("Your message says tests passed, but no tool result in this turn shows a test run (\"passed\"). "
                    "Run the tests now, or say they were not run.")
        for m in CHECK_PASS_CLAIM.finditer(body):
            cid = m.group(1)
            if not re.search(re.escape(cid) + r"\s+PASS", results):
                return (f"Your message says {cid} passed, but no tool result in this turn shows \"{cid} PASS\". "
                        "Run the gate now and report what it printed, or say it was not run.")
    return ""


def main() -> int:
    cursor = "--cursor" in sys.argv
    try:
        data = json.load(sys.stdin)
    except Exception:
        return 0
    if data.get("stop_hook_active") or (cursor and int(data.get("loop_count") or 0) >= 1):
        return 0
    path = data.get("transcript_path")
    if not path:
        return 0
    try:
        with open(path, encoding="utf-8") as f:
            text, tools, results = last_turn(f.read().splitlines())
    except Exception:
        return 0
    why = verdict(text, tools, None if cursor else results)
    if why:
        print(json.dumps({"followup_message": why} if cursor else {"decision": "block", "reason": why}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
