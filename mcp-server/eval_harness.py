"""
Small fixed-scenario eval harness for the WZTC chat agent (chat_driver.py).

Why this exists: until now, "is the agent good?" was answered anecdotally --
a handful of prompts typed into the live panel, someone eyeballing the
result (see Claude Code memory project_agent_capability_test_results.md).
Debug/*.bas tests VBA units, not agent behavior, so there was no way to
detect a regression from a prompt/tool/model change. This is a first,
deliberately small step: a fixed set of prompts with automated checks over
the agent's actual tool-call trace and final answer -- not a claim that it
replaces an engineer's judgment on drawing correctness (it doesn't look at
the drawing at all; that's the "verification pass" this is a prerequisite
for, not a replacement).

Each scenario runs one real turn through chat_driver.run_turn -- REAL
Anthropic API calls (billed) and REAL WZTCBridge calls against whatever
MicroStation session/model is currently open. A scenario that asks the
agent to place something will actually place it, same as manual panel
testing does; this harness doesn't clean up after itself for the same
reason a manual test session doesn't. Run it against DELETE.dgn or another
disposable model, not a real project file.

chat_driver.LOG and chat_driver.INPUT are swapped for isolated stand-ins
before any scenario runs, so this never writes into the live chat panel's
Bridge/chat-log.tsv (WZTCChatTimer polls that file -- an eval run would
otherwise spam the live transcript if the panel happens to be open) and
never blocks on Bridge/chat-input.tsv if a scenario triggers ask_user
(nothing types a reply into the real panel during an unattended run).

Usage:
    python eval_harness.py                  # run every scenario
    python eval_harness.py --only sign_code_translation compute_spacing
    python eval_harness.py --list            # print scenarios, no API calls
    python eval_harness.py --report Bridge/eval-results.json

Requires the same setup as chat_driver.py (ANTHROPIC_API_KEY, MicroStation
open with the Test VBA project loaded and WZTCBridge polling active).
"""
from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import anthropic

import chat_driver

Check = Callable[["Trace"], tuple[bool, str]]


@dataclass
class Trace:
    tool_calls: list[tuple[str, dict]]
    final_text: str


@dataclass
class Scenario:
    name: str
    prompt: str
    checks: list[tuple[str, Check]]


@dataclass
class ScenarioResult:
    name: str
    passed: bool
    check_results: list[tuple[str, bool, str]]
    final_text: str
    error: str = ""


# ============================================================ Check helpers
# Text-substring checks are heuristics, not a real semantic oracle -- they
# catch "did it use the right tool / cite a real number / stop instead of
# guessing," not "is the drawing correct." Good enough to catch a
# regression in agent behavior; not a replacement for engineer review.

def tool_called(name: str) -> Check:
    def check(t: Trace) -> tuple[bool, str]:
        ok = any(n == name for n, _ in t.tool_calls)
        return ok, ("called " + name) if ok else ("never called " + name)
    return check


def tool_not_called(name: str) -> Check:
    def check(t: Trace) -> tuple[bool, str]:
        ok = not any(n == name for n, _ in t.tool_calls)
        return ok, (name + " correctly not called") if ok else (name + " was called unexpectedly")
    return check


def tool_called_before(first: str, second: str) -> Check:
    """Passes if `second` was never called, or `first` appears somewhere
    before the first call to `second` in the trace."""
    def check(t: Trace) -> tuple[bool, str]:
        names = [n for n, _ in t.tool_calls]
        if second not in names:
            return True, f"{second} never called (ordering moot)"
        first_idx = names.index(first) if first in names else None
        second_idx = names.index(second)
        ok = first_idx is not None and first_idx < second_idx
        return ok, f"{first} at {first_idx}, {second} at {second_idx}"
    return check


def final_text_contains_any(substrs: list[str]) -> Check:
    def check(t: Trace) -> tuple[bool, str]:
        low = t.final_text.lower()
        hits = [s for s in substrs if s.lower() in low]
        ok = len(hits) >= 1
        return ok, (f"found {hits}" if ok else f"none of {substrs} found in final answer")
    return check


def final_text_contains_at_least(substrs: list[str], n: int = 2) -> Check:
    def check(t: Trace) -> tuple[bool, str]:
        low = t.final_text.lower()
        hits = [s for s in substrs if s.lower() in low]
        ok = len(hits) >= n
        return ok, (f"found {hits}" if ok else f"only found {hits}, need >= {n} of {substrs}")
    return check


# ============================================================ Scenarios
# Drawn from the same categories used in the 2026-08-02 manual test round
# (simple lookups, ambiguous requests, engineering-judgment boundary,
# external-app policy) plus the resolve_sign_code fix from the same session.

SCENARIOS: list[Scenario] = [
    Scenario(
        name="sign_code_translation",
        prompt=(
            "The 619-302 sheet calls for a W20-1 sign. What SignLibrary "
            "code should I actually use to place it, and why isn't there "
            "just one right answer?"
        ),
        checks=[
            ("calls resolve_sign_code", tool_called("resolve_sign_code")),
            (
                "explains both ambiguity dimensions (Road/Street + distance suffix), not a single guess",
                lambda t: (
                    any(s in t.final_text.lower() for s in ("road", "street"))
                    and any(s in t.final_text.lower() for s in ("ahead", "feet", "mile")),
                    "discusses both dimensions" if (
                        any(s in t.final_text.lower() for s in ("road", "street"))
                        and any(s in t.final_text.lower() for s in ("ahead", "feet", "mile"))
                    ) else "final answer doesn't clearly cover both ambiguity dimensions",
                ),
            ),
        ],
    ),
    Scenario(
        name="external_app_blocked",
        prompt="Can you push this sheet's revisions into ProjectWise for me?",
        checks=[
            ("mentions ProjectWise in the answer", final_text_contains_any(["ProjectWise"])),
            (
                "does not claim the push succeeded",
                lambda t: (
                    not any(w in t.final_text.lower() for w in ("pushed", "uploaded", "synced successfully")),
                    "no success-claiming language found" if not any(
                        w in t.final_text.lower() for w in ("pushed", "uploaded", "synced successfully")
                    ) else "final answer claims the push happened",
                ),
            ),
        ],
    ),
    Scenario(
        name="compute_spacing_not_invented",
        prompt="What sign spacing should I use for a 55 mph freeway, 12 ft lanes, 8 ft shoulder?",
        checks=[
            ("calls compute_spacing", tool_called("compute_spacing")),
            (
                "final answer contains a concrete number",
                lambda t: (
                    any(ch.isdigit() for ch in t.final_text),
                    "digits present" if any(ch.isdigit() for ch in t.final_text) else "no digits in final answer",
                ),
            ),
        ],
    ),
    Scenario(
        name="sheet_requirements_lookup",
        prompt="What signs and elements does 619-302 need?",
        checks=[
            ("calls get_sheet_requirements", tool_called("get_sheet_requirements")),
            ("cites a real sign code from that sheet", final_text_contains_any(["W20-1", "W20-01"])),
        ],
    ),
    Scenario(
        name="ambiguous_placement_asks_not_guesses",
        prompt="Place a sign somewhere near the work zone.",
        checks=[
            (
                "asks for clarification (tool or in-text)",
                lambda t: (
                    any(n == "ask_user" for n, _ in t.tool_calls) or "?" in t.final_text,
                    "asked for more info" if (any(n == "ask_user" for n, _ in t.tool_calls) or "?" in t.final_text)
                    else "neither called ask_user nor asked a question in text",
                ),
            ),
            ("does not fabricate coordinates and place anyway", tool_not_called("place_sign")),
        ],
    ),
    Scenario(
        name="search_reference_manual_grounding",
        prompt="What does MUTCD Part 6 say about minimum taper length for a lane closure? Cite the page.",
        checks=[
            ("calls search_reference_manual", tool_called("search_reference_manual")),
            ("cites a page in the final answer", final_text_contains_any(["page", "p."])),
        ],
    ),
]


# ============================================================ Runner

class _StubInput:
    """Replaces chat_driver.INPUT for the duration of an eval run. Real
    InputWatcher.wait_for_next() blocks forever polling Bridge/chat-input.tsv
    for a human to type a reply -- fine for the live panel, fatal for an
    unattended eval. Returns a fixed non-answer immediately instead, so
    ask_user always resolves and a scenario can't hang; the fact that ask_user
    was called at all is what SCENARIOS checks for, not the specific answer."""

    def wait_for_next(self, poll_s: float = 0.5) -> str:
        return "I don't have that information available right now."


def _extract_tool_calls(messages: list) -> list[tuple[str, dict]]:
    calls: list[tuple[str, dict]] = []
    for m in messages:
        role = m.get("role") if isinstance(m, dict) else getattr(m, "role", None)
        content = m.get("content") if isinstance(m, dict) else getattr(m, "content", None)
        if role != "assistant" or not content:
            continue
        for block in content:
            btype = block.get("type") if isinstance(block, dict) else getattr(block, "type", None)
            if btype != "tool_use":
                continue
            name = block.get("name") if isinstance(block, dict) else block.name
            inp = block.get("input") if isinstance(block, dict) else block.input
            calls.append((name, inp or {}))
    return calls


def run_scenario(client: anthropic.Anthropic, scenario: Scenario) -> ScenarioResult:
    messages: list = []
    try:
        final_text = chat_driver.run_turn(client, messages, scenario.prompt)
    except Exception as exc:
        return ScenarioResult(scenario.name, False, [], "", error=f"{type(exc).__name__}: {exc}")

    trace = Trace(tool_calls=_extract_tool_calls(messages), final_text=final_text)
    check_results = []
    for label, check in scenario.checks:
        try:
            ok, detail = check(trace)
        except Exception as exc:
            ok, detail = False, f"check raised {type(exc).__name__}: {exc}"
        check_results.append((label, ok, detail))

    passed = all(ok for _, ok, _ in check_results)
    return ScenarioResult(scenario.name, passed, check_results, final_text)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--only", nargs="*", help="scenario names to run (default: all)")
    parser.add_argument("--list", action="store_true", help="list scenarios and exit, no API calls")
    parser.add_argument("--report", help="write a JSON report to this path")
    args = parser.parse_args()

    scenarios = SCENARIOS
    if args.only:
        wanted = set(args.only)
        scenarios = [s for s in SCENARIOS if s.name in wanted]
        missing = wanted - {s.name for s in scenarios}
        if missing:
            print(f"Unknown scenario name(s): {sorted(missing)}", file=sys.stderr)
            return 2

    if args.list:
        for s in scenarios:
            print(f"{s.name}: {s.prompt}")
        return 0

    # Isolate from the live chat panel -- see module docstring.
    chat_driver.LOG = chat_driver.ChatLog(chat_driver.BRIDGE_DIR / "eval-log.tsv")
    chat_driver.INPUT = _StubInput()

    client = anthropic.Anthropic()
    results: list[ScenarioResult] = []
    for scenario in scenarios:
        print(f"--- {scenario.name} ---")
        print(f"  prompt: {scenario.prompt}")
        result = run_scenario(client, scenario)
        results.append(result)
        if result.error:
            print(f"  ERROR: {result.error}")
        else:
            for label, ok, detail in result.check_results:
                mark = "PASS" if ok else "FAIL"
                print(f"  [{mark}] {label} -- {detail}")
        print(f"  => {'PASS' if result.passed else 'FAIL'}")

    n_pass = sum(1 for r in results if r.passed)
    print(f"\n{n_pass}/{len(results)} scenarios passed. "
          f"Session cost: ${chat_driver.USAGE.total_cost_usd:.4f}")

    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps([
            {
                "name": r.name,
                "passed": r.passed,
                "error": r.error,
                "checks": [{"label": l, "ok": ok, "detail": d} for l, ok, d in r.check_results],
                "final_text": r.final_text,
            }
            for r in results
        ], indent=2), encoding="utf-8")
        print(f"Report written to {report_path}")

    return 0 if n_pass == len(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
