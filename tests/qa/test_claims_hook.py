"""scripts/qa_claims_hook.py: a message that claims a check with no tool call in the turn is blocked."""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "scripts"))
import qa_claims_hook as h  # noqa: E402


def line(role, content):
    return json.dumps({"type": role, "message": {"role": role, "content": content}})


HUMAN = line("user", "where is alg 9 dont see it?")
TOOL = line("assistant", [{"type": "tool_use", "id": "t1", "name": "Bash", "input": {}}])
RESULT = line("user", [{"type": "tool_result", "tool_use_id": "t1", "content": "ok"}])


def final(text):
    return line("assistant", [{"type": "text", "text": text}])


def test_check_claim_without_tool_call_blocks():
    text, tools, _res = h.last_turn([HUMAN, final("I know. I checked: the model list, result: 4 models.")])
    assert tools == 0 and h.verdict(text, tools)


def test_check_claim_after_tool_call_passes():
    text, tools, _res = h.last_turn([HUMAN, TOOL, RESULT, final("I know. I checked: the model list, result: 4 models.")])
    assert tools == 1 and h.verdict(text, tools) == ""


def test_tool_call_in_an_earlier_turn_does_not_count():
    text, tools, _res = h.last_turn([HUMAN, TOOL, RESULT, final("done"), line("user", "ok now what"), final("I verified it earlier.")])
    assert tools == 0 and h.verdict(text, tools)


def test_hedge_about_state_blocks_but_plain_hedge_passes():
    assert h.verdict("I think the file is still in the model list.", 1)
    assert h.verdict("I think this naming is clearer.", 1) == ""


def test_quoted_text_is_ignored():
    assert h.verdict("The rule says `I checked` must be backed by a call.\n> I verified it", 0) == ""


def line_cursor(role, content):
    return json.dumps({"role": role, "message": {"content": content}})


def result(text):
    return line("user", [{"type": "tool_result", "tool_use_id": "t1", "content": text}])


def test_tests_passed_claim_needs_a_test_result_in_the_turn():
    said = final("354 passed after the change.")
    t, n, r = h.last_turn([HUMAN, TOOL, result("....  354 passed in 9s"), said])
    assert h.verdict(t, n, r) == ""
    t, n, r = h.last_turn([HUMAN, TOOL, result("file written"), said])
    assert h.verdict(t, n, r)


def test_gate_check_pass_claim_needs_that_check_in_a_tool_result():
    said = final("CAD-TXT-05 PASS: 85 text boxes, none overlap.")
    t, n, r = h.last_turn([HUMAN, TOOL, result("ok CAD-TXT-05 PASS: 85 text boxes in the model"), said])
    assert h.verdict(t, n, r) == ""
    t, n, r = h.last_turn([HUMAN, TOOL, result("ok CAD-TXT-03 PASS: clean"), said])
    assert h.verdict(t, n, r)
    t, n, r = h.last_turn([HUMAN, TOOL, result("XX CAD-TXT-05 FAIL: 3 overlaps"), said])
    assert h.verdict(t, n, r)


def test_cursor_transcript_format_is_read_and_skips_result_checks():
    lines = [line_cursor("user", [{"type": "text", "text": "where is it?"}]),
             line_cursor("assistant", [{"type": "text", "text": "I know. I checked: the model list."}])]
    t, n, r = h.last_turn(lines)
    assert n == 0 and h.verdict(t, n, None)                      # no tool call in the turn: blocked
    lines.insert(1, line_cursor("assistant", [{"type": "tool_use", "name": "Shell", "input": {}}]))
    t, n, r = h.last_turn(lines)
    assert n == 1 and h.verdict(t, n, None) == ""
    assert h.verdict("354 passed", 1, None) == ""                # results unavailable in Cursor: not judged


def test_no_claim_no_block():
    assert h.verdict("The ticks are on level XAX_P.", 0) == ""
