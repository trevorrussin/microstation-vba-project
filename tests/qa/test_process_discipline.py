"""Process discipline enforced in code (2026-10-24 review of the "turn off the circles" run):
hook scoping + dedupe, failed-attempt budget, evidence labels, COM-script guard, process check."""
from __future__ import annotations

import json
import os
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import retrieval  # noqa: E402
import task_guard  # noqa: E402
from qa import claims  # noqa: E402
from qa.checks import process  # noqa: E402

PASTED_REVIEW = ("what do you think?\n<pasted_content id=\"ab12\">Level Display levels references alignments model view "
                 "reference levels alignment station playbook</pasted_content id=\"ab12\">")


# ------------------------------------------------------------------ retrieval hook scoping and dedupe
def test_pasted_material_does_not_arm_the_hook():
    assert retrieval.hook_matches(PASTED_REVIEW) == []


def test_a_discussion_question_does_not_arm_the_hook():
    assert retrieval.hook_matches("what do you think about the levels and references approach in the alignment playbook?") == []


def test_one_stray_keyword_is_not_a_topic():
    assert retrieval.hook_matches("add a new level for the sign") == [] or len(retrieval.hook_matches("add a new level for the sign")) <= retrieval.HOOK_MAX_TOPICS


def test_a_real_task_still_matches_and_is_capped():
    m = retrieval.hook_matches("attach the reference file, level display shows nothing, turn on levels in the model view")
    assert m and len(m) <= retrieval.HOOK_MAX_TOPICS
    assert m[0]["topic"] == "references-levels-models-measuring"


@pytest.fixture
def pending(tmp_path, monkeypatch):
    monkeypatch.setattr(retrieval, "PENDING", str(tmp_path / "pending.json"))
    return tmp_path


def test_same_demand_is_not_repeated_once_the_stop_hook_gave_up(pending):
    topics = ["references-levels-models-measuring"]
    assert not retrieval.already_told(topics)
    retrieval.set_pending(topics, need_search=True)
    assert not retrieval.already_told(topics)  # not given up yet
    p = retrieval.load_pending()
    p["blocks"] = retrieval.MAX_PENDING_BLOCKS
    json.dump(p, open(retrieval.PENDING, "w"))
    assert retrieval.already_told(topics)
    assert not retrieval.already_told(["alignments-horizontal-vertical"])  # a different demand is new


def test_stop_hook_gives_up_after_the_cap_and_says_so(pending, monkeypatch):
    monkeypatch.setattr(retrieval, "has_fresh_receipt", lambda t, now=None: False)
    retrieval.set_pending(["references-levels-models-measuring"], need_search=True)
    msgs = [retrieval.stop_message() for _ in range(retrieval.MAX_PENDING_BLOCKS + 1)]
    assert all(msgs[:retrieval.MAX_PENDING_BLOCKS]) and "NOT backed" in msgs[-1]
    assert retrieval.stop_message() == ""  # a failed lookup never traps the session


# ------------------------------------------------------------------ failed-attempt budget
@pytest.fixture
def guard(tmp_path, monkeypatch):
    tasks = tmp_path / "tasks"
    (tasks / "T1").mkdir(parents=True)
    json.dump({"id": "T1", "type": "cad.drafting", "status": "in_progress", "created": time.time(), "candidates": []},
              open(tasks / "T1" / "task.json", "w"))
    monkeypatch.setattr(task_guard, "TASKS", tasks)
    monkeypatch.setattr(task_guard, "ATTEMPTS", tmp_path / "attempts.jsonl")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    return tmp_path


def _fresh():
    task_guard._cache.update(at=0.0, open=None)


def test_two_failed_edits_then_the_third_is_refused(guard):
    for _ in range(task_guard.MAX_FAILED_ATTEMPTS):
        task_guard.record_attempts(["RUN_REGISTRY_COMMAND"], [{"status": "ERROR", "note": "Cannot modify library level"}])
    _fresh()
    with pytest.raises(task_guard.ReplanRequired) as ex:
        task_guard.require_open_task(["RUN_REGISTRY_COMMAND"])
    assert "qa_gate.py replan" in str(ex.value) and "escalate" in str(ex.value)


def test_reads_are_never_refused_by_the_failure_budget(guard):
    for _ in range(5):
        task_guard.record_attempts(["MOVE_ELEMENT"], RuntimeError("boom"))
    _fresh()
    task_guard.require_open_task(["GET_ELEMENTS_RANGE", "DESCRIBE_DRAWING_STATE"])


def test_a_success_resets_the_count(guard):
    task_guard.record_attempts(["MOVE_ELEMENT"], [{"status": "ERROR"}])
    task_guard.record_attempts(["MOVE_ELEMENT"], [{"status": "OK"}])
    task_guard.record_attempts(["MOVE_ELEMENT"], [{"status": "ERROR"}])
    _fresh()
    assert task_guard.failed_streak("T1") == 1
    task_guard.require_open_task(["MOVE_ELEMENT"])


def test_replan_reopens_the_door(guard):
    for _ in range(2):
        task_guard.record_attempts(["MOVE_ELEMENT"], RuntimeError("x"))
    task_guard.record_replan("T1", "library levels refuse changes; hand the Level Display step to the engineer")
    _fresh()
    assert task_guard.failed_streak("T1") == 0
    task_guard.require_open_task(["MOVE_ELEMENT"])


def test_bridge_logs_failures_and_the_exception_still_propagates(guard, monkeypatch):
    import bridge_client
    b = bridge_client.Bridge(request_file=guard / "rq.tsv", response_file=guard / "rs.tsv")
    monkeypatch.setattr(b, "_call_batch_impl", lambda ops: (_ for _ in ()).throw(bridge_client.BridgeError("no response")))
    monkeypatch.setattr(task_guard, "require_open_task", lambda ops: None)
    with pytest.raises(bridge_client.BridgeError):
        b.call_batch([("MOVE_ELEMENT", {"elementId": "1"})])
    assert task_guard.failed_streak("T1") == 1


# ------------------------------------------------------------------ evidence labels
def test_observed_and_inferred_need_evidence():
    t = {}
    with pytest.raises(ValueError):
        claims.add(t, "observed", "0 elements on the Draft_ levels", "", set())
    assert claims.add(t, "observed", "0 elements on the Draft_ levels", "scan of 5 levels, 0 rows", set())["id"] == 1


def test_verified_must_cite_a_check_that_passed():
    t = {}
    with pytest.raises(ValueError):
        claims.add(t, "verified", "circles are hidden in the view", "I looked at a screenshot", {"COM-01"})
    with pytest.raises(ValueError):
        claims.add(t, "verified", "circles are hidden in the view", "COM-01", {"CAD-LVL-01"})
    c = claims.add(t, "verified", "circles are hidden in the view", "COM-01 and VIS-01", {"COM-01"})
    assert c["level"] == "verified"


def test_open_claims_block_and_upgrade_or_withdraw_clears_them():
    t = {}
    c = claims.add(t, "hypothesis", "the rings are transient Civil graphics", "", set())
    assert claims.open_claims(t)
    claims.add(t, "inferred", "", "0 elements on 5 levels; Exa: Sight Visibility is transient", set(), c["id"])
    assert not claims.open_claims(t) and claims.inferred_claims(t)
    claims.add(t, "unknown", "which level draws the rings", "", set())
    claims.withdraw(t, 2)
    assert not claims.open_claims(t)


def test_supported_and_contradicted_levels():
    t = {}
    with pytest.raises(ValueError):
        claims.add(t, "supported", "the rings come from Civil display", "", set())      # needs evidence
    s = claims.add(t, "supported", "the rings come from Civil display", "0 elements on 5 levels; Exa note", set())
    assert not claims.open_claims(t) and [c["id"] for c in claims.inferred_claims(t)] == [s["id"]]
    claims.add(t, "contradicted", "the rings are saved elements", "range scan found 0 type-15 elements", set())
    assert [c["level"] for c in claims.open_claims(t)] == ["contradicted"]            # blocks close until resolved


# ------------------------------------------------------------------ process check
def ev(*kinds):
    return [{"kind": k if isinstance(k, str) else k[0], "at": i, **({} if isinstance(k, str) else k[1])} for i, k in enumerate(kinds)]


def test_good_run_passes():
    e = ev("task_open", "read", "search", "attempt_fail", "read", ("replan", {}), "attempt_ok", "verify_pass", ("outcome", {"value": "pass"}))
    assert process.violations(e) == []


def test_edit_before_task_or_read_fails():
    assert any(v.startswith("P1") for v in process.violations(ev("attempt_ok", "task_open")))
    assert any(v.startswith("P2") for v in process.violations(ev("task_open", "attempt_ok")))


def test_third_failed_edit_in_a_row_fails():
    e = ev("task_open", "read", "attempt_fail", "attempt_fail", "attempt_fail")
    assert any(v.startswith("P3") for v in process.violations(e))


def test_replan_without_new_reading_fails():
    e = ev("task_open", "read", "attempt_fail", "attempt_fail", "replan")
    assert any(v.startswith("P4") for v in process.violations(e))


def test_blocked_needs_an_escalation_reason_and_pass_needs_a_verifying_run():
    assert any(v.startswith("P5") for v in process.violations(ev("task_open", ("outcome", {"value": "blocked"}))))
    assert not process.violations(ev("task_open", ("escalate", {"reason": "library levels refuse agent changes"}), ("outcome", {"value": "blocked"})))
    assert any(v.startswith("P5") for v in process.violations(ev("task_open", ("outcome", {"value": "pass"}))))


# ------------------------------------------------------------------ COM-script guard
import qa_pretool_hook as pre  # noqa: E402

COM_MUTATE = "from ms_connect import get_microstation_app\napp = get_microstation_app()\napp.CadInputQueue.SendKeyin('LEVEL DISPLAY OFF \"X\"')\n"
COM_READ = "from ms_connect import get_microstation_app\napp = get_microstation_app()\nprint(app.ActiveModelReference.Name)\n"


def test_com_script_that_edits_is_refused_without_an_open_task(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    s = tmp_path / "hide_levels.py"
    s.write_text(COM_MUTATE)
    msg = pre.assess(f"python {s}", str(tmp_path))
    assert "QA gate refuses" in msg and "SendKeyin" in msg and "references-levels-models-measuring" in msg


def test_inline_python_is_checked_too(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    assert pre.assess("python -c \"from ms_connect import get_microstation_app as g; g().CadInputQueue.SendKeyin('x')\"")


def test_read_only_com_script_passes(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    s = tmp_path / "look.py"
    s.write_text(COM_READ)
    assert pre.assess(f"python {s}", str(tmp_path)) == ""


def test_exempt_maintenance_script_and_explicit_scratch_marker_pass(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    assert pre.assess("python mcp-server/hot_reload.py Modules/WZTCBridge.bas", str(ROOT)) == ""
    s = tmp_path / "scratch.py"
    s.write_text(COM_MUTATE)
    assert pre.assess(f"QA_ALLOW_NO_TASK=1 python {s}", str(tmp_path)) == ""


def test_com_edit_passes_with_an_open_task_and_nothing_left_to_read(guard, tmp_path):
    s = tmp_path / "edit.py"
    s.write_text(COM_MUTATE)
    assert pre.assess(f"python {s}", str(tmp_path)) == ""


def test_com_edit_is_refused_when_the_failure_budget_is_spent(guard, tmp_path):
    for _ in range(2):
        task_guard.record_attempts(["RUN_REGISTRY_COMMAND"], RuntimeError("x"))
    s = tmp_path / "edit.py"
    s.write_text(COM_MUTATE)
    assert "replan" in pre.assess(f"python {s}", str(tmp_path))


def test_non_python_commands_pass(tmp_path):
    assert pre.assess("git status") == ""


def test_a_command_that_only_names_a_com_file_runs_nothing(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    s = tmp_path / "edit.py"
    s.write_text(COM_MUTATE)
    assert pre.assess(f"grep -n SendKeyin {s}", str(tmp_path)) == ""
    assert pre.assess(f"git diff -- {s}", str(tmp_path)) == ""
    assert pre.assess(f"python {s}", str(tmp_path))  # running it is still refused


def test_stdin_python_is_scanned_as_code(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "none")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    cmd = "python - <<'EOF'\n" + COM_MUTATE + "EOF"
    assert pre.assess(cmd, str(tmp_path))
