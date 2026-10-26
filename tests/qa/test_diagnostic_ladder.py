"""Diagnostic ladder, experience memory, reversible A/B experiment, trajectory metrics (2026-10-24, after the circles run)."""
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

import experience  # noqa: E402
import ladders  # noqa: E402
import task_guard  # noqa: E402
from qa import experiment, trajectory  # noqa: E402
from qa.checks import process  # noqa: E402

CIRCLES = "Hide the yellow and white circles on the plan view, they are not showing as design linework"


# ------------------------------------------------------------------ ladder matching and gating
def test_circles_request_matches_the_ladder_and_unrelated_requests_do_not():
    assert ladders.match(CIRCLES)["id"] == "unwanted-or-missing-graphics"
    assert ladders.match("Move the sign 10 ft left of the alignment") is None
    assert ladders.match("circle") is None  # one keyword is not a symptom


def _task():
    t = {"id": "T", "created": 0.0}
    ladders.attach(t, ladders.match(CIRCLES))
    return t


def test_edit_rung_is_gated_until_every_cheaper_rung_has_a_result():
    t = _task()
    assert [g.split(":")[0] for g in ladders.gaps(t)] == ["ladder R1 (check)", "ladder R2 (ask the engineer)",
                                                          "ladder R3 (ask the engineer)", "ladder R4 (check)"]
    ladders.record(t, "R1", "negative", "find_elements_near at the circles: zero stored elements within 50 ft")
    ladders.record(t, "R2", "handed_off", "asked the engineer to turn View Display off for Draft_Sight levels")
    assert len(ladders.gaps(t)) == 2
    ladders.record(t, "R3", "negative", "engineer reports View Options - Civil items already off")
    ladders.record(t, "R4", "na", "no references attached to this model")
    assert ladders.gaps(t) == []


def test_a_positive_rung_ends_the_diagnosis():
    t = _task()
    ladders.record(t, "R1", "negative", "found nothing stored at the symptom in the sampled box")
    ladders.record(t, "R2", "positive", "engineer turned Draft_Sight View Display off and the circles vanished")
    assert ladders.gaps(t) == []


def test_rung_results_need_evidence_and_an_engineer_step_cannot_be_checked_by_the_agent_as_handed_off():
    t = _task()
    with pytest.raises(ValueError):
        ladders.record(t, "R1", "negative", "ok")
    with pytest.raises(ValueError):
        ladders.record(t, "R1", "handed_off", "agent step must be checked, not handed off")
    with pytest.raises(ValueError):
        ladders.record(t, "R9", "negative", "no such rung exists in this ladder at all")
    with pytest.raises(ValueError):
        ladders.record(t, "R1", "maybe", "status outside the allowed list of four")


@pytest.fixture
def guard(tmp_path, monkeypatch):
    tasks = tmp_path / "tasks"
    (tasks / "T1").mkdir(parents=True)
    t = {"id": "T1", "type": "cad.drafting", "status": "in_progress", "created": time.time(), "candidates": []}
    ladders.attach(t, ladders.match(CIRCLES))
    json.dump(t, open(tasks / "T1" / "task.json", "w"))
    monkeypatch.setattr(task_guard, "TASKS", tasks)
    monkeypatch.setattr(task_guard, "ATTEMPTS", tmp_path / "attempts.jsonl")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    return tasks / "T1" / "task.json"


def test_bridge_guard_refuses_edits_while_the_ladder_is_open_and_allows_reads(guard):
    task_guard.require_open_task(["GET_ELEMENT_VERTICES"])
    with pytest.raises(task_guard.PlaybookNotRead) as ex:
        task_guard.require_open_task(["RUN_REGISTRY_COMMAND"])
    assert "ladder R1" in str(ex.value) and "qa_gate rung" in str(ex.value)


def test_bridge_guard_lets_the_edit_through_once_the_ladder_is_done(guard):
    t = json.load(open(guard))
    for rid, st, who in (("R1", "negative", 0), ("R2", "handed_off", 0), ("R3", "handed_off", 0), ("R4", "negative", 0)):
        ladders.record(t, rid, st, "recorded by the test with enough text to count as evidence")
    json.dump(t, open(guard, "w"))
    task_guard._cache.update(at=0.0, open=None)
    task_guard.require_open_task(["RUN_REGISTRY_COMMAND"])


# ------------------------------------------------------------------ PROC-01 rule P6
def ev(*items):
    return [{"kind": k if isinstance(k, str) else k[0], "at": i, **({} if isinstance(k, str) else k[1])} for i, k in enumerate(items)]


def test_p6_flags_a_scripted_edit_before_the_cheap_rungs():
    e = ev("task_open", ("ladder_open", {"gating": ["R1", "R2"]}), "read", ("rung", {"rung": "R1", "status": "negative"}), "attempt_fail")
    assert any(v.startswith("P6") and "R2" in v for v in process.violations(e))


def test_p6_passes_when_every_gating_rung_has_a_result_or_one_is_positive():
    done = ev("task_open", ("ladder_open", {"gating": ["R1", "R2"]}), "read", ("rung", {"rung": "R1", "status": "negative"}),
              ("rung", {"rung": "R2", "status": "handed_off"}), "attempt_ok", "verify_pass")
    pos = ev("task_open", ("ladder_open", {"gating": ["R1", "R2"]}), "read", ("rung", {"rung": "R1", "status": "positive"}), "attempt_ok")
    assert process.violations(done) == [] and process.violations(pos) == []


# ------------------------------------------------------------------ experience memory
@pytest.fixture
def exp(tmp_path, monkeypatch):
    monkeypatch.setattr(experience, "PATH", str(tmp_path / "experience.jsonl"))
    return experience


def test_unresolved_stays_unresolved_and_findings_need_a_level(exp):
    r = exp.add("Hide yellow circles on the plan view in the alt A file", "unresolved",
                [{"what": "COM level change", "result": "rejected"}], [{"level": "observed", "text": "Cannot modify library level attribute"}])
    assert r["status"] == "unresolved"
    with pytest.raises(ValueError):
        exp.add("Hide yellow circles on the plan view", "solved", [], [])
    with pytest.raises(ValueError):
        exp.add("Hide yellow circles on the plan view", "unresolved", [], [{"level": "certain", "text": "x"}])
    with pytest.raises(ValueError):
        exp.add("too short", "unresolved", [], [])


def test_search_finds_the_similar_problem_and_not_an_unrelated_one(exp):
    exp.add("Hide yellow circles on the plan view in the alt A file", "handed_off", [], [{"level": "reported", "text": "GUI step"}])
    exp.add("Offset the profile grade line 5 ft below the existing ground", "resolved", [], [{"level": "observed", "text": "copy_parallel"}])
    hit = exp.search("turn off the circles showing in the plan view")
    assert [h["symptom"].startswith("Hide yellow") for h in hit] == [True]
    assert exp.search("place a stop sign on the shoulder") == []
    assert "UNRESOLVED row is not a solution" in exp.for_task("circles in the plan view yellow hide")


def test_an_escalated_task_becomes_a_handed_off_or_unresolved_row_and_a_closed_task_a_resolved_one():
    esc = {"id": "T", "status": "escalated", "desc": "Hide Sight Visibility yellow circles", "escalation": "GUI only", "claims": [
        {"level": "hypothesis", "text": "a guess that must not become a finding"}, {"level": "observed", "text": "API refused the change"}],
        "ladder": {"id": "x", "rungs": {"R2": {"status": "handed_off", "evidence": "asked engineer", "at": 1}}}}
    row = experience.from_task(esc)
    assert row["status"] == "handed_off"
    assert [f["level"] for f in row["findings"]] == ["observed", "reported"]  # the hypothesis is not stored as a finding
    esc["ladder"] = {}
    assert experience.from_task(esc)["status"] == "unresolved"
    assert experience.from_task({**esc, "status": "closed"})["status"] == "resolved"
    assert experience.from_task({**esc, "status": "in_progress"}) is None


def test_the_seeded_circles_row_is_in_the_tracked_memory_and_marked_reported_not_verified():
    rows = [r for r in experience.rows() if "Sight Visibility" in r["symptom"]]
    assert rows and rows[0]["status"] == "handed_off"
    assert all(f["level"] in experience.LEVELS for f in rows[0]["findings"])


# ------------------------------------------------------------------ reversible A/B experiment
def _img(tmp_path, name, blobs):
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (100, 100), "white")
    d = ImageDraw.Draw(im)
    for box, col in blobs:
        d.rectangle(box, fill=col)
    p = str(tmp_path / name)
    im.save(p)
    return p


class World:
    """Stand-in for MicroStation: a state dict and a picture that depends on it."""

    def __init__(self, tmp_path, fail_restore=False, hides_more=False, side_effect=False):
        self.tmp, self.state, self.n = tmp_path, {"civil_circles": "on", "levels": "A"}, 0
        self.fail_restore, self.hides_more, self.side_effect = fail_restore, hides_more, side_effect

    def snapshot(self):
        return dict(self.state)

    def apply(self):
        self.state["civil_circles"] = "off"
        if self.side_effect:
            self.state["levels"] = "B"

    def restore(self):
        self.state["civil_circles"] = "on"
        if not self.fail_restore:
            self.state["levels"] = "A"

    def capture(self):
        self.n += 1
        blobs = [((60, 60, 80, 80), "black")]  # linework that must stay
        if self.state["civil_circles"] == "on":
            blobs.append(((10, 10, 30, 30), "yellow"))
        if self.hides_more and self.state["civil_circles"] == "off":
            blobs = []
        return _img(self.tmp, f"c{self.n}.png", blobs)


def _run(w, **kw):
    return experiment.run_ab("Civil circles are drawn by the Civil display setting", w.snapshot, w.apply, w.restore, w.capture,
                             region=(0, 0, 50, 50), protect=[(55, 55, 90, 90)], expect_changed={"civil_circles"}, **kw)


def test_ab_supported_when_only_the_target_region_changes_and_state_is_restored(tmp_path):
    w = World(tmp_path)
    r = _run(w)
    assert r["verdict"] == "SUPPORTED" and w.state == {"civil_circles": "on", "levels": "A"}


def test_ab_refuted_when_the_target_does_not_change(tmp_path):
    w = World(tmp_path)
    w.apply = lambda: None
    assert _run(w)["verdict"] == "REFUTED"


def test_ab_side_effect_when_the_change_also_hid_protected_linework(tmp_path):
    r = _run(World(tmp_path, hides_more=True))
    assert r["verdict"] == "SIDE_EFFECT"


def test_ab_side_effect_when_an_unnamed_state_key_changed(tmp_path):
    r = _run(World(tmp_path, side_effect=True))
    assert r["verdict"] == "SIDE_EFFECT" and r["unexpected_state"] == ["levels"]


def test_ab_restore_failure_is_reported_and_overrides_the_verdict(tmp_path):
    r = _run(World(tmp_path, fail_restore=True, side_effect=True))
    assert r["verdict"] == "RESTORE_FAILED" and r["restore_diff"] == ["levels"]


def test_ab_restores_even_when_apply_raises(tmp_path):
    w = World(tmp_path)

    def boom():
        w.state["civil_circles"] = "off"
        raise RuntimeError("apply failed half way")

    w.apply = boom
    r = _run(w)
    assert r["verdict"] == "INCONCLUSIVE" and w.state["civil_circles"] == "on"


def test_ab_different_sized_captures_are_inconclusive(tmp_path):
    from PIL import Image
    w = World(tmp_path)
    cap = w.capture
    calls = []

    def capture():
        p = cap()
        if calls:
            Image.open(p).resize((50, 50)).save(p)
        calls.append(1)
        return p

    w.capture = capture
    assert _run(w)["verdict"] == "INCONCLUSIVE"


# ------------------------------------------------------------------ trajectory metrics
def test_trajectory_counts_redundant_failures_streaks_and_unread_edits():
    task = {"id": "T", "created": 0.0, "status": "escalated"}
    att = [{"kind": "attempt_fail", "op": "RUN_REGISTRY_COMMAND", "note": "Cannot modify library level", "at": 10 + i} for i in range(3)]
    att.append({"kind": "attempt_ok", "op": "MOVE_ELEMENT", "at": 20})
    events = [{"kind": "task_open", "at": 0}] + [{"kind": r["kind"], "at": r["at"]} for r in att]
    m = trajectory.task_metrics(task, events, att)
    assert (m["edits"], m["failed_edits"], m["redundant_failures"], m["longest_fail_streak"]) == (4, 3, 2, 3)
    assert m["read_before_edit"] is False and any(v.startswith("P2") for v in m["violations"])
    s = trajectory.summarize([m])
    assert s["redundant_failure_rate"] == 0.5 and s["verified_success_rate"] == 0.0 and s["tasks_with_logged_edits"] == 1


def test_trajectory_with_no_logged_edit_has_no_timing_instead_of_a_made_up_one():
    m = trajectory.task_metrics({"id": "T", "created": 0.0, "status": "closed", "last_status": "PASS"}, [{"kind": "read", "at": 5}], [])
    assert m["read_before_edit"] is None and m["secs_to_first_read"] is None and m["verified_success"] is True
