"""Round-2 audit changes: validators, COM cross-check, explicit task states, partial coverage, sign-off history,
corpus sensitivity."""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import qa_gate  # noqa: E402
import qa_stop_hook  # noqa: E402
import task_guard  # noqa: E402
from qa import corpus, geom  # noqa: E402
from qa.checks import com_spot, coverage, operations_check as oc  # noqa: E402
from tests.qa.corpus_cases import cases  # noqa: E402
from tests.qa.test_qa_checks import el, ctx, _txt  # noqa: E402
from tests.qa.test_qa_integrity import _vis  # noqa: E402


# ------------------------------------------------------------------ geometry helpers
def test_nearest_on_polyline_gives_a_signed_distance():
    line = [(0.0, 0.0), (10.0, 0.0)]
    assert geom.nearest_on_polyline((5.0, 3.0), line)[1] == pytest.approx(3.0)    # left of a->b
    assert geom.nearest_on_polyline((5.0, -2.0), line)[1] == pytest.approx(-2.0)  # right
    assert geom.nearest_on_polyline((15.0, 0.0), line)[0] == pytest.approx(5.0)   # beyond the end: distance to the end point


def test_box_to_polyline_distance_is_zero_when_touching_or_crossing_and_positive_otherwise():
    line = [(0.0, 0.0), (100.0, 50.0)]
    assert geom.dist_box_polyline((40, 15, 60, 25), line) == 0.0       # crosses the diagonal
    assert geom.dist_box_polyline((40, 40, 70, 48), line) == pytest.approx(5 / 1.25 ** 0.5)  # above it: corner (70,40) is 4.47 ft off y=0.5x
    assert geom.dist_box_polyline((0, 0, 1, 1), line) == 0.0           # contains the start point


def test_y_at_x_interpolates_and_reports_when_the_line_does_not_span_x():
    assert geom.y_at_x([(0.0, 10.0), (100.0, 14.0)], 50.0) == pytest.approx(12.0)
    assert geom.y_at_x([(0.0, 10.0), (100.0, 14.0)], 150.0) is None


# ------------------------------------------------------------------ relations through the real check
def _rel(expect, geom_, boxes=None):
    c = ctx([])
    c["geom"], c["geom_boxes"] = geom_, boxes or {}
    return oc.relations({"expect": expect}, c)


def test_a_relation_row_missing_its_inputs_is_uncertain_not_a_crash():
    assert _rel([{"id": "R", "rel": "offset_curve", "a": "1", "b": "2"}], {"1": [(0, 0), (1, 0)], "2": [(0, 1), (1, 1)]})["status"] == "UNCERTAIN"
    assert _rel([{"id": "R", "rel": "side", "a": "t", "b": "g"}], {"g": [(0, 0), (10, 0)]}, {"t": (1, 1, 2, 2)})["status"] == "UNCERTAIN"


def test_relation_geometry_that_cannot_be_read_is_uncertain():
    assert _rel([{"id": "R", "rel": "distance", "a": "t", "b": "g", "min": 1}], {})["status"] == "UNCERTAIN"


def test_left_right_side_uses_the_direction_of_the_reference():
    g = {"g": [(0.0, 0.0), (10.0, 0.0)]}
    assert _rel([{"id": "R", "rel": "side", "a": "t", "b": "g", "target": "left"}], g, {"t": (4, 2, 6, 3)})["status"] == "PASS"
    assert _rel([{"id": "R", "rel": "side", "a": "t", "b": "g", "target": "left"}], g, {"t": (4, -3, 6, -2)})["status"] == "FAIL"


def test_expect_cli_builds_rich_relation_rows_and_rejects_unknown_kinds(tmp_path, monkeypatch):
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path))
    qa_gate.save_task({"id": "T1", "type": "cad.drafting", "desc": "d", "status": "in_progress", "created": 1, "expect": []})
    ns = argparse.Namespace(task="T1", id="V1", desc="clearance", where=None, count=None, value=None, target=None, rel="vclear", a="5", b="6",
                            tol=None, params='{"min": 14.5, "ve": 5, "x": [4800, 4850]}')
    assert qa_gate.cmd_expect(ns) == 0
    row = qa_gate.load_task("T1")["expect"][0]
    assert row["rel"] == "vclear" and row["min"] == 14.5 and row["x"] == [4800, 4850] and row["ve"] == 5
    with pytest.raises(SystemExit):
        qa_gate.cmd_expect(argparse.Namespace(task="T1", id="V2", desc="", where=None, count=None, value=None, target=None, rel="telepathy",
                                              a="1", b="2", tol=None, params=None))
    with pytest.raises(SystemExit):  # offset_curve without the distance
        qa_gate.cmd_expect(argparse.Namespace(task="T1", id="V3", desc="", where=None, count=None, value=None, target=None, rel="offset_curve",
                                              a="1", b="2", tol=None, params=None))


# ------------------------------------------------------------------ intent
def test_start_banner_states_what_the_gate_will_hold_the_operations_to():
    lines = oc.intent_lines("offset the proposed curb 12 ft to the right")
    assert any("12" in x for x in lines) and any("right" in x for x in lines)
    assert any("NO number" in x for x in oc.intent_lines("offset the curb"))


# ------------------------------------------------------------------ COM cross-check
def _com_ctx(rows, **kw):
    c = ctx([rows[0]], **kw)
    c["state_rows"], c["state_fp"] = rows, "fp"
    return c


def _with_com(monkeypatch, answer):
    monkeypatch.setattr(com_spot, "read_com", (lambda ids: answer(ids)) if callable(answer) else (lambda ids: {k: v for k, v in answer.items() if k in ids}))


def test_com_check_confirms_agreeing_evidence_and_flags_every_kind_of_disagreement(monkeypatch):
    rows = [el(1, box=(0, 0, 10, 1)), el(2, box=(5, 5, 6, 6))]
    good = {"1": {"low": (0.0, 0.0), "high": (10.0, 1.0), "level": "TWZCD_P"}, "2": {"low": (5.0, 5.0), "high": (6.0, 6.0), "level": "TWZCD_P"}}
    _with_com(monkeypatch, good)
    assert com_spot.verify({}, _com_ctx(rows))["status"] == "PASS"
    moved = dict(good, **{"2": {"low": (5.0, 5.0), "high": (6.0, 9.0), "level": "TWZCD_P"}})
    _with_com(monkeypatch, moved)
    assert com_spot.verify({}, _com_ctx(rows))["status"] == "UNCERTAIN"
    relevel = dict(good, **{"1": {"low": (0.0, 0.0), "high": (10.0, 1.0), "level": "Default"}})
    _with_com(monkeypatch, relevel)
    assert com_spot.verify({}, _com_ctx(rows))["status"] == "UNCERTAIN"
    _with_com(monkeypatch, {"1": good["1"]})  # element 2 not found through COM
    assert "not found" in " ".join(com_spot.verify({}, _com_ctx(rows))["evidence"])

    def boom(ids):
        raise RuntimeError("COM down")
    _with_com(monkeypatch, boom)
    assert com_spot.verify({}, _com_ctx(rows))["status"] == "UNCERTAIN"
    assert com_spot.verify({}, ctx([]))["status"] == "NA"


def test_com_sample_always_includes_the_changed_elements_and_is_reproducible():
    rows = [el(i) for i in range(1, 1000)]
    ids1, _ = com_spot.pick(rows, ["999", "5"], "fp")
    ids2, _ = com_spot.pick(rows, ["999", "5"], "fp")
    assert ids1 == ids2 and ids1[:2] == ["999", "5"] and len(ids1) == 2 + com_spot.SAMPLE


# ------------------------------------------------------------------ explicit task states
def _hook(tmp_path, monkeypatch, task):
    monkeypatch.setattr(qa_stop_hook, "TASKS", str(tmp_path))
    d = tmp_path / task["id"]
    d.mkdir(exist_ok=True)
    (d / "task.json").write_text(json.dumps(task))
    monkeypatch.setattr(sys, "stdin", io.StringIO("{}"))
    monkeypatch.setattr(sys, "argv", ["qa_stop_hook.py"])
    rc = qa_stop_hook.main()
    return rc, json.loads((d / "task.json").read_text())


def test_an_old_task_becomes_expired_and_is_reported_not_forgotten(tmp_path, monkeypatch, capsys):
    rc, t = _hook(tmp_path, monkeypatch, {"id": "T1", "type": "cad.drafting", "status": "in_progress", "created": time.time() - 14 * 3600, "last_status": "FAIL"})
    assert t["status"] == "expired" and "EXPIRED" in capsys.readouterr().out


def test_a_task_the_hook_gives_up_on_becomes_blocked(tmp_path, monkeypatch, capsys):
    base = {"id": "T1", "type": "cad.drafting", "status": "in_progress", "created": time.time(), "last_status": "FAIL", "last_run": time.time()}
    t = base
    for _ in range(qa_stop_hook.MAX_BLOCKS + 1):
        _, t = _hook(tmp_path, monkeypatch, t)
    assert t["status"] == "blocked" and "6 stop attempts" in t["state_reason"]
    assert "BLOCKED" in capsys.readouterr().out


def test_expired_and_blocked_tasks_stop_counting_as_open_until_reopened(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path)
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path))
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    qa_gate.save_task({"id": "T1", "type": "cad.drafting", "desc": "d", "status": "blocked", "created": time.time() - 14 * 3600, "state_reason": "x"})
    with pytest.raises(task_guard.QAGateRequired):
        task_guard.require_open_task(["MOVE_ELEMENT"])
    assert qa_gate.cmd_reopen(argparse.Namespace(task="T1")) == 0
    task_guard._cache.update(at=0.0, open=None)
    task_guard.require_open_task(["MOVE_ELEMENT"])  # reopened: counts again even though it is old
    assert qa_gate.cmd_reopen(argparse.Namespace(task="T1")) == 1  # only blocked/expired tasks reopen


# ------------------------------------------------------------------ partial coverage
def test_a_passing_proxy_check_is_reported_partial_never_verified():
    c = ctx([el(1, level="Default", typ="LINE", color="0", weight="0")])
    c["results"] = {k: {"status": "PASS"} for k in ("CAD-LVL-01", "CAD-SYM-01", "CAD-WT-01", "CAD-Z-01", "CAD-LOC-01", "CAD-DUP-01", "CAD-DEG-01")}
    coverage.report({}, c)
    rows = {r["requirement"]: r for r in c["coverage"]["requirements"]}
    assert rows["R-LOC"]["status"] == "partial" and "location is right for the design" in rows["R-LOC"]["remaining"]
    assert rows["R-LVL"]["status"] == "verified"


# ------------------------------------------------------------------ sign-off keeps the original result
def test_a_signoff_records_what_the_result_was_before_it(tmp_path):
    r, _ = _vis(tmp_path, ran=False, signoff={"items": ["VIS-01"], "by": "eng", "at": "now"})
    assert r["status"] == "PASS" and r["signed_off_from"]["status"] == "UNCERTAIN"
    r, _ = _vis(tmp_path / "b" if (tmp_path / "b").mkdir() is None else tmp_path, issues=[{"severity": "major", "description": "x", "where": ""}],
                signoff={"items": ["VIS-01"], "by": "eng", "at": "now"})
    assert r["status"] == "PASS" and r["signed_off_from"]["status"] == "FAIL"


# ------------------------------------------------------------------ the corpus is sensitive
def test_a_broken_gate_makes_the_corpus_report_false_passes(monkeypatch):
    """A corpus that cannot fail proves nothing: disable two checks and the defects they catch must show up as false passes."""
    monkeypatch.setattr(oc, "relations", lambda t, c: {"status": "PASS", "observed": "", "expected": "", "evidence": [], "fix": ""})
    monkeypatch.setattr(oc, "operations", lambda t, c: {"status": "PASS", "observed": "", "expected": "", "evidence": [], "fix": ""})
    rep = corpus.evaluate(cases())
    assert "REL-wrong-side" in rep["false_pass"] and "OP-move-wrong-distance" in rep["false_pass"]
    assert rep["false_pass_rate"] > 0.2
