"""Release-integrity regression tests (2026-10-23 audit): a PASS must describe the artifact that is released, the
independent controls must not be satisfiable by the agent, and coverage gaps must not hide behind a green gate.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

import pytest
import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import qa_gate  # noqa: E402
import op_recorder  # noqa: E402
from qa import changeset, operations  # noqa: E402
from qa.checks import coverage, visual, operations_check as oc  # noqa: E402
from tests.qa.test_qa_checks import el, ctx, _txt  # noqa: E402
from tests.qa.test_qa_operations import _ctx as op_ctx, _t as op_task  # noqa: E402


# ------------------------------------------------------------------ change set: vertices, fingerprint
def _line(eid="1"):
    return el(eid, typ="LineString", box=(0, 0, 10, 10))


def test_vertex_edit_inside_an_unchanged_range_box_is_detected():
    row = _line()
    before = changeset.snapshot([row], {"1": changeset.vhash([[0, 0, 0], [10, 10, 0]])})
    same = changeset.diff(before, [row], {"1": changeset.vhash([[0, 0, 0], [10, 10, 0]])})
    assert not same["modified"]
    moved = changeset.diff(before, [row], {"1": changeset.vhash([[0, 0, 0], [5, 5, 0], [10, 10, 0]])})
    assert [r["elementId"] for r in moved["modified"]] == ["1"] and "vertices" in moved["modified_fields"]["1"]


def test_snapshot_without_vertex_hashes_does_not_flag_every_element():
    before = changeset.snapshot([_line()])  # an older task's snapshot
    assert not changeset.diff(before, [_line()], {"1": "abc"})["modified"]


def test_vertex_hash_ignores_read_back_noise_but_not_real_edits():
    assert changeset.vhash([[0, 0, 0]]) == changeset.vhash([[0.00001, 0, 0]])
    assert changeset.vhash([[0, 0, 0]]) != changeset.vhash([[0.01, 0, 0]])


def test_fingerprint_changes_with_vertices_properties_and_model():
    rows, vh = [_line()], {"1": "aaa"}
    base = changeset.fingerprint(rows, vh, "f|Design")
    assert base == changeset.fingerprint([_line()], {"1": "aaa"}, "f|Design")
    assert base != changeset.fingerprint(rows, {"1": "bbb"}, "f|Design")
    assert base != changeset.fingerprint([el("1", typ="LineString", level="Default", box=(0, 0, 10, 10))], vh, "f|Design")
    assert base != changeset.fingerprint(rows, vh, "f|Other")


# ------------------------------------------------------------------ close / STATE-01
def _open_task(tmp_path, monkeypatch, **kw):
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path))
    t = {"id": "T1", "type": "cad.drafting", "desc": "d", "status": "in_progress", "created": 1, "last_status": "PASS",
         "verified_fp": "aaa", "verified_model": "f|m", "runs": 1, "expect": [], "candidates": [], "citations": []}
    t.update(kw)
    qa_gate.save_task(t)
    monkeypatch.setattr(qa_gate, "_ops", lambda: object())
    return t


def _close(monkeypatch, fp, model="f|m"):
    monkeypatch.setattr(qa_gate, "state_fingerprint", lambda ops: (fp, model, None))
    return qa_gate.cmd_close(argparse.Namespace(task="T1"))


def test_close_is_blocked_when_the_drawing_changed_after_the_pass(tmp_path, monkeypatch, capsys):
    _open_task(tmp_path, monkeypatch)
    assert _close(monkeypatch, "bbb") == 1
    assert "changed after verification" in capsys.readouterr().out
    assert qa_gate.load_task("T1")["status"] == "in_progress"


def test_close_is_blocked_in_a_different_model(tmp_path, monkeypatch, capsys):
    _open_task(tmp_path, monkeypatch)
    assert _close(monkeypatch, "aaa", model="f|Other") == 1
    assert "Activate that model" in capsys.readouterr().out


def test_close_is_blocked_when_the_state_cannot_be_read(tmp_path, monkeypatch):
    _open_task(tmp_path, monkeypatch)

    def boom(ops):
        raise RuntimeError("bridge down")

    monkeypatch.setattr(qa_gate, "state_fingerprint", boom)
    assert qa_gate.cmd_close(argparse.Namespace(task="T1")) == 1


def test_close_succeeds_on_the_verified_state(tmp_path, monkeypatch):
    _open_task(tmp_path, monkeypatch)
    assert _close(monkeypatch, "aaa") == 0
    assert qa_gate.load_task("T1")["status"] == "closed"


def test_state_row_detects_a_drawing_that_changed_during_the_run(monkeypatch):
    monkeypatch.setattr(qa_gate, "_ops", lambda: object())
    monkeypatch.setattr(qa_gate, "state_fingerprint", lambda ops: ("same", "f|m", None))
    assert qa_gate._state_row({}, {"state_fp": "same"})["status"] == "PASS"
    assert qa_gate._state_row({}, {"state_fp": "other"})["status"] == "UNCERTAIN"
    assert qa_gate._state_row({}, {})["status"] == "UNCERTAIN"


def test_started_in_another_model_makes_the_change_set_uncertain():
    r = coverage.changes({}, {"model_mismatch": "task started in a|Design but the active model is a|Default", "changes": {"created": []}})
    assert r["status"] == "UNCERTAIN"


# ------------------------------------------------------------------ engineer sign-off
@pytest.fixture
def signoff_env(tmp_path, monkeypatch):
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path))
    monkeypatch.setattr(qa_gate, "KEY_PATH", str(tmp_path / "key"))
    t = {"id": "T1", "type": "generic", "desc": "d", "status": "in_progress", "created": 1, "expect": [], "candidates": [], "citations": []}
    qa_gate.save_task(t)
    return t


def _sign(monkeypatch, tty=True, answer="SIGN OFF"):
    monkeypatch.setattr(qa_gate, "_is_tty", lambda: tty)
    monkeypatch.setattr("builtins.input", lambda prompt="": answer)
    return qa_gate.cmd_signoff(argparse.Namespace(task="T1", items="VIS-01,R-VCLR", reason="reviewed"))


def test_an_agent_without_a_terminal_cannot_sign_off(signoff_env, monkeypatch):
    assert _sign(monkeypatch, tty=False) == 1
    assert qa_gate.load_signoff(signoff_env, None)["items"] == []


def test_signoff_needs_the_confirmation_phrase(signoff_env, monkeypatch):
    assert _sign(monkeypatch, answer="yes") == 1


def test_a_real_signoff_counts_and_a_forged_file_does_not(signoff_env, monkeypatch, tmp_path):
    assert _sign(monkeypatch) == 0
    assert qa_gate.load_signoff(signoff_env, None)["items"] == ["VIS-01", "R-VCLR"]
    p = tmp_path / "T1" / "signoff.json"
    so = json.load(open(p))
    so["items"].append("OP-01")  # tampered after signing
    json.dump(so, open(p, "w"))
    assert qa_gate.load_signoff(signoff_env, None) == {"items": [], "invalid": "signature does not verify"}
    json.dump({"items": ["VIS-01"], "fp": "", "sig": "made-up"}, open(p, "w"))
    assert qa_gate.load_signoff(signoff_env, None)["items"] == []


def test_signoff_is_void_after_the_drawing_changes(signoff_env, monkeypatch, tmp_path):
    t = dict(signoff_env, type="cad.drafting")
    qa_gate.save_task(t)
    monkeypatch.setattr(qa_gate, "_ops", lambda: object())
    monkeypatch.setattr(qa_gate, "state_fingerprint", lambda ops: ("fp1", "f|m", None))
    assert _sign(monkeypatch) == 0
    assert qa_gate.load_signoff(t, "fp1")["items"]
    assert qa_gate.load_signoff(t, "fp2") == {"items": [], "invalid": "drawing changed since the sign-off"}


# ------------------------------------------------------------------ vision review authority
def _vis(tmp_path, issues=(), disputes=None, ran=True, results=None, signoff=None, agent_issues=()):
    img = tmp_path / "render-1.png"
    img.write_bytes(b"png")
    json.dump({"ran": ran, "issues": list(issues), "model": "m", "t": time.time() + 100, "note": "no api key"}, open(tmp_path / "vision-review.json", "w"))
    json.dump({"t": time.time() + 200, "ok": not agent_issues, "issues": list(agent_issues), "checked": "looked at it",
               "disputes": disputes or {}}, open(tmp_path / "agent-review.json", "w"))
    c = {"task_dir": str(tmp_path), "images": [str(img)], "results": results or {}, "signoff": signoff or {}}
    return visual.review({}, c), c


ISSUE = {"severity": "major", "description": "level shows ?", "where": "legend"}


def test_the_agents_own_review_can_never_pass_vis01(tmp_path):
    r, c = _vis(tmp_path, ran=False)
    assert r["status"] == "UNCERTAIN" and any(n["id"] == "VIS-01" for n in c["needs_signoff"])


def test_engineer_signoff_passes_vis01_when_vision_is_unavailable(tmp_path):
    r, _ = _vis(tmp_path, ran=False, signoff={"items": ["VIS-01"], "by": "eng", "at": "now"})
    assert r["status"] == "PASS" and "sign-off" in r["observed"]


def test_agent_listed_issues_block_even_with_a_signoff(tmp_path):
    r, _ = _vis(tmp_path, ran=False, signoff={"items": ["VIS-01"]}, agent_issues=["label on the line"])
    assert r["status"] == "FAIL"


def test_open_reviewer_issue_fails(tmp_path):
    assert _vis(tmp_path, issues=[ISSUE])[0]["status"] == "FAIL"


def test_dispute_with_passing_deterministic_evidence_dismisses_the_issue(tmp_path):
    d = {"0": {"reason": "cell header level is empty by design", "evidence": ["CAD-LVL-01", "CAD-SYM-01"]}}
    r, _ = _vis(tmp_path, issues=[ISSUE], disputes=d, results={"CAD-LVL-01": {"status": "PASS"}, "CAD-SYM-01": {"status": "PASS"}})
    assert r["status"] == "PASS" and any("disputed #0" in e for e in r["evidence"])


def test_dispute_without_evidence_or_with_a_failing_check_does_not_dismiss(tmp_path):
    res = {"CAD-LVL-01": {"status": "FAIL"}}
    for d in ({"0": {"reason": "just wrong", "evidence": []}}, {"0": "legacy free text"},
              {"0": {"reason": "levels are fine", "evidence": ["CAD-LVL-01"]}}, {"0": {"reason": "x", "evidence": ["CAD-NOPE-99"]}}):
        r, _ = _vis(tmp_path, issues=[ISSUE], disputes=d, results=res)
        assert r["status"] == "FAIL" and "dispute rejected" in " ".join(r["evidence"]), d


def test_issue_that_contradicts_a_deterministic_pass_without_naming_elements_is_discarded(tmp_path):
    guess = dict(ISSUE, contradicts_deterministic=True, elements=[])
    assert _vis(tmp_path, issues=[guess])[0]["status"] == "PASS"
    named = dict(ISSUE, contradicts_deterministic=True, elements=["35043"])
    assert _vis(tmp_path, issues=[named])[0]["status"] == "FAIL"


def test_render_shows_existing_linework_and_labels_cells_by_their_content_levels(tmp_path):
    cell = el(9, typ="CELL", level="", cellName="480_Vert", childLevels="NSB_P|NSB_P_T", box=(5, 5, 9, 9))
    existing = [_line("1"), el("2", typ="CELL", box=(0, 0, 1, 1))]
    out = visual.render_elements([cell], lambda eid: {"path_vertices": [(0, 0), (10, 10)]}, str(tmp_path / "r.png"), "t",
                                 extent=[0, 0, 20, 20], context_rows=existing + [cell])
    assert os.path.getsize(out) > 1000
    assert [e["elementId"] for e in visual._context_rows(existing + [cell], [cell], [0, 0, 20, 20])] == ["1", "2"]


# ------------------------------------------------------------------ operation records
def test_unrecorded_edit_blocks_unless_independently_checked_or_signed_off():
    c = op_ctx([], changed=["9"])
    r = oc.operations(op_task(), c)
    assert r["status"] == "UNCERTAIN" and any(n["id"] == "OP-01" for n in c["needs_signoff"])
    rel = op_task(expect=[{"id": "R1", "rel": "offset", "a": "9", "b": "8", "value": 12}])
    assert oc.operations(rel, op_ctx([], changed=["9"]))["status"] == "NA"  # a declared geometric relation covers it
    c2 = op_ctx([], changed=["9"])
    c2["signoff"] = {"items": ["OP-01"], "by": "eng", "at": "now"}
    assert oc.operations(op_task(), c2)["status"] == "PASS"


class _FakeBridge:
    def _call_batch_impl(self, ops):
        op, p = ops[0]
        if op == "GET_ELEMENT_VERTICES":
            return [{"status": "OK", "rows": [{"x": 1, "y": 2, "z": 0}]}]
        return [{"status": "OK", "rows": [{"elementId": p["ids"], "type": "CELL", "cellName": FAKE_CELL, "rangeLowX": "0", "rangeLowY": "0",
                                          "rangeHighX": "4", "rangeHighY": "4"}]}]


FAKE_CELL = "480_Vert"


def _record_place_cell(tmp_path, monkeypatch, params):
    monkeypatch.setattr(op_recorder, "LOG", tmp_path / "oplog.jsonl")
    ops = [("PLACE_CELL", params)]
    tok = op_recorder.begin(_FakeBridge(), ops)
    op_recorder.finish(_FakeBridge(), tok, ops, [{"status": "OK", "elementId": "77", "note": "placed"}])
    return [json.loads(x) for x in open(tmp_path / "oplog.jsonl")]


def test_place_cell_is_recorded_and_its_contract_measures_name_and_position(tmp_path, monkeypatch):
    recs = _record_place_cell(tmp_path, monkeypatch, {"cellName": "480_Vert", "ptX": 2.0, "ptY": 2.0, "angleDeg": 0})
    assert recs[0]["op"] == "PLACE_CELL" and recs[0]["new_ids"] == ["77"] and recs[0]["contract"]
    assert operations.validate(recs[0])["status"] == "PASS"
    wrong_name = dict(recs[0], params=dict(recs[0]["params"], cellName="240_1"))
    assert operations.validate(wrong_name)["status"] == "FAIL"
    far = dict(recs[0], params=dict(recs[0]["params"], ptX=500.0))
    assert operations.validate(far)["status"] == "FAIL"


# ------------------------------------------------------------------ coverage: mandatory requirements, gaps, dispositions
ALL_PASS = {k: {"status": "PASS"} for k in ("CAD-LVL-01", "CAD-SYM-01", "CAD-LOC-01", "CAD-DUP-01", "CAD-DEG-01", "CAD-TXT-02",
                                             "CAD-TXT-03", "CAD-TXT-04", "CAD-TXT-05", "CAD-WT-01", "CAD-Z-01")}


def _cov_ctx(**kw):
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))], **kw)
    c["results"] = dict(ALL_PASS)
    return c


def test_mandatory_requirement_whose_check_did_not_pass_blocks():
    c = _cov_ctx()
    c["results"].pop("CAD-TXT-04")
    r = coverage.report({}, c)
    assert r["status"] == "UNCERTAIN" and any("R-TXT-LINE" in e for e in r["evidence"])


def test_all_mandatory_requirements_verified_does_not_block_on_advisory_items():
    r = coverage.report({}, _cov_ctx())
    assert r["status"] == "WARN"  # R-TXT-WORD etc. are listed for the engineer, not blocking


def test_mandatory_manual_requirement_needs_the_engineers_signoff():
    c = _cov_ctx(model_name="727.21.101 1960s CL1 - Profile 1")
    r = coverage.report({}, c)
    assert r["status"] == "UNCERTAIN" and any(n["id"] == "R-VCLR" for n in c["needs_signoff"])
    c2 = _cov_ctx(model_name="727.21.101 1960s CL1 - Profile 1")
    c2["signoff"] = {"items": ["R-VCLR"]}
    assert coverage.report({}, c2)["status"] == "WARN"


def test_candidate_marked_applicable_but_unverifiable_needs_signoff():
    t = {"candidates": [{"cite": "Chapter 99 - 99.9.9 Something (p.1)", "doc": "nysdot/hdm/chapt_99.pdf", "applies": "yes", "reason": "applies here"}]}
    c = _cov_ctx()
    assert coverage.report(t, c)["status"] == "UNCERTAIN" and c["needs_signoff"][0]["id"] == "GAP-0"
    c2 = _cov_ctx()
    c2["signoff"] = {"items": ["GAP-0"]}
    assert coverage.report(t, c2)["status"] == "WARN"


def test_dismissing_a_section_the_registry_ties_to_the_changed_elements_is_a_contradiction():
    cand = {"cite": "Chapter 20 — 20.9.6 Text (p.29)", "doc": "nysdot/hdm/chapt_20.pdf", "applies": "no", "reason": "not relevant to this task"}
    r = coverage.dispositions({"candidates": [cand]}, _cov_ctx())
    assert r["status"] == "UNCERTAIN" and "contradict" in r["observed"]
    assert coverage.dispositions({"candidates": [dict(cand, applies="yes")]}, _cov_ctx())["status"] == "PASS"
    assert coverage.dispositions({"candidates": [dict(cand, cite="Chapter 20 — 20.1 Intro (p.2)")]}, _cov_ctx())["status"] == "PASS"


def test_registry_flags_and_new_kinds_load():
    reg = coverage.registry()
    mand = {r["id"] for k in reg["kinds"] for r in k["requirements"] if r.get("mandatory")}
    assert {"R-LVL", "R-TXT-LINE", "R-VCLR"} <= mand and "R-Z" not in mand  # Z is NA in 3D models, so it cannot be mandatory
    assert {"K-SIGN", "K-STRIPING"} <= {k["id"] for k in reg["kinds"]}


# ------------------------------------------------------------------ catalogs and traceability
def test_failure_mode_ids_are_unique():
    ids = [r["id"] for r in yaml.safe_load(open(os.path.join(ROOT, "qa", "failure_modes.yaml"), encoding="utf-8"))]
    assert len(ids) == len(set(ids))


def test_provenance_names_the_gate_standards_editions_and_state(tmp_path, monkeypatch):
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path))
    t = {"id": "T1", "candidates": [{"doc": "nysdot/hdm/chapt_20.pdf"}], "citations": []}
    c = {"run_model": "f|m", "state_fp": "abc", "oplog": [1, 2], "signoff": {"items": ["VIS-01"], "by": "eng"},
         "coverage": {"registry_version": 2, "requirements": [{"requirement": "R-LVL", "kinds": ["K-ALL"], "check": "CAD-LVL-01", "status": "verified"}]}}
    pv = qa_gate._provenance(t, c)
    assert pv["state_fingerprint"] == "abc" and pv["model"] == "f|m" and pv["oplog_records"] == 2
    assert pv["criteria_sha"] and pv["applicability_sha"] and pv["standards_registry_sha"]
    assert "nysdot/hdm/chapt_20.pdf" in pv["standards_documents"]
    assert pv["requirement_rationale"][0]["kinds"] == ["K-ALL"] and pv["signoff"]["items"] == ["VIS-01"]


def test_a_placed_polylines_length_counts_as_a_number_used_by_the_operation():
    rec = {"op": "PLACE_POLYLINE", "params": {"verticesTSV": "4700,-300,0|4720,-300,0"}}
    assert any(abs(n - 20.0) < 1e-9 for n in operations.numbers_in_params(rec))


def test_state_read_sends_one_vertex_request_per_element(monkeypatch):
    """Multi-op bridge batches return vertex rows attributed to the wrong elements (live 2026-10-23), so the
    vertex hashes behind the fingerprint must come from single-op requests."""
    class Bridge:
        def _call_batch_impl(self, ops):
            assert len(ops) == 1, "vertex reads must not be batched"
            return [{"status": "OK", "rows": [{"x": "1", "y": "2", "z": "0"}, {"x": "3", "y": "4", "z": "0"}]}]

    class Ops:
        _bridge = Bridge()

        @staticmethod
        def get_element_cad_props(min_id=0, ids="", max_rows=0):
            return {"truncated": False, "rows": [_line("1"), _line("2"), el("3", typ="CELL")]}

    st = qa_gate.read_state(Ops)
    assert set(st["vh"]) == {"1", "2"} and not st["vh_capped"]
