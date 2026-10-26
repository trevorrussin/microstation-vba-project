"""Retrieval receipts: a playbook must be READ (not just named) before edits or a chat answer; plans are diffed
against the playbook. Uses the real profile-cutting-workflow playbook."""
from __future__ import annotations

import json
import os
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import retrieval  # noqa: E402
import task_guard  # noqa: E402
import tutorial_kb  # noqa: E402

TOPIC = "profile-cutting-workflow"


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(retrieval, "RECEIPTS", str(tmp_path / "receipts.jsonl"))
    monkeypatch.setattr(retrieval, "PENDING", str(tmp_path / "pending.json"))
    monkeypatch.setattr(task_guard, "TASKS", tmp_path / "tasks")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    return tmp_path


def _task(tmp, applies=None):
    d = tmp / "tasks" / "t1"
    d.mkdir(parents=True)
    t = {"id": "t1", "type": "cad.drafting", "status": "in_progress", "created": time.time(),
         "candidates": [{"cite": "PLAYBOOK x", "doc": f"playbook:{TOPIC}", "applies": applies, "reason": "n/a here"}]}
    (d / "task.json").write_text(json.dumps(t))
    task_guard._cache.update(at=0.0, open=None)


# A/B: the edit is refused until the playbook is read, then allowed
def test_edit_refused_until_playbook_read(env):
    _task(env)
    with pytest.raises(task_guard.PlaybookNotRead) as e:
        task_guard.require_open_task(["PLACE_POLYLINE"])
    assert TOPIC in str(e.value)
    tutorial_kb.get_playbook(TOPIC)
    task_guard._cache.update(at=0.0, open=None)
    task_guard.require_open_task(["PLACE_POLYLINE"])


def test_reads_are_never_blocked(env):
    _task(env)
    task_guard.require_open_task(["GET_ELEMENTS_RANGE"])


def test_waived_playbook_does_not_block(env):
    _task(env, applies="no")
    task_guard.require_open_task(["PLACE_POLYLINE"])


# F: retrieval fails -> no receipt, so the gate stays closed
def test_failed_lookup_writes_no_receipt(env):
    out = tutorial_kb.get_playbook("no-such-playbook")
    assert out.startswith("no playbook")
    assert not retrieval.has_fresh_receipt("no-such-playbook")
    assert not os.path.exists(retrieval.RECEIPTS)


def test_receipt_goes_stale_when_playbook_changes(env, monkeypatch):
    tutorial_kb.get_playbook(TOPIC)
    assert retrieval.has_fresh_receipt(TOPIC)
    monkeypatch.setattr(retrieval, "playbook_hash", lambda t: "changed")
    assert not retrieval.has_fresh_receipt(TOPIC)


# chat answers: prompt matches -> pending -> Stop holds the turn until read
def test_stop_holds_turn_until_playbook_read(env):
    unread = retrieval.unread_for_text("make profile cut sheets from sta 10 to 30")
    assert TOPIC in [p["topic"] for p in unread]
    retrieval.set_pending([p["topic"] for p in unread])
    assert TOPIC in retrieval.stop_message()
    assert {TOPIC, "profile-sheet-annotation"} <= {p["topic"] for p in unread}  # related playbook is required too
    for p in unread:
        tutorial_kb.get_playbook(p["topic"])
    assert retrieval.stop_message() == ""


def test_stop_gives_up_after_limit_and_says_so(env):
    retrieval.set_pending([TOPIC])
    for _ in range(retrieval.MAX_PENDING_BLOCKS):
        assert "not consulted" in retrieval.stop_message()
    assert "NOT backed" in retrieval.stop_message()
    assert retrieval.stop_message() == ""


def test_no_nag_once_read(env):
    tutorial_kb.get_playbook(TOPIC)
    assert TOPIC not in [p["topic"] for p in retrieval.unread_for_text("cut a profile along the alignment")]


# E: plan omits a required setting / invents a value
FULL_PLAN = ("N5 Place Named Boundary Civil Profile, Detail Scale 1\" = 40', Drawing Seed Profile Only, "
             "Method Station Limits, Vertical Exaggeration 5, Available Profile Height 50, Top Clearance 1, "
             "Bottom Clearance 1, Elevation Datum Spacing 5, Station Datum Spacing 100, Profile Shifts Datum Stations, "
             "tick Use Active Vertical")


def test_diff_flags_omitted_setting(env):
    r = retrieval.check_plan_against_playbook(TOPIC, "N5 Place Named Boundary, Civil Profile, Vertical Exaggeration 5")
    row = next(x for x in r["steps_with_omissions"] if x["step"] == "N5")
    assert any("Available Profile Height" in m or m == "50" for m in row["missing"])
    assert r["status"] == "REVIEW"


def test_diff_flags_invented_number(env):
    r = retrieval.check_plan_against_playbook(TOPIC, FULL_PLAN + ", Vertical Exaggeration 7.25")
    assert "7.25" in r["values_not_in_playbook"]


def test_diff_full_plan_has_no_omissions_for_n5(env):
    r = retrieval.check_plan_against_playbook(TOPIC, FULL_PLAN)
    missing = [m for x in r["steps_with_omissions"] if x["step"] == "N5" for m in x["missing"]]
    for setting in ("Vertical Exaggeration", "Available Profile Height", "Top Clearance", "Elevation Datum Spacing",
                    "Station Datum Spacing", "Use Active Vertical", "Method Station Limits", "50", "100"):
        assert setting not in missing
    assert not any("65+77" in m for m in missing)  # example-only station values are not required


def test_diff_unknown_playbook(env):
    assert retrieval.check_plan_against_playbook("nope", "x")["status"] == "NO_PLAYBOOK"


# ---- other sources: standards, written tutorial sources, searches
STD = "nysdot/hdm/chapt_20.pdf"
DOC = "Creating Profiles and Sheets from Alignment 7.pdf"


def _task2(tmp, cands=(), sources=()):
    d = tmp / "tasks" / "t2"
    d.mkdir(parents=True, exist_ok=True)
    t = {"id": "t2", "type": "cad.drafting", "status": "in_progress", "created": time.time(),
         "candidates": list(cands), "sources": list(sources)}
    (d / "task.json").write_text(json.dumps(t))
    task_guard._cache.update(at=0.0, open=None)
    return t


def _std(applies=None):
    return {"cite": "Chapter 20 20.7", "doc": STD, "applies": applies, "reason": "x" * 12 if applies else ""}


def _src(applies=None):
    return {"cite": "SOURCE " + DOC, "doc": "tutorial:" + DOC, "applies": applies, "reason": ""}


def test_undispositioned_standard_blocks_edit(env):
    _task2(env, cands=[_std()])
    with pytest.raises(task_guard.PlaybookNotRead) as e:
        task_guard.require_open_task(["PLACE_POLYLINE"])
    assert "disposition standard" in str(e.value)


def test_standard_yes_needs_the_read(env):
    t = _task2(env, cands=[_std("yes")])
    assert retrieval.gaps(t) == [f"read standard {STD}"]
    retrieval.record_read("standard", STD)  # what get_standard_section / read_standard_pages do on non-empty text
    assert retrieval.gaps(t) == []


def test_standard_no_needs_only_reason(env):
    assert retrieval.gaps(_task2(env, cands=[_std("no")])) == []


def test_standard_receipt_matches_by_file_name(env):
    retrieval.record_read("standard", "hdm/chapt_20.pdf")
    assert retrieval.has_read("standard", STD)


def test_written_source_must_be_read(env):
    t = _task2(env, sources=[_src("yes")])
    assert retrieval.gaps(t) == [f"read tutorial {DOC}"]
    assert tutorial_kb.read_source(DOC, "1", "1").strip()
    assert retrieval.gaps(t) == []


def test_unmatched_source_name_writes_no_receipt(env):
    assert tutorial_kb.read_source("no such file", "1", "1").startswith("no tutorial source")
    assert not os.path.exists(retrieval.RECEIPTS)


def test_playbook_sources_resolve_to_written_docs(env):
    keys = [d["key"] for d in retrieval.doc_sources_for([TOPIC])]
    assert DOC in keys
    assert not any(k.lower().endswith(".mp4") for k in keys)


def test_dispose_yes_refused_until_read(env, monkeypatch, capsys):
    import qa_gate
    t = _task2(env, cands=[_std()])
    monkeypatch.setattr(qa_gate, "load_task", lambda tid=None: json.loads((env / "tasks" / "t2" / "task.json").read_text()))
    monkeypatch.setattr(qa_gate, "save_task", lambda tt: (env / "tasks" / "t2" / "task.json").write_text(json.dumps(tt)))
    a = type("A", (), {"task": "t2", "cite": "Chapter 20", "applies": "yes", "reason": "governs text and levels"})()
    with pytest.raises(SystemExit) as e:
        qa_gate.cmd_dispose(a)
    assert "needs the source read" in str(e.value)
    retrieval.record_read("standard", STD)
    assert qa_gate.cmd_dispose(a) == 0


def test_chat_turn_waits_for_a_search_too(env):
    tutorial_kb.get_playbook(TOPIC)
    retrieval.set_pending([], need_search=True)
    assert "search_standards" in retrieval.stop_message()
    retrieval.record_search("standards", "profile sheet scale")
    assert retrieval.stop_message() == ""


def test_search_before_the_prompt_does_not_count(env):
    retrieval.record_search("standards", "old query")
    time.sleep(0.01)
    retrieval.set_pending([], need_search=True)
    assert retrieval.stop_message() != ""
