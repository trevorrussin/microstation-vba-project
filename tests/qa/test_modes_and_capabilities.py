"""Production vs capability-development modes, and the capability catalog (2026-10-24, after the circles run)."""
from __future__ import annotations

import json
import os
import sys
import time
import types

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import capabilities  # noqa: E402
import modes  # noqa: E402
import qa_gate  # noqa: E402
import qa_mode_hook  # noqa: E402
import task_guard  # noqa: E402


def _task(tasks_dir, tid, **kw):
    d = tasks_dir / tid
    d.mkdir(parents=True)
    t = {"id": tid, "type": "cad.drafting", "status": "in_progress", "created": time.time(), "candidates": []}
    t.update(kw)
    json.dump(t, open(d / "task.json", "w"))
    return t


@pytest.fixture
def tasks(tmp_path, monkeypatch):
    d = tmp_path / "tasks"
    d.mkdir()
    monkeypatch.setattr(modes, "TASKS", d)
    monkeypatch.setattr(task_guard, "TASKS", d)
    monkeypatch.setattr(task_guard, "ATTEMPTS", tmp_path / "attempts.jsonl")
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    return d


# ------------------------------------------------------------------ what counts as tooling
@pytest.mark.parametrize("path", ["mcp-server/wztc_ops.py", "Modules/WZTCBridge.bas", "ord-addin/src/Bridge.cs", "Data/command-registry.tsv",
                                  "Data/workflow-contracts/profile-design.yaml", "qa/criteria.yaml", "tests/qa/test_x.py", "scripts/qa_gate.py",
                                  "Data/diagnostic-ladders/a.yaml", ".claude/settings.json"])
def test_tooling_paths_are_protected(path):
    assert modes.is_protected(path) and modes.is_protected(os.path.join(ROOT, path))


@pytest.mark.parametrize("path", ["dev-notes/agent-log.md", "Data/experience/experience.jsonl", "scripts/_tmp_my_scratch.py", "outputs/x.json", "CLAUDE.md"])
def test_ordinary_paths_are_not_protected(path):
    assert not modes.is_protected(path)


# ------------------------------------------------------------------ the rule applies only while a production task is open
def test_no_open_task_means_no_restriction(tasks):
    assert modes.assess_path("mcp-server/wztc_ops.py") == ""
    assert modes.assess_command("python mcp-server/hot_reload.py Modules/X.bas") == ""


def test_production_task_refuses_tooling_edits_and_loaders(tasks):
    _task(tasks, "P1")  # no mode field = production
    msg = modes.assess_path("mcp-server/wztc_ops.py")
    assert "capability development" in msg and "--mode development" in msg and "P1" in msg
    assert modes.assess_command("python mcp-server/hot_reload.py Modules/WZTCBridge.bas")
    assert modes.assess_command("powershell ord-addin\\build.ps1")
    assert modes.assess_command("echo x >> mcp-server/wztc_ops.py")
    assert modes.assess_command("sed -i s/a/b/ Modules/WZTCExec.bas")


def test_production_task_still_allows_reading_and_ordinary_work(tasks):
    _task(tasks, "P1")
    assert modes.assess_command("grep -n cif_exec mcp-server/wztc_ops.py") == ""
    assert modes.assess_command("python scripts/qa_gate.py status") == ""
    assert modes.assess_path("dev-notes/agent-log.md") == ""
    assert modes.assess_path("Data/experience/experience.jsonl") == ""


def test_development_task_lifts_the_restriction(tasks):
    _task(tasks, "D1", mode="development", scratch="none")
    assert modes.assess_path("mcp-server/wztc_ops.py") == ""
    assert modes.assess_command("python mcp-server/hot_reload.py Modules/X.bas") == ""


def test_a_production_task_still_restricts_while_a_development_task_is_also_open(tasks):
    _task(tasks, "D1", mode="development", scratch="none")
    _task(tasks, "P1")
    assert modes.assess_path("mcp-server/wztc_ops.py")


def test_expired_or_closed_production_task_does_not_restrict(tasks):
    _task(tasks, "P1", status="closed")
    _task(tasks, "P2", created=time.time() - 13 * 3600)
    assert modes.assess_path("mcp-server/wztc_ops.py") == ""


def test_hook_maps_each_tool_to_the_right_check(tasks):
    _task(tasks, "P1")
    assert qa_mode_hook.assess("Edit", {"file_path": os.path.join(ROOT, "mcp-server", "task_guard.py")})
    assert qa_mode_hook.assess("Write", {"file_path": "Modules/New.bas"})
    assert qa_mode_hook.assess("Bash", {"command": "python mcp-server/hot_reload.py Modules/A.bas"})
    assert qa_mode_hook.assess("Edit", {"file_path": "dev-notes/agent-log.md"}) == ""
    assert qa_mode_hook.assess("Read", {"file_path": "mcp-server/task_guard.py"}) == ""


# ------------------------------------------------------------------ development tasks never edit an unnamed drawing
def test_code_only_development_task_refuses_every_drawing_edit(tasks):
    _task(tasks, "D1", mode="development", scratch="none")
    task_guard.require_open_task(["GET_ELEMENT_VERTICES"])
    with pytest.raises(task_guard.QAGateRequired) as ex:
        task_guard.require_open_task(["MOVE_ELEMENT"])
    assert "code-only" in str(ex.value)


def test_development_task_on_a_named_scratch_file_may_edit(tasks):
    _task(tasks, "D1", mode="development", scratch="scratch.dgn")
    task_guard.require_open_task(["MOVE_ELEMENT"])


def _args(**kw):
    base = dict(type="generic", desc="Investigate per-view level display on library levels", id="M1", files=[], extent=None, ids=None,
                text_heights=None, is3d=False, may_modify=False, no_standards=True, no_snapshot=True, mode="production", scratch=None, mode_reason=None)
    base.update(kw)
    return types.SimpleNamespace(**base)


@pytest.fixture
def gate(tmp_path, monkeypatch):
    monkeypatch.setattr(qa_gate, "TASKS", str(tmp_path / "gt"))
    os.makedirs(qa_gate.TASKS)
    active = {"fileName": "X727.11_fea_rwy_alt_a.dgn"}
    monkeypatch.setattr(qa_gate, "_ops", lambda: types.SimpleNamespace(describe_drawing_state=lambda: dict(active)))
    return active


def test_start_defaults_to_production_and_says_what_that_forbids(gate, capsys):
    qa_gate.cmd_start(_args())
    out = capsys.readouterr().out
    assert "production mode" in out and "Editing the tooling" in out
    assert qa_gate.load_task("M1")["mode"] == "production"


def test_development_start_needs_a_scratch_and_a_reason(gate):
    with pytest.raises(SystemExit):
        qa_gate.cmd_start(_args(mode="development"))
    with pytest.raises(SystemExit):
        qa_gate.cmd_start(_args(mode="development", scratch="none", mode_reason="short"))


def test_development_start_refuses_when_the_active_file_is_not_the_named_scratch(gate):
    with pytest.raises(SystemExit) as ex:
        qa_gate.cmd_start(_args(mode="development", scratch="scratch.dgn", mode_reason="probe per-view level display API on a scratch file"))
    assert "X727.11_fea_rwy_alt_a.dgn" in str(ex.value)


def test_development_start_accepts_the_matching_scratch_and_code_only(gate):
    gate["fileName"] = "Scratch.dgn"
    qa_gate.cmd_start(_args(mode="development", scratch="scratch.dgn", mode_reason="probe per-view level display API on a scratch file"))
    t = qa_gate.load_task("M1")
    assert (t["mode"], t["scratch"]) == ("development", "scratch.dgn")
    qa_gate.cmd_start(_args(id="M2", mode="development", scratch="none", mode_reason="write the capability search and its tests (no drawing)"))
    assert qa_gate.load_task("M2")["scratch"] == "none"


# ------------------------------------------------------------------ capability catalog
def test_catalog_is_built_from_what_exists():
    entries, _ = capabilities.catalog()
    kinds = {e["kind"] for e in entries}
    assert kinds == {"workflow", "procedure", "tool", "keyin", "diagnosis", "gui"}
    ids = {e["id"] for e in entries}
    assert {"workflow:profile-design", "tool:place_sign", "tool:move_element", "diagnosis:unwanted-or-missing-graphics",
            "procedure:references-levels-models-measuring", "keyin:ACTIVE_LEVEL"} <= ids
    assert not any(i.startswith("keyin:MOVE_ELEMENT") and i == "keyin:MOVE_ELEMENT" for i in ids), "direct_api rows are tools, not key-ins"


def test_profile_request_finds_the_workflow_first_and_marks_engineer_steps():
    r = capabilities.find("Create a proposed vertical alignment profile using these PVI elevations and curve lengths")
    assert r["status"] == "FOUND" and r["matches"][0]["id"] == "workflow:profile-design"
    m = r["matches"][0]
    assert m["status"] == "partial" and m["engineer_steps"] and m["production_ok"]
    assert m["executors"][1].startswith("create_profile_exact"), "the contract's override beats the kind default"


def test_symptom_request_finds_the_ladder():
    r = capabilities.find("turn off the yellow circles in the plan view")
    assert r["matches"][0]["id"] == "diagnosis:unwanted-or-missing-graphics"


def test_simple_edit_request_finds_the_bridge_tool():
    r = capabilities.find("move this element 5 ft to the right")
    assert "tool:move_element" in [m["id"] for m in r["matches"]]


def test_blocked_and_unverified_rows_are_not_usable_in_production():
    r = capabilities.find("export to IFC")
    blocked = [m for m in r["matches"] if m["id"] == "keyin:IFC_EXPORT"]
    assert blocked and blocked[0]["status"] == "blocked" and blocked[0]["production_ok"] is False
    assert "NOT for production" in capabilities.format_result(r)


def test_unrelated_request_is_unsupported_with_mode_specific_guidance():
    p = capabilities.find("design the drainage network for the new culvert", mode="production")
    d = capabilities.find("design the drainage network for the new culvert", mode="development")
    assert p["status"] == d["status"] == "UNSUPPORTED" and p["nearest"]
    assert "Stop" in p["guidance"] and "scratch file" in d["guidance"] and p["guidance"] != d["guidance"]
    assert "NO SUPPORTED CAPABILITY" in capabilities.format_result(p)


def test_partial_word_matches_are_flagged_low_confidence():
    r = capabilities.find("write a poem")
    assert r["status"] != "FOUND" or all(m["confidence"] == "low" for m in r["matches"])


def test_policy_executor_lists_are_plain_strings():
    """A ': ' inside an unquoted YAML flow list silently turns an item into a mapping (found while building this)."""
    pol = capabilities._policy()
    for k in list(pol["kinds"].values()) + list(pol["overrides"].values()):
        assert all(isinstance(x, str) for x in k["executors"]), k
