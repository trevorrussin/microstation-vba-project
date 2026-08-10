"""Geometry-faithful scorecard + sandbox + harness history P0."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import chat_history
import sheet_geometry_faithful as gf
import sheet_sandbox
import sheet_scorecard


def test_geometry_faithful_mid_drift():
    compiled = {
        "plan": {
            "planByAlign": {
                "1": [{
                    "kind": "dimension",
                    "primitiveId": "1:buffer:dim",
                    "tip1": [0.0, 0.0],
                    "tip2": [100.0, 0.0],
                    "text": "100",
                }],
            }
        }
    }
    # mid should be (50,0); registry drifted to (80,0)
    rows = [{
        "primitiveId": "1:buffer:dim",
        "kind": "dimension",
        "elementIds": ["1"],
        "midX": 80.0,
        "midY": 0.0,
    }]
    fails = gf.check_geometry_faithfulness(compiled, rows)
    assert fails and "drifted" in fails[0]


def test_duplicate_signs():
    rows = [
        {"kind": "sign", "primitiveId": "1:W20-05RA:sign", "alignIdx": 1,
         "x": 10.0, "y": 20.0, "elementIds": ["a"]},
        {"kind": "sign", "primitiveId": "1:W20-05RA:sign", "alignIdx": 1,
         "x": 10.2, "y": 20.1, "elementIds": ["b"]},
    ]
    fails = gf.check_duplicate_signs(rows)
    assert fails and "duplicate sign" in fails[0]


def test_scorecard_includes_faithful_failures():
    compiled = {
        "plan": {"planByAlign": {}, "channelizingByAlign": {},
                 "symbolsByAlign": {}, "hatch": []},
        "counts": {},
        "gateFailures": [],
    }
    # Empty expectations → pass presence; add flood via placed
    rows = [
        {"kind": "dimension", "primitiveId": f"extra:{i}", "elementIds": [str(i)]}
        for i in range(5)
    ]
    # expected byKind empty — flood check only when expected>0
    sc = sheet_scorecard.build_placement_scorecard(compiled, registry_rows=rows)
    assert sc["geometryFaithful"] is True


def test_auto_visual_missing_pv():
    compiled = {
        "plan": {
            "symbolsByAlign": {
                "1": [{"kind": "protectiveVehicle", "primitiveId": "1:pv",
                       "x": 1, "y": 2, "id": "pv1"}],
            },
            "planByAlign": {},
        }
    }
    fails = gf.check_automated_visual_rules(compiled, registry_rows=[])
    assert any("protectiveVehicle" in f for f in fails)


def test_sandbox_offset_and_keep(tmp_path, monkeypatch):
    monkeypatch.setattr(sheet_sandbox, "STATE_PATH", tmp_path / "sandbox-state.json")
    monkeypatch.setattr(sheet_sandbox, "CHECKPOINT_DIR", tmp_path / "cps")
    monkeypatch.setattr(sheet_sandbox, "_BRIDGE", tmp_path)
    out = sheet_sandbox.begin_sandbox(
        upstream_edge=[100.0, 200.0, 0.0],
        downstream_edge=[50.0, 200.0, 0.0],
        offset_y_ft=2000.0,
        sheet_num="619-311",
    )
    assert out["upstream_edge"][1] == 2200.0
    assert out["sandbox"]["status"] == "active"
    kept = sheet_sandbox.keep_sandbox()
    assert kept["sandbox"]["status"] == "kept"


def test_harness_preflight_clears_orphan():
    messages = [
        {"role": "assistant", "content": [
            {"type": "tool_use", "id": "toolu_x", "name": "x", "input": {}},
        ]},
    ]
    # Repair drops unanswered tool_use; empty history is safe.
    chat_history.harness_preflight_or_clear(messages)
    assert messages == []
    # Inject a still-broken shape repair might miss: tool_result with no prior use
    messages[:] = [
        {"role": "user", "content": "hi"},
        {"role": "user", "content": [
            {"type": "tool_result", "tool_use_id": "toolu_orphan", "content": "x"},
        ]},
    ]
    # Repair drops the orphan tool_result user msg
    chat_history._repair_tool_pairing(messages)
    leftover = chat_history.harness_history_issues(messages)
    # After repair should be clean (just "hi")
    assert leftover == []
    assert messages[0]["content"] == "hi"


def test_plan_sheets_have_build_guides():
    spec_dir = ROOT / "Data" / "sheet-specs"
    guides = list(spec_dir.glob("619-*.build.md"))
    assert len(guides) >= 60
    assert (spec_dir / "619-311.build.md").is_file()
    # Spot-check a generated sibling
    assert (spec_dir / "619-302.build.md").is_file()
    raw = json.loads((spec_dir / "619-302.json").read_text(encoding="utf-8"))
    assert raw["sheet"].get("buildGuide") == "619-302.build.md"
