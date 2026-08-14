"""Build ledger + overlap classify + Tier 1 scorecard stacks."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import build_ledger as bl  # noqa: E402
import build_overlap as ov  # noqa: E402
import sheet_scorecard as sc  # noqa: E402


def test_ledger_append_and_retention(tmp_path, monkeypatch):
    path = tmp_path / "ledger.jsonl"
    monkeypatch.setattr(bl, "LEDGER_PATH", path)
    monkeypatch.setattr(bl, "MAX_ROWS", 3)
    for i in range(5):
        bl.append_build(sheet_num="619-311", origin=[float(i), 0.0],
                        path_vertices=[[float(i), 0], [float(i) + 10, 0]])
    rows = bl.load_builds()
    assert len(rows) == 3
    assert rows[0]["origin"][0] == 2.0


def test_tier1_duplicate_hash():
    rows = [
        {"elementId": "1", "type": "TEXT", "cx": 10.04, "cy": 20.01, "w": 12.0, "h": 2.0, "text": "BUFFER SPACE"},
        {"elementId": "2", "type": "TEXT", "cx": 10.02, "cy": 20.00, "w": 12.02, "h": 2.01, "text": "BUFFER SPACE"},
        {"elementId": "3", "type": "TEXT", "cx": 50.0, "cy": 20.0, "w": 12.0, "h": 2.0, "text": "LANE TAPER"},
    ]
    dups = ov.tier1_duplicates(rows)
    assert len(dups) == 1
    assert dups[0]["count"] == 2


def test_classify_same_origin_rebuild(tmp_path, monkeypatch):
    path = tmp_path / "ledger.jsonl"
    monkeypatch.setattr(bl, "LEDGER_PATH", path)
    bl.append_build(
        sheet_num="619-311", origin=[76000.0, 288000.0],
        path_vertices=[[76000, 288000], [78600, 288000], [78600, 290000]],
        bbox={"lowX": 75900, "lowY": 287900, "highX": 78700, "highY": 290100},
        lateral_half_width=20.0,
    )
    r = ov.classify(
        sheet_num="619-311",
        origin=[76010.0, 288005.0],
        path_vertices=[[76000, 288000], [78600, 288000]],
        lateral_half_width=20.0,
        ledger_rows=bl.load_builds(),
        model_rows=[],
    )
    assert r["verdict"] == "rebuild_same_origin"
    assert r["blocking"] is False
    assert r["nextTool"] == "clear_plan_elements"


def test_classify_other_sheet_path_conflict(tmp_path, monkeypatch):
    path = tmp_path / "ledger.jsonl"
    monkeypatch.setattr(bl, "LEDGER_PATH", path)
    corridor = [[0.0, 0.0], [500.0, 0.0]]
    bl.append_build(
        sheet_num="619-302", origin=[0.0, 0.0],
        path_vertices=corridor, lateral_half_width=40.0,
        bbox={"lowX": -80, "lowY": -80, "highX": 580, "highY": 80},
    )
    r = ov.classify(
        sheet_num="619-311",
        origin=[2000.0, 2000.0],
        path_vertices=corridor,
        lateral_half_width=20.0,
        ledger_rows=bl.load_builds(),
        model_rows=[],
    )
    assert r["verdict"] == "collision_other_sheet"
    assert r["conflicts"]


def test_scorecard_fails_on_model_stacks():
    compiled = {
        "gateFailures": [],
        "counts": {},
        "plan": {"planByAlign": {}, "channelizingByAlign": {},
                 "symbolsByAlign": {}, "hatch": []},
    }
    model = [
        {"elementId": "10", "type": "TEXT", "cx": 1.0, "cy": 1.0, "w": 8.0, "h": 2.0, "text": "ROLL AHEAD DISTANCE"},
        {"elementId": "11", "type": "TEXT", "cx": 1.0, "cy": 1.0, "w": 8.0, "h": 2.0, "text": "ROLL AHEAD DISTANCE"},
        {"elementId": "12", "type": "TEXT", "cx": 1.0, "cy": 1.0, "w": 8.0, "h": 2.0, "text": "ROLL AHEAD DISTANCE"},
        {"elementId": "13", "type": "TEXT", "cx": 1.0, "cy": 1.0, "w": 8.0, "h": 2.0, "text": "ROLL AHEAD DISTANCE"},
    ]
    bad = sc.build_placement_scorecard(compiled, registry_rows=[], model_rows=model)
    assert bad["passed"] is False
    assert any("stacked 4x TEXT" in f for f in bad["failures"])
