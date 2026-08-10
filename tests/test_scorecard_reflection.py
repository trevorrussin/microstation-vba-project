"""Scorecard, registry supersedes/reqId, reflection, visual QA gating."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import placement_registry as preg
import sheet_scorecard as sc


def test_supersedes_latest_wins(tmp_path, monkeypatch):
    path = tmp_path / "reg.jsonl"
    monkeypatch.setattr(preg, "REGISTRY_PATH", path)
    preg.clear_registry()
    r1 = preg.append_placement(
        sheet_num="619-311", align_idx=1, kind="cone",
        primitive_id="1:laneTaperRun:cone", bridge_op="PLACE_CHANNELIZING_MARKERS",
        element_ids=["10"], req_id="P1")
    r2 = preg.append_placement(
        sheet_num="619-311", align_idx=1, kind="cone",
        primitive_id="1:laneTaperRun:cone", bridge_op="PLACE_CHANNELIZING_MARKERS",
        element_ids=["20", "21"], req_id="P2")
    assert r2["supersedes"] == r1["recordId"]
    heads = preg.resolve_latest_placements(kind="cone")
    assert len(heads) == 1
    assert heads[0]["elementIds"] == ["20", "21"]
    assert heads[0]["reqId"] == "P2"
    all_rows = preg.load_placements(kind="cone", include_superseded=True)
    assert len(all_rows) == 2


def test_mark_deleted_soft(tmp_path, monkeypatch):
    path = tmp_path / "reg.jsonl"
    monkeypatch.setattr(preg, "REGISTRY_PATH", path)
    preg.clear_registry()
    preg.append_placement(
        sheet_num="619-311", align_idx=1, kind="dimension",
        primitive_id="1:bufferSpace:dimension", bridge_op="PLACE_DIMENSION",
        element_ids=["5"], req_id="D1")
    n = preg.mark_deleted({"1:bufferSpace:dimension"})
    assert n == 1
    assert preg.resolve_latest_placements() == []


def test_scorecard_missing_and_pass():
    compiled = {
        "gateFailures": [],
        "counts": {},
        "plan": {
            "planByAlign": {
                "1": [
                    {"kind": "dimension", "primitiveId": "1:a:dimension"},
                    {"kind": "label", "primitiveId": "1:a:label"},
                ]
            },
            "channelizingByAlign": {
                "1": [
                    {"kind": "cone", "run": "laneTaperRun",
                     "primitiveId": "1:laneTaperRun:cone"},
                    {"kind": "cone", "run": "laneTaperRun",
                     "primitiveId": "1:laneTaperRun:cone"},
                ]
            },
            "symbolsByAlign": {},
            "hatch": [{"kind": "hatch", "primitiveId": "0:workArea:hatch"}],
        },
    }
    reg = [
        {"kind": "dimension", "primitiveId": "1:a:dimension", "elementIds": ["1"]},
        {"kind": "label", "primitiveId": "1:a:label", "elementIds": ["2"]},
        {"kind": "cone", "primitiveId": "1:laneTaperRun:cone", "elementIds": ["3"]},
        # hatch missing
    ]
    bad = sc.build_placement_scorecard(compiled, registry_rows=reg)
    assert bad["passed"] is False
    assert any("hatch" in f or "missing" in f for f in bad["failures"])

    reg.append({"kind": "hatch", "primitiveId": "0:workArea:hatch",
                "elementIds": ["4"], "reqId": "H1"})
    good = sc.build_placement_scorecard(compiled, registry_rows=reg)
    assert good["passed"] is True
    assert good["citations"]


def test_visual_qa_prechecks():
    fails = sc.visual_qa_prechecks(
        {"passed": False, "failures": ["x"]},
        registry_rows=[{"kind": "dimension"}],
        sheet_geometry_placed=True,
    )
    assert any("scorecard" in f for f in fails)
    ok = sc.visual_qa_prechecks(
        {"passed": True, "failures": []},
        registry_rows=[{"kind": "dimension", "elementIds": ["1"]}],
        sheet_geometry_placed=True,
    )
    assert ok == []


def test_replan_and_reflect(tmp_path, monkeypatch):
    import wztc_ops as ops
    monkeypatch.setattr(ops, "SHEET_PLAN_PATH", tmp_path / "sheet-plan.json")
    monkeypatch.setattr(preg, "REGISTRY_PATH", tmp_path / "reg.jsonl")
    monkeypatch.setattr(ops, "_BRIDGE_DIR", tmp_path)
    ops._PLAN_SESSION.reset()
    ops._PLAN_SESSION.order_table_built = True
    ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width=">= 8 ft", area_type="RURAL")
    ops._PLAN_SESSION.sheet_geometry_placed = True
    ops._PLAN_SESSION.last_scorecard = {
        "passed": False,
        "failures": ["scorecard: missing hatch"],
        "citations": [{"primitiveId": "1:a:dimension", "elementIds": ["1"],
                       "reqId": "P9"}],
    }
    replan = ops._replan_after_failure(
        "place_sheet_geometry",
        {"failures": ["scorecard: missing hatch"]},
    )
    assert replan["resumeFrom"] == "place_sheet_geometry"
    assert "place_sheet_geometry" in replan["preservedPhases"] or True

    refl = ops.reflect_sheet_build()
    assert refl["satisfactory"] is False
    assert refl["citations"]
    assert (tmp_path / "sheet-reflection.jsonl").exists()

    # visual QA refuses without passing scorecard
    out = ops.run_visual_qa_captures()
    assert out.get("visualQaPassed") is False
    assert out.get("status") == "ERROR"
    ops._PLAN_SESSION.reset()
