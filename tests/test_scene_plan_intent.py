"""Intent fields, durable plan save/load, placement registry."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import placement_registry as preg
import sheet_compile as sc
import wztc_ops as ops


def test_annotation_style_defaults():
    style = sc.annotation_style({})
    assert style["dimensionText"] == "lengthOnly"
    assert style["featureLabel"] == "nameOnly"
    assert style["overlayDimSide"] == "opposite"
    assert style["offsetsFt"]["dimOutward"] == 15.0


def test_annotation_style_override():
    style = sc.annotation_style({
        "annotationStyle": {
            "overlayDimSide": "same",
            "offsetsFt": {"symbolLabel": 30},
        }
    })
    assert style["overlayDimSide"] == "same"
    assert style["offsetsFt"]["symbolLabel"] == 30.0
    assert style["offsetsFt"]["dimOutward"] == 15.0  # untouched default


def test_channelizing_representation_default_and_override():
    assert sc.channelizing_representation({})["mode"] == "markers"
    rep = sc.channelizing_representation({
        "representation": {"mode": "markers", "markerHalfSizeFt": 2.0}
    })
    assert rep["markerHalfSizeFt"] == 2.0


def test_619311_has_annotation_style():
    spec = json.loads(
        (ROOT / "Data" / "sheet-specs" / "619-311.json").read_text(encoding="utf-8"))
    assert "annotationStyle" in spec
    chan = next(s for s in spec["symbols"]["items"] if s["id"] == "channelizingDevices")
    assert chan["representation"]["mode"] == "markers"


def test_label_text_policies():
    style = sc.annotation_style({})
    assert sc._label_text("LANE TAPER", 720, style) == "LANE TAPER"
    style["featureLabel"] = "nameAndLength"
    assert "720" in sc._label_text("LANE TAPER", 720, style)


def test_primitive_id():
    assert sc._primitive_id(1, "bufferSpace", "dimension") == "1:bufferSpace:dimension"


def test_placement_registry_roundtrip(tmp_path, monkeypatch):
    path = tmp_path / "placement-registry.jsonl"
    monkeypatch.setattr(preg, "REGISTRY_PATH", path)
    preg.clear_registry()
    preg.append_placement(
        sheet_num="619-311", align_idx=1, kind="cone",
        primitive_id="1:laneTaperRun:cone", bridge_op="PLACE_CHANNELIZING_MARKERS",
        element_ids=["100", "101"],
        spec_ref={"zone": "laneTaper", "run": "laneTaperRun"},
    )
    rows = preg.load_placements(kind="cone", run="laneTaperRun")
    assert len(rows) == 1
    assert rows[0]["elementIds"] == ["100", "101"]
    n = preg.mark_deleted({"1:laneTaperRun:cone"})
    assert n == 1
    assert preg.load_placements(kind="cone") == []


def test_parse_created_ids():
    assert preg.parse_created_ids({"createdElementIds": "1,2,3"}) == ["1", "2", "3"]
    assert preg.parse_created_ids({"elementId": 9}) == ["9"]


def test_sheet_plan_save_load(tmp_path, monkeypatch):
    plan_path = tmp_path / "sheet-plan.json"
    monkeypatch.setattr(ops, "SHEET_PLAN_PATH", plan_path)
    monkeypatch.setattr(preg, "REGISTRY_PATH", tmp_path / "reg.jsonl")
    ops._PLAN_SESSION.reset()
    ops._PLAN_SESSION.order_table_built = True
    ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width=">= 8 ft", area_type="RURAL")
    ops._PLAN_SESSION.required_aligns = {1, 2}
    ops._PLAN_SESSION.aligns_ready = {1, 2}
    ops._PLAN_SESSION.stations_placed_aligns = {1}
    ops._PLAN_SESSION.signs_placed_rows = {(1, "W20-1")}
    ops._PLAN_SESSION.work_area_edges = {
        "upstream": [1.0, 2.0, 0.0], "downstream": [3.0, 4.0, 0.0]}
    saved = ops._save_sheet_plan()
    assert saved is not None and saved.exists()

    ops._PLAN_SESSION.reset()
    loaded = ops._load_sheet_plan(plan_path)
    assert loaded["loaded"] is True
    assert ops._PLAN_SESSION.sheet_plan_active()
    assert ops._PLAN_SESSION.designer_inputs.sheet_num == "619-311"
    assert ops._PLAN_SESSION.stations_placed_aligns == {1}
    assert (1, "W20-1") in ops._PLAN_SESSION.signs_placed_rows
    assert ops._PLAN_SESSION.work_area_edges["upstream"][0] == 1.0

    st = ops.get_plan_status()
    assert st["sheetPlanActive"] is True
    assert st.get("updatedAt")
    ops._PLAN_SESSION.reset()
