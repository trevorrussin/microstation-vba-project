"""Spec-driven designer-input ask list (no invented 60 mph)."""
from __future__ import annotations

import json

import plan_workflow
import sheet_spec
import wztc_ops


def test_619311_ask_count_derives_closure_and_sign_class():
    spec = sheet_spec.load("619-311")
    out = sheet_spec.required_designer_inputs(spec, locked={})
    assert out["askCount"] == 5
    ids = [i["id"] for i in out["toAsk"]]
    assert ids == [
        "preconstructionPostedSpeedMph",
        "laneWidthFt",
        "shoulderWidthBand",
        "areaType",
        "exposureCondition",
    ]
    derived = {d["id"]: d["value"] for d in out["derived"]}
    assert derived["closureType"] == "LANE CLOSURE OR ENCROACHMENT"
    assert derived["signSizeClass"] == "NON-FREEWAY"


def test_619311_speed_options_exclude_60():
    spec = sheet_spec.load("619-311")
    out = sheet_spec.required_designer_inputs(spec)
    speed = next(i for i in out["toAsk"] if i["id"] == "preconstructionPostedSpeedMph")
    blob = json.dumps(speed["askUserChoice"])
    assert "60" not in blob
    labels = [o["label"] for o in speed["askUserChoice"]["options"]]
    assert any("45" in lab for lab in labels)
    assert any(lab == "Other" for lab in labels)
    other = next(o for o in speed["askUserChoice"]["options"] if o["label"] == "Other")
    assert "60" not in other["description"]
    assert "25" in other["description"] or "45" in json.dumps(labels)


def test_validate_rejects_60_on_619311():
    spec = sheet_spec.load("619-311")
    bad = sheet_spec.validate_designer_input_value(
        spec, "preconstructionPostedSpeedMph", 60)
    assert bad["ok"] is False
    good = sheet_spec.validate_designer_input_value(
        spec, "preconstructionPostedSpeedMph", 45)
    assert good["ok"] is True


def test_locked_speed_not_reasked():
    spec = sheet_spec.load("619-311")
    out = sheet_spec.required_designer_inputs(spec, locked={"speed": 45})
    ids = [i["id"] for i in out["toAsk"]]
    assert "preconstructionPostedSpeedMph" not in ids
    assert any(l["id"] == "preconstructionPostedSpeedMph" for l in out["locked"])


def test_all_locked_means_zero_asks():
    spec = sheet_spec.load("619-311")
    out = sheet_spec.required_designer_inputs(spec, locked={
        "speed": 45,
        "lane_width": 12,
        "shoulder_width": ">= 8 ft",
        "area_type": "URBAN",
        "exposure_condition": (
            "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC"
        ),
    })
    assert out["askCount"] == 0


def test_prompt_points_at_spec_lookup():
    import prompts
    text = prompts.WZTC_SYSTEM_PROMPT_ADDENDUM
    assert "get_required_designer_inputs" in text
    assert '{"label": "45"' not in text


def test_wztc_ops_wrapper_and_plan_next_tool(monkeypatch):
    wztc_ops._PLAN_SESSION.reset()
    out = wztc_ops.get_required_designer_inputs("619-311")
    assert out["status"] == "OK"
    assert out["askCount"] == 5
    sess = wztc_ops.PlanSession()
    done = plan_workflow.stage_done(sess)
    act = plan_workflow.next_action(sess, done)
    assert act["nextTool"] == "get_required_designer_inputs"
    wztc_ops._PLAN_SESSION.reset()


def test_tools_for_turn_omits_catalogs_when_plan_active():
    import chat_driver

    wztc_ops._PLAN_SESSION.reset()
    chat_driver._SESSION.mode = "wztc"
    names_idle = {chat_driver._tool_registered_name(t) for t in chat_driver.tools_for_turn()}
    assert "place_two_way_highway" in names_idle
    assert "place_lane_highway" in names_idle
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    names_active = {chat_driver._tool_registered_name(t) for t in chat_driver.tools_for_turn()}
    assert "place_two_way_highway" in names_active
    assert "get_required_designer_inputs" in names_active
    assert "place_lane_highway" not in names_active
    assert "list_registry_commands" not in names_active
    assert "list_cells" not in names_active
    wztc_ops._PLAN_SESSION.reset()
    chat_driver._SESSION.mode = "general"
