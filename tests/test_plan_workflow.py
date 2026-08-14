"""Sheet-plan checklist: active only after build_wztc_order_table."""
from __future__ import annotations

import pytest

import plan_workflow
import wztc_ops


@pytest.fixture(autouse=True)
def _clean():
    wztc_ops._PLAN_SESSION.reset()
    yield
    wztc_ops._PLAN_SESSION.reset()


def test_get_plan_status_inactive_outside_sheet_build():
    st = wztc_ops.get_plan_status()
    assert st["sheetPlanActive"] is False
    assert st.get("nextTool") is None


def test_checklist_advances_after_order_table_lock():
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    wztc_ops._PLAN_SESSION.required_aligns = {1, 2}
    wztc_ops._PLAN_SESSION.lock_sign_rows([
        {"align_idx": 1, "sign_num": "W20-01RF"},
        {"align_idx": 2, "sign_num": "G20-02"},
    ])
    st = wztc_ops.get_plan_status()
    assert st["sheetPlanActive"] is True
    assert st["currentStep"] == "corridor_ready"
    assert st["nextTool"] == "propose_corridor_source"
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    wztc_ops._PLAN_SESSION.lock_sign_rows([
        {"align_idx": 1, "sign_num": "W20-01RF"},
    ])
    wztc_ops._PLAN_SESSION.aligns_ready = {1, 2}
    with pytest.raises(ValueError, match="PLAN_GATE"):
        wztc_ops.place_sign(
            "W20-01RF", "Non-Freeway", "One Side",
            0, 0, 0, 0, -1, align_idx=1)


def test_adjust_view_refuses_after_compiler_during_sheet_plan():
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    wztc_ops._PLAN_SESSION.sheet_geometry_placed = True
    with pytest.raises(ValueError, match="run_visual_qa_captures"):
        wztc_ops.adjust_view(center_x=1, center_y=2, width=100, height=50)


def test_adjust_view_gate_skipped_when_no_sheet_plan():
    """General CAD: gate must not fire (COM may still fail in unit tests)."""
    # No order table — sheet_plan_active False. Gate skipped; navigate may error.
    try:
        wztc_ops.adjust_view(center_x=1, center_y=2, width=100, height=50)
    except ValueError as e:
        assert "PLAN_GATE" not in str(e)
        assert "run_visual_qa_captures" not in str(e)
    except Exception:
        # COM / MicroStation unavailable — OK for this unit test
        pass


def test_plan_gate_format_includes_accepted():
    with pytest.raises(ValueError, match="accepted: URBAN"):
        plan_workflow.raise_plan_gate(
            "needs area_type",
            missing=["area_type"],
            accepted=["URBAN", "RURAL", "FREEWAY"],
            next_tool="place_sheet_geometry",
        )


def test_run_sheet_build_noop_outside_sheet_plan():
    out = wztc_ops.run_sheet_build(
        upstream_edge=[0, 0], downstream_edge=[100, 0])
    assert out["sheetPlanActive"] is False


def test_run_sheet_build_requires_edges_when_corridor_missing():
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    wztc_ops._PLAN_SESSION.required_aligns = {1, 2}
    wztc_ops._PLAN_SESSION.lock_sign_rows([
        {"align_idx": 1, "sign_num": "W20-01RF", "side": "One Side"},
    ])
    with pytest.raises(ValueError, match="upstream_edge"):
        wztc_ops.run_sheet_build()


def test_checklist_next_tool_is_propose_corridor_after_order_table():
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    wztc_ops._PLAN_SESSION.required_aligns = {1, 2}
    st = wztc_ops.get_plan_status()
    assert st["nextTool"] == "propose_corridor_source"


def test_empty_locked_signs_not_vacuous_done_for_619311():
    """Stale plan with order_table_built but empty lockedSignRows must not
    skip PLACE_SIGN on sheets that list roadside signs."""
    s = wztc_ops._PLAN_SESSION
    s.order_table_built = True
    s.lock_designer_inputs(
        sheet_num="619-311", speed=55, road_type="Non-Freeway",
        lane_width=12, shoulder_width=">= 8 ft", area_type="URBAN",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    s.required_aligns = {1, 2}
    s.aligns_ready = {1, 2}
    s.stations_placed_aligns = {1, 2}
    s.locked_sign_rows = set()
    s.locked_sign_details = []
    done = plan_workflow.stage_done(s)
    assert done["signs_placed"] is False
    st = wztc_ops.get_plan_status()
    assert st["nextTool"] == "build_wztc_order_table"


def test_complete_real_road_points_at_guide_cleanup():
    s = wztc_ops._PLAN_SESSION
    s.order_table_built = True
    s.lock_designer_inputs(
        sheet_num="619-311", speed=55, road_type="Non-Freeway",
        lane_width=12, shoulder_width=">= 8 ft", area_type="URBAN",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    s.required_aligns = {1, 2}
    s.aligns_ready = {1, 2}
    s.stations_placed_aligns = {1, 2}
    s.lock_sign_rows([
        {"align_idx": 1, "sign_num": "W20-01RF"},
        {"align_idx": 2, "sign_num": "G20-02"},
    ])
    s.signs_placed_rows = set(s.locked_sign_rows)
    s.sign_attrs_applied = True
    s.sheet_geometry_placed = True
    s.geometry_qa_passed = True
    s.visual_qa_passed = True
    s.real_road_edge = True
    st = wztc_ops.get_plan_status()
    assert st["currentStep"] == "complete"
    assert st["nextTool"] == "delete_construction_guides"
    assert "place_two_way_highway" in st["nextStep"]

