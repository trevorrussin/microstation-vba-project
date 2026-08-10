"""Control-loop gates: locked designer fill + anti-fish refuses."""
from __future__ import annotations

import pytest

import wztc_ops


@pytest.fixture(autouse=True)
def _clean_plan_session():
    wztc_ops._PLAN_SESSION.reset()
    yield
    wztc_ops._PLAN_SESSION.reset()


def test_merge_fills_blank_area_type():
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    merged = wztc_ops._merge_locked_designer_inputs(
        "619-311", 45, 12, "8", area_type="")
    assert merged["area_type"] == "RURAL"
    assert "area_type" in merged["filledFromLock"]


def test_merge_refuses_conflicting_area_type():
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    with pytest.raises(ValueError, match="conflicts with locked"):
        wztc_ops._merge_locked_designer_inputs(
            "619-311", 45, 12, "8", area_type="URBAN")


def test_find_reference_linework_refuses_default_mid_plan():
    wztc_ops._PLAN_SESSION.order_table_built = True
    with pytest.raises(ValueError, match="assemble_corridor"):
        wztc_ops.find_reference_linework("Default")


def test_find_elements_near_refuses_wide_radius_mid_plan():
    wztc_ops._PLAN_SESSION.order_table_built = True
    with pytest.raises(ValueError, match="too wide"):
        wztc_ops.find_elements_near(0.0, 0.0, 5000.0)


def test_define_alignment_refuses_mid_plan_when_spec_locked():
    wztc_ops._PLAN_SESSION.order_table_built = True
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=45, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="RURAL",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    with pytest.raises(ValueError, match="assemble_corridor"):
        wztc_ops.define_alignment_segment(1, [[0, 0, 0], [100, 0, 0]])
