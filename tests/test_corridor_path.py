"""Corridor path helpers and lock/snap without MicroStation."""
from __future__ import annotations

import corridor_path as cp
import sheet_spec
import wztc_ops


def test_nearest_and_snap_on_straight_path():
    pts = [[0, 0, 0], [1000, 0, 0]]
    n = cp.nearest_station(pts, 400, 50)
    assert abs(n["stationFt"] - 400) < 1e-6
    assert abs(n["y"]) < 1e-6
    mid = cp.point_at_station(pts, 250)
    assert abs(mid[0] - 250) < 1e-6


def test_offset_goes_right_of_travel():
    pts = [[0, 0, 0], [100, 0, 0]]
    off = cp.offset_polyline(pts, 10)
    assert off[0][1] < 0  # +X travel, right is -Y


def test_length_check_urban_55():
    spec = sheet_spec.load("619-311")
    resolved = sheet_spec.resolve(spec, 55, 12, "8 ft", "URBAN")
    ap = cp.sheet_approach_ft(spec, resolved)
    assert ap["upstreamFt"] > 1000
    assert ap["downstreamFt"] > 0
    bad = cp.length_check(100.0, ap, 200.0)
    assert bad["ok"] is False
    assert "Extend" in (bad["note"] or "")
    good = cp.length_check(ap["bothSidesFt"] + 500, ap, 200.0)
    assert good["ok"] is True


def test_propose_ladder_recommends_last_placed():
    wztc_ops._LAST_PLACED_ROAD = None
    wztc_ops._PLAN_SESSION.reset()
    out = wztc_ops.propose_corridor_source()
    labels = [o["label"] for o in out["askUserChoice"]["options"]]
    assert not any("Recommended" in lab for lab in labels)
    wztc_ops._remember_placed_road(
        road_type="two_way_undivided", lanes=4, lane_width_ft=12,
        shoulder_width_ft=8, yellow_gap_ft=2, side="right",
        verts=[[0, 0], [6000, 0]], x1=0, y1=0, x2=6000, y2=0, length=6000,
    )
    out2 = wztc_ops.propose_corridor_source()
    assert out2["lastPlacedAvailable"] is True
    assert "Recommended" in out2["askUserChoice"]["options"][0]["label"]
    locked = wztc_ops.lock_corridor_path("last_placed")
    assert locked["status"] == "OK"
    assert locked["closedSideDerived"] is None  # no designer inputs yet
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=55, road_type="Non-Freeway",
        lane_width=12, shoulder_width="8", area_type="URBAN",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)
    locked2 = wztc_ops.lock_corridor_path("last_placed")
    assert locked2["closed_side"] == "right"
    snap = wztc_ops.snap_work_area_to_path(
        "ends", p1=[2500, 40], p2=[2740, -10])
    assert snap["status"] == "OK"
    assert abs(snap["workLenFt"] - 240) < 1.0
    assert abs(snap["upstream_edge"][1]) < 1e-6
    wztc_ops._LAST_PLACED_ROAD = None
    wztc_ops._PLAN_SESSION.reset()


def test_prompt_names_corridor_tools():
    import prompts
    text = prompts.WZTC_SYSTEM_PROMPT_ADDENDUM
    assert "propose_corridor_source" in text
    assert "snap_work_area_to_path" in text
