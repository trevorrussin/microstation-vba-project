"""resolve_sheet_lateral: closed-side → outward_sign / half_len."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))

import wztc_ops  # noqa: E402


def setup_function():
    wztc_ops._PLAN_SESSION.reset()


def test_eb_right_real_road_half_len_and_outward():
    # Travel +X through bay; Align1 points west; right → outward_sign +1 → −Y
    wztc_ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=55, road_type="Non-Freeway",
        lane_width=12, shoulder_width=">= 8 ft", area_type="URBAN",
    )
    wztc_ops._PLAN_SESSION.order_table_built = True
    r = wztc_ops.resolve_sheet_lateral(
        [34500.0, 290187.0, 0.0], [34600.0, 290187.0, 0.0],
        closed_side="right", real_road_edge=True,
    )
    assert r["status"] == "OK"
    assert r["outward_sign"] == 1.0
    assert r["half_len"] == 20.0  # 12 + 8
    assert r["outwardUnit"][1] < 0  # south
    assert r["closed_outward"][1] < 0
    assert wztc_ops._PLAN_SESSION.lateral_outward_sign == 1.0
    assert wztc_ops._PLAN_SESSION.lateral_half_len == 20.0
    assert wztc_ops._PLAN_SESSION.closed_outward_y < 0
    assert wztc_ops._PLAN_SESSION.opposite_half_len is None


def test_eb_left_outward_negative():
    r = wztc_ops.resolve_sheet_lateral(
        [0.0, 0.0, 0.0], [100.0, 0.0, 0.0],
        closed_side="left", lane_width_ft=12, shoulder_width_ft=0,
        real_road_edge=True,
    )
    assert r["outward_sign"] == -1.0
    assert r["half_len"] == 12.0
    assert r["outwardUnit"][1] > 0  # north for EB left


def test_abstract_ticks_when_not_real_road():
    r = wztc_ops.resolve_sheet_lateral(
        [0.0, 0.0, 0.0], [100.0, 0.0, 0.0],
        closed_side="right", lane_width_ft=12, shoulder_width_ft=8,
        real_road_edge=False,
    )
    assert r["half_len"] == 40.0
    assert r["real_road_edge"] is False


def test_apply_locked_lateral_prefers_session():
    wztc_ops._PLAN_SESSION.lateral_outward_sign = 1.0
    wztc_ops._PLAN_SESSION.lateral_half_len = 20.0
    o, h, meta = wztc_ops._apply_locked_lateral(-1.0, 40.0, True)
    assert o == 1.0 and h == 20.0 and meta["usedLockedLateral"]
    o2, h2, meta2 = wztc_ops._apply_locked_lateral(-1.0, 40.0, False)
    assert o2 == -1.0 and h2 == 40.0 and not meta2["usedLockedLateral"]
