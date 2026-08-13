"""Unit tests for orthogonal intersection + ramp gore geometry."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from road_junctions import (  # noqa: E402
    orthogonal_intersection_lines,
    ramp_gore_lines,
    strip_placeable_segments,
)
from lane_highway import travel_width_ft  # noqa: E402


def test_travel_widths():
    assert travel_width_ft("one_way", lanes=3) == 36.0
    assert travel_width_ft("two_way", lanes=4, yellow_gap_ft=2.0) == 50.0
    assert travel_width_ft(
        "divided", lanes_per_direction=2, median_width_ft=20.0,
    ) == 68.0
    assert travel_width_ft(
        "twlt", lanes_per_direction=2, twlt_width_ft=12.0,
    ) == 60.0


def test_plus_intersection_arms_and_marks():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=2,
        secondary_lanes=2,
        primary_bearing_deg=0.0,
        junction="plus",
        has_turning_lanes=False,
        turn_arrows=False,
    )
    arms = {s["arm"] for s in segs}
    assert "primary_neg" in arms and "primary_pos" in arms
    assert "secondary_left" in arms and "secondary_right" in arms
    assert sum(1 for s in segs if s["kind"] == "crosswalk") == 8
    assert sum(1 for s in segs if s["kind"] == "stop_bar") == 4
    assert not any(
        (s.get("arm") or "").startswith("center_extension") for s in segs
    )


def test_edges_meet_box_center_stops_at_stop_bar():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=2,
        secondary_lanes=2,
        has_turning_lanes=False,
        turn_arrows=False,
    )
    # Secondary travel width 26 → box half 13; stop at 13+8+4=25
    stop_s = 13.0 + 8.0 + 4.0
    primary_edges = [
        s for s in segs
        if s["arm"].startswith("primary") and s["kind"] == "edge"
    ]
    # Edges should reach the box (|x| ≈ 13)
    assert any(abs(min(s["x1"], s["x2"]) - 13.0) < 0.5 or
               abs(max(s["x1"], s["x2"]) + 13.0) < 0.5 or
               abs(max(s["x1"], s["x2"]) - 13.0) < 0.5 or
               abs(min(s["x1"], s["x2"]) + 13.0) < 0.5
               for s in primary_edges)
    yellow = [
        s for s in segs
        if s["arm"].startswith("primary") and s["kind"] == "yellow"
    ]
    for s in yellow:
        for x in (s["x1"], s["x2"]):
            assert abs(x) >= stop_s - 1e-3, (s, stop_s)


def test_stub_edges_connect_to_primary_box():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=2,
        secondary_lanes=2,
        turn_arrows=False,
        has_turning_lanes=False,
    )
    # Primary half-width = 13. Stub edges should reach y≈±13
    stub_edges = [
        s for s in segs
        if s["arm"].startswith("secondary") and s["kind"] == "edge"
    ]
    ys = []
    for s in stub_edges:
        ys.extend([s["y1"], s["y2"]])
    assert any(abs(y) - 13.0 < 0.5 for y in ys)


def test_turn_arrow_metas_continuous_shared_options():
    """Equal lanes_in/out at a + → SALS/SARS (2-lane) or SALS/SAS/SARS (3)."""
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=220.0,
        secondary_stub_ft=90.0,
        primary_lanes=4,
        secondary_lanes=2,
        turn_arrows=True,
    )
    arrows = [s for s in segs if s.get("kind") == "turn_arrow"]
    assert arrows
    names = {s["cellName"] for s in arrows}
    assert "SLONLY" not in names
    assert "SALS" in names and "SARS" in names
    primary = [
        s["cellName"] for s in arrows
        if s["arm"] == "primary_neg" and s["cellName"] != "SLONLY"
    ]
    assert primary == ["SALS", "SARS"]
    # West approach travel +X → angle -90; south travel +Y → angle 0
    west = [s for s in arrows if s["arm"] == "primary_neg"]
    assert west and abs(west[0]["angleDeg"] + 90.0) < 1e-6
    assert all(s["y"] < 0 for s in west)
    south = [
        s for s in arrows
        if s["arm"] == "secondary_right" and s["cellName"] in ("SALS", "SARS")
    ]
    assert len(south) == 2
    assert south[0]["x"] == south[1]["x"] and south[0]["y"] == south[1]["y"]
    assert abs(south[0]["angleDeg"]) < 1e-6


def test_turn_arrow_metas_six_lane_shared():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=280.0,
        secondary_stub_ft=100.0,
        primary_lanes=6,
        secondary_lanes=2,
        turn_arrows=True,
    )
    primary = [
        s["cellName"] for s in segs
        if s.get("kind") == "turn_arrow"
        and s["arm"] == "primary_neg"
        and s["cellName"] != "SLONLY"
    ]
    assert primary == ["SALS", "SAS", "SARS"]
    assert not any(
        s.get("cellName") == "SLONLY" and s["arm"].startswith("primary")
        for s in segs if s.get("kind") == "turn_arrow"
    )


def test_turn_arrow_metas_dedicated_when_lane_drops():
    """3 toward, 2 through → SAL+SLONLY aligned with through; median on away."""
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=280.0,
        secondary_stub_ft=110.0,
        primary_lanes=6,
        secondary_lanes=2,
        primary_lanes_out=2,
        turn_arrows=True,
    )
    primary = [
        s for s in segs
        if s.get("kind") == "turn_arrow" and s["arm"] == "primary_neg"
    ]
    names = [s["cellName"] for s in primary]
    assert names[0] == "SAL"
    assert "SLONLY" in names
    non_only = [c for c in names if c != "SLONLY"]
    assert non_only == ["SAL", "SAS", "SARS"]
    # SAL / SAS / SARS share the same station (x for E-W)
    arrows_no_only = [s for s in primary if s["cellName"] != "SLONLY"]
    xs = {round(s["x"], 3) for s in arrows_no_only}
    assert len(xs) == 1
    only = next(s for s in primary if s["cellName"] == "SLONLY")
    sal = next(s for s in primary if s["cellName"] == "SAL")
    # ONLY further upstream (smaller x on west approach), same lane y
    assert only["x"] < sal["x"] - 30.0
    assert abs(only["y"] - sal["y"]) < 1e-6


def test_lane_drop_three_into_two_after():
    """3 into the box / 2 after: approach (south on west arm) has 2 dash rows."""
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=300.0,
        secondary_stub_ft=120.0,
        primary_lanes=6,
        secondary_lanes=2,
        primary_lanes_out=2,
        turn_arrows=False,
    )
    dashes = [
        s for s in segs
        if s.get("arm") == "primary_neg" and s.get("kind") == "lane"
        and s.get("style") == "dashed"
    ]
    south_rows = {
        round(0.5 * (s["y1"] + s["y2"]), 1)
        for s in dashes
        if 0.5 * (s["y1"] + s["y2"]) < 0
    }
    north_rows = {
        round(0.5 * (s["y1"] + s["y2"]), 1)
        for s in dashes
        if 0.5 * (s["y1"] + s["y2"]) > 0
    }
    # South = EB into box = 3 lanes → 2 dashed separators; north away = 2 → 1
    assert len(south_rows) == 2
    assert len(north_rows) == 1

def test_primary_does_not_cross_box():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=2,
        secondary_lanes=2,
        has_turning_lanes=False,
        turn_arrows=False,
    )
    primary_edges = [
        s for s in segs
        if s["arm"].startswith("primary") and s["kind"] == "edge"
    ]
    for s in primary_edges:
        xs = (s["x1"], s["x2"])
        assert max(xs) <= -13.0 + 1e-6 or min(xs) >= 13.0 - 1e-6


def test_tee_only_one_stub():
    segs = orthogonal_intersection_lines(
        100.0, 200.0,
        primary_road_type="one_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=2,
        secondary_lanes=2,
        junction="tee",
        tee_side="right",
        has_turning_lanes=False,
        turn_arrows=False,
    )
    arms = {s["arm"] for s in segs}
    assert "secondary_tee" in arms
    assert "secondary_left" not in arms
    assert sum(1 for s in segs if s["kind"] == "stop_bar") == 3


def test_twlt_auto_dotted_center():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="twlt",
        secondary_road_type="two_way",
        primary_length_ft=240.0,
        secondary_stub_ft=100.0,
        primary_lanes_per_direction=2,
        secondary_lanes=2,
        junction="plus",
        turn_arrows=False,
    )
    ext = [s for s in segs if (s.get("arm") or "").startswith("center_extension")]
    assert ext
    assert all(s["kind"] == "yellow" and s["style"] == "dashed" for s in ext)


def test_explicit_turning_lanes_flag():
    segs = orthogonal_intersection_lines(
        0.0, 0.0,
        primary_road_type="two_way",
        secondary_road_type="two_way",
        primary_length_ft=200.0,
        secondary_stub_ft=80.0,
        primary_lanes=4,
        secondary_lanes=2,
        has_turning_lanes=True,
        turn_arrows=False,
    )
    assert any(
        (s.get("arm") or "").startswith("center_extension") for s in segs
    )


def test_stub_too_short_rejects():
    try:
        orthogonal_intersection_lines(
            0.0, 0.0,
            primary_road_type="two_way",
            secondary_road_type="two_way",
            primary_length_ft=200.0,
            secondary_stub_ft=5.0,
            primary_lanes=2,
            secondary_lanes=2,
        )
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_ramp_gore_nose_and_arms():
    segs = ramp_gore_lines(
        0.0, 100.0, 200.0, 100.0,
        mainline_lanes=3,
        ramp_angle_deg=15.0,
        gore_station_ft=80.0,
        ramp_length_ft=120.0,
        ramp_lanes=1,
        side="right",
        gore_mark_ft=40.0,
    )
    placeable = strip_placeable_segments(segs)
    arms = {s["arm"] for s in placeable}
    assert "mainline" in arms and "ramp" in arms and "gore" in arms
    nose = next(s for s in segs if s["kind"] == "gore_nose")
    assert abs(nose["x1"] - 80.0) < 1e-6
    assert abs(nose["y1"] - 64.0) < 1e-6
    assert sum(1 for s in placeable if s["kind"] == "gore") == 2


def test_ramp_gore_rejects_bad_station():
    try:
        ramp_gore_lines(
            0.0, 0.0, 50.0, 0.0,
            mainline_lanes=2,
            ramp_angle_deg=10.0,
            gore_station_ft=100.0,
            ramp_length_ft=50.0,
        )
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_ramp_diverges_toward_side():
    segs = ramp_gore_lines(
        0.0, 0.0, 100.0, 0.0,
        mainline_lanes=2,
        ramp_angle_deg=20.0,
        gore_station_ft=40.0,
        ramp_length_ft=50.0,
        side="right",
        gore_mark_ft=0.0,
    )
    ramp_edges = [
        s for s in segs
        if s.get("arm") == "ramp" and s.get("kind") == "edge" and s.get("style") == "solid"
    ]
    nose_y = 0.0 - 24.0
    end_ys = [s["y2"] for s in ramp_edges]
    assert min(end_ys) < nose_y + 1.0


def test_ramp_gore_curved_mainline():
    verts = [
        [0.0, 100.0],
        [100.0, 100.0],
        [200.0, 140.0],
        [300.0, 140.0],
    ]
    segs = ramp_gore_lines(
        vertices=verts,
        mainline_lanes=2,
        ramp_angle_deg=15.0,
        gore_station_ft=150.0,
        ramp_length_ft=80.0,
        ramp_lanes=1,
        side="right",
        gore_mark_ft=20.0,
    )
    placeable = strip_placeable_segments(segs)
    arms = {s["arm"] for s in placeable}
    assert "mainline" in arms and "ramp" in arms
    main_edges = [
        s for s in placeable
        if s.get("arm") == "mainline" and s.get("kind") == "edge" and s.get("style") == "solid"
    ]
    assert any(s.get("vertices") for s in main_edges) or any(
        abs(float(s["x2"]) - float(s["x1"])) > 1.0
        or abs(float(s["y2"]) - float(s["y1"])) > 1.0
        for s in main_edges
    )
    nose = next(s for s in segs if s["kind"] == "gore_nose")
    # Nose is on ramp-side outer (off = 24 for 2*12) along the path
    assert nose["x1"] != 0.0
    assert 100.0 < nose["x1"] < 300.0
