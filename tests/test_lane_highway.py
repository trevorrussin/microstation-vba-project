"""Unit tests for road-strip geometry (no MicroStation required)."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from lane_highway import (  # noqa: E402
    divided_highway_lines,
    lane_highway_lines,
    twlt_highway_lines,
    two_way_highway_lines,
)


def test_three_lane_horizontal_counts():
    segs = lane_highway_lines(3, 0.0, 100.0, 1000.0, 100.0, side="right")
    solids = [s for s in segs if s["style"] == "solid"]
    dashed = [s for s in segs if s["style"] == "dashed"]
    assert len(solids) == 2
    # 2 dashed rows * (1000/40)=25 segs each
    assert len(dashed) == 50
    assert solids[0]["y1"] == 100.0 and solids[0]["y2"] == 100.0
    assert solids[1]["y1"] == 64.0 and solids[1]["y2"] == 64.0  # 100 - 3*12


def test_one_lane_solids_only():
    segs = lane_highway_lines(1, 10.0, 50.0, 110.0, 50.0)
    assert all(s["style"] == "solid" for s in segs)
    assert len(segs) == 2
    assert segs[1]["y1"] == 38.0  # 50 - 12


def test_four_lane_dashed_row_count():
    segs = lane_highway_lines(4, 0.0, 0.0, 400.0, 0.0)
    rows = {s["row"] for s in segs if s["style"] == "dashed"}
    assert rows == {1, 2, 3}
    assert sum(1 for s in segs if s["style"] == "solid") == 2


def test_dash_gap_lengths():
    segs = lane_highway_lines(2, 0.0, 0.0, 100.0, 0.0, dash_ft=10.0, gap_ft=30.0)
    dashed = [s for s in segs if s["style"] == "dashed"]
    # 100/40 = 2 full periods + leftover? t=0..10, t=40..50, t=80..90 => 3 segs
    assert len(dashed) == 3
    for s in dashed:
        assert abs((s["x2"] - s["x1"]) - 10.0) < 1e-9


def test_left_side_goes_positive_y_for_plus_x():
    segs = lane_highway_lines(1, 0.0, 0.0, 100.0, 0.0, side="left")
    bot = [s for s in segs if s["row"] == 1][0]
    assert bot["y1"] == 12.0


def test_custom_lane_width_spacing():
    segs = lane_highway_lines(2, 0.0, 100.0, 200.0, 100.0, lane_width_ft=14.0)
    solids = [s for s in segs if s["style"] == "solid"]
    assert solids[0]["y1"] == 100.0
    assert solids[1]["y1"] == 72.0  # 100 - 2*14
    dashed = [s for s in segs if s["style"] == "dashed"]
    assert dashed and all(abs(s["y1"] - 86.0) < 1e-9 for s in dashed)  # row at 100-14


def test_rejects_bad_lanes():
    try:
        lane_highway_lines(0, 0.0, 0.0, 10.0, 0.0)
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_one_way_shoulders():
    segs = lane_highway_lines(
        2, 0.0, 100.0, 100.0, 100.0, shoulder_width_ft=8.0, side="right",
    )
    shoulders = [s for s in segs if s["kind"] == "shoulder"]
    assert len(shoulders) == 2
    ys = sorted(s["y1"] for s in shoulders)
    # travel edges at 100 and 76; shoulders at 108 and 68
    assert ys == [68.0, 108.0]


def test_two_way_two_lane_no_dashed():
    segs = two_way_highway_lines(2, 0.0, 100.0, 200.0, 100.0, side="right")
    assert all(s["style"] == "solid" for s in segs)
    kinds = [s["kind"] for s in segs]
    assert kinds == ["edge", "yellow", "yellow", "edge"]
    # offsets: 0, 12, 14, 26  => y = 100, 88, 86, 74
    ys = [s["y1"] for s in segs]
    assert ys == [100.0, 88.0, 86.0, 74.0]


def test_two_way_four_lane_one_dashed_each_side():
    segs = two_way_highway_lines(4, 0.0, 0.0, 100.0, 0.0)
    solids = [s for s in segs if s["style"] == "solid"]
    dashed = [s for s in segs if s["style"] == "dashed"]
    assert [s["kind"] for s in solids] == ["edge", "yellow", "yellow", "edge"]
    # rows: 0 edge, 1 dashed, 2 yellow, 3 yellow, 4 dashed, 5 edge
    assert {s["row"] for s in dashed} == {1, 4}
    # 2 dashed rows * 3 segs on 100ft @ 10/30
    assert len(dashed) == 6
    # yellow gap 2ft after first yellow at offset 24: second at 26
    yellows = [s for s in solids if s["kind"] == "yellow"]
    assert abs(yellows[0]["y1"] - yellows[1]["y1"]) == 2.0


def test_two_way_six_lane_two_dashed_each_side():
    segs = two_way_highway_lines(6, 0.0, 200.0, 40.0, 200.0)
    dashed_rows = sorted({s["row"] for s in segs if s["style"] == "dashed"})
    assert dashed_rows == [1, 2, 5, 6]
    yellows = [s for s in segs if s["kind"] == "yellow"]
    assert len(yellows) == 2
    assert abs(yellows[0]["y1"] - yellows[1]["y1"]) == 2.0


def test_two_way_rejects_odd_lanes():
    try:
        two_way_highway_lines(3, 0.0, 0.0, 10.0, 0.0)
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_two_way_custom_yellow_gap():
    segs = two_way_highway_lines(
        2, 0.0, 50.0, 100.0, 50.0, yellow_gap_ft=3.0, lane_width_ft=12.0,
    )
    yellows = [s for s in segs if s["kind"] == "yellow"]
    assert abs(yellows[0]["y1"] - yellows[1]["y1"]) == 3.0


def test_divided_three_per_dir_median():
    segs = divided_highway_lines(
        3, 0.0, 100.0, 200.0, 100.0, median_width_ft=20.0, side="right",
    )
    # rows: edge, dash, dash, yellow, yellow, dash, dash, edge
    solids = [s for s in segs if s["style"] == "solid"]
    kinds = [s["kind"] for s in solids]
    assert kinds == ["edge", "yellow", "yellow", "edge"]
    yellows = [s for s in solids if s["kind"] == "yellow"]
    assert abs(yellows[0]["y1"] - yellows[1]["y1"]) == 20.0
    # first yellow at off=36 (3*12), second at 56
    assert yellows[0]["y1"] == 64.0  # 100-36
    assert yellows[1]["y1"] == 44.0  # 100-56
    assert solids[-1]["y1"] == 8.0   # 100 - (56+36) = 100-92


def test_divided_requires_median():
    try:
        divided_highway_lines(2, 0.0, 0.0, 10.0, 0.0, median_width_ft=0.0)
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_divided_with_shoulders():
    segs = divided_highway_lines(
        2, 0.0, 50.0, 100.0, 50.0, median_width_ft=10.0, shoulder_width_ft=6.0,
    )
    assert sum(1 for s in segs if s["kind"] == "shoulder") == 2


def test_twlt_two_per_dir():
    segs = twlt_highway_lines(
        2, 0.0, 100.0, 100.0, 100.0, twlt_width_ft=12.0, side="right",
    )
    # edge, dash, yellow-dash, yellow-dash, dash, edge
    yellows = [s for s in segs if s["kind"] == "yellow"]
    assert yellows and all(s["style"] == "dashed" for s in yellows)
    # yellow rows at off 24 and 36
    y_offs = sorted({round(100.0 - s["y1"], 6) for s in yellows})
    assert y_offs == [24.0, 36.0]
    edges = [s for s in segs if s["kind"] == "edge"]
    assert len(edges) == 2
    assert edges[0]["y1"] == 100.0
    assert edges[1]["y1"] == 40.0  # 100 - (36+12)


def test_twlt_one_per_dir_no_white_dash():
    segs = twlt_highway_lines(1, 0.0, 0.0, 80.0, 0.0, twlt_width_ft=14.0)
    white_dash = [s for s in segs if s["style"] == "dashed" and s["kind"] == "lane"]
    assert white_dash == []
    yellow_rows = sorted({s["row"] for s in segs if s["kind"] == "yellow"})
    assert yellow_rows == [1, 2]
