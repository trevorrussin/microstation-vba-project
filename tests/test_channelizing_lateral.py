"""Channelizing offsets: align = left edge of closed lane (channelizing)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))

import alignment_geometry as ag
import sheet_compile
import sheet_spec


def _segs(y: float, x0: float, x1: float):
    verts = [{
        "segIndex": "0", "isArc": "N",
        "sx": str(x0), "sy": str(y), "sz": "0",
        "ex": str(x1), "ey": str(y), "ez": "0",
        "segLen": str(abs(x1 - x0)),
        "cx": "0", "cy": "0", "radius": "0",
        "startAngle": "0", "sweepAngle": "0",
    }]
    return ag.parse_vertices(verts)


def test_longitudinal_on_channelizing_line():
    spec = sheet_spec.load("619-311")
    res = sheet_spec.resolve(spec, 55, 12, ">= 8 ft", "URBAN", None, None)
    segs = _segs(290188.0, 34500.0, 32000.0)
    prims = sheet_compile.compile_channelizing(
        spec, res, 1, segs, lane_width_ft=12, shoulder_width_ft=8,
        outward_sign=1.0)
    long = [p for p in prims if p["run"] == "longitudinalRun"]
    assert long
    assert all(abs(p["y"] - 290188.0) < 0.05 for p in long)


def test_lane_taper_tip_at_outer_travel_edge():
    spec = sheet_spec.load("619-311")
    res = sheet_spec.resolve(spec, 55, 12, ">= 8 ft", "URBAN", None, None)
    segs = _segs(290188.0, 34500.0, 32000.0)
    prims = sheet_compile.compile_channelizing(
        spec, res, 1, segs, lane_width_ft=12, shoulder_width_ft=8,
        outward_sign=1.0)
    lane = [p for p in prims if p["run"] == "laneTaperRun"]
    tip = max(lane, key=lambda p: p["stationFt"])
    toe = min(lane, key=lambda p: p["stationFt"])
    assert tip["y"] < toe["y"]  # south tip (outer) vs channelizing toe
    assert abs(toe["y"] - 290188.0) < 0.05
    assert abs(tip["y"] - 290176.0) < 1.0  # outer travel (± station sample)


def test_shoulder_taper_to_eop():
    spec = sheet_spec.load("619-311")
    res = sheet_spec.resolve(spec, 55, 12, ">= 8 ft", "URBAN", None, None)
    segs = _segs(290188.0, 34500.0, 32000.0)
    prims = sheet_compile.compile_channelizing(
        spec, res, 1, segs, lane_width_ft=12, shoulder_width_ft=8,
        outward_sign=1.0)
    sh = [p for p in prims if p["run"] == "shoulderTaperRun"]
    ys = [p["y"] for p in sh]
    assert min(ys) <= 290168.0 + 0.05  # EOP
    assert max(ys) >= 290176.0 - 0.05  # travel outer
