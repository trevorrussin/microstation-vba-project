"""Road striping gap-fill: ask the gaps, never re-ask, never default silently."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import road_inputs as ri  # noqa: E402


def _ids(rows) -> list[str]:
    return [r["id"] for r in rows]


def test_engineer_words_map_to_tools():
    """They say 'highway' or 'intersection', not a Python function name."""
    assert ri.resolve_tool("highway") == "place_two_way_highway"
    assert ri.resolve_tool("divided") == "place_divided_highway"
    assert ri.resolve_tool("intersection") == "place_orthogonal_intersection"
    assert ri.resolve_tool("ramp") == "place_ramp_gore"
    assert ri.resolve_tool("twlt") == "place_twlt_highway"
    # Real tool names pass through.
    assert ri.resolve_tool("place_lane_highway") == "place_lane_highway"
    assert ri.resolve_tool("nonsense") == ""


def test_unknown_tool_is_an_error_not_a_guess():
    out = ri.get_required_road_inputs("teleporter")
    assert out["status"] == "ERROR" and out["found"] is False


def test_vague_request_asks_only_lanes_and_path():
    """'build me a curved highway' — two questions, not six."""
    out = ri.get_required_road_inputs("highway", {})
    assert out["ready"] is False
    assert set(_ids(out["missing"])) == {"lanes", "path"}
    # Everything else is defaulted, and reported so it can be stated.
    assert {d["id"] for d in out["assumedDefaults"]} >= {
        "lane_width_ft", "shoulder_width_ft", "side"}


def test_specific_request_asks_nothing():
    """'S-curve, 2000 ft, two bends, four lanes' — do not re-ask any of it."""
    known = {"lanes": 4, "path": [[0.0, 0.0], [100.0, 0.0]]}
    out = ri.get_required_road_inputs("highway", known)
    assert out["ready"] is True
    assert out["missing"] == []
    assert set(out["answered"]) >= {"lanes", "path"}


def test_path_satisfied_by_any_of_its_forms():
    for known in ({"vertices": [[0, 0], [1, 1]]},
                  {"path_vertices": [[0, 0], [1, 1]]},
                  {"x1": 0, "y1": 0, "x2": 10, "y2": 0}):
        out = ri.get_required_road_inputs("highway", dict(known, lanes=2))
        assert out["ready"] is True, known


def test_divided_requires_median_width():
    """A divided road is not defined without it — must not default to 0."""
    out = ri.get_required_road_inputs(
        "divided", {"lanes_per_direction": 2, "path": [[0, 0], [1, 1]]})
    assert _ids(out["missing"]) == ["median_width_ft"]


def test_tee_side_only_asked_for_a_tee():
    plus = ri.get_required_road_inputs("intersection", {"junction": "plus"})
    tee = ri.get_required_road_inputs("intersection", {"junction": "tee"})
    assert "tee_side" not in _ids(plus["missing"])
    assert "tee_side" in _ids(tee["missing"])


def test_median_asked_only_for_a_divided_arm():
    plain = ri.get_required_road_inputs(
        "intersection", {"junction": "plus", "primary_road_type": "two_way"})
    div = ri.get_required_road_inputs(
        "intersection", {"junction": "plus", "primary_road_type": "divided"})
    assert "primary_median_width_ft" not in _ids(plain["missing"])
    assert "primary_median_width_ft" in _ids(div["missing"])


def test_derived_inputs_are_never_asked():
    out = ri.get_required_road_inputs("intersection", {"junction": "plus"})
    assert "has_turning_lanes" not in _ids(out["missing"])
    assert "has_turning_lanes" in _ids(out["derived"])


def test_junction_point_needs_both_coords():
    out = ri.get_required_road_inputs("intersection", {"junction_x": 100.0})
    assert "junction_point" in _ids(out["missing"])
    out2 = ri.get_required_road_inputs(
        "intersection", {"junction_x": 100.0, "junction_y": 200.0})
    assert "junction_point" not in _ids(out2["missing"])


def test_missing_rows_carry_ask_payloads():
    """The agent must not have to invent option lists."""
    out = ri.get_required_road_inputs("highway", {})
    lanes = next(m for m in out["missing"] if m["id"] == "lanes")
    assert lanes["allowed"] == [2, 4, 6, 8]
    path = next(m for m in out["missing"] if m["id"] == "path")
    values = [o["value"] for o in path["options"]]
    # Proposing must be an option, and clicking points must not be the only one.
    assert "synthesize" in values and "element" in values
    assert values[0] == "synthesize"


def test_side_default_carries_the_warning():
    out = ri.get_required_road_inputs("highway", {})
    side = next(d for d in out["assumedDefaults"] if d["id"] == "side")
    assert "outer edge" in side["note"].lower()


def test_announce_defaults_is_readable():
    out = ri.get_required_road_inputs("highway", {})
    txt = ri.announce_defaults(out["assumedDefaults"])
    assert txt.startswith("Assumed:") and "12.0" in txt
    assert ri.announce_defaults([]) == ""


def test_every_tool_spec_is_well_formed():
    for tool, spec in ri.ROAD_TOOL_INPUTS.items():
        assert spec.get("label")
        seen = set()
        for inp in spec["inputs"]:
            assert inp["id"] not in seen, f"{tool} duplicates {inp['id']}"
            seen.add(inp["id"])
            assert inp["kind"] in (ri.REQUIRED, ri.DEFAULTABLE, ri.DERIVED)
            if inp["kind"] == ri.DEFAULTABLE:
                assert "default" in inp, f"{tool}.{inp['id']} needs a default"
