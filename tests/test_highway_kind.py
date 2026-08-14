"""Sheet vs placed-highway caution (all 619 specs, not 311-only)."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import sheet_spec as ss  # noqa: E402


def test_619311_is_two_way_undivided():
    spec = ss.load("619-311")
    assert spec is not None
    kinds = ss.highway_kinds(spec)
    assert kinds == ["two_way_undivided"]
    ok = ss.highway_kind_match(spec, "two_way_undivided")
    assert ok["mismatch"] is False
    bad = ss.highway_kind_match(spec, "divided")
    assert bad["mismatch"] is True
    assert "askUserChoice" in bad
    freeway = ss.highway_kind_match(spec, "freeway")
    assert freeway["mismatch"] is True
    twlt = ss.highway_kind_match(spec, "twlt")
    assert twlt["mismatch"] is True


def test_619302_accepts_divided_or_freeway():
    spec = ss.load("619-302")
    if spec is None:
        return
    kinds = ss.highway_kinds(spec)
    assert "divided" in kinds or "freeway" in kinds
    assert ss.highway_kind_match(spec, "divided")["mismatch"] is False
    assert ss.highway_kind_match(spec, "two_way_undivided")["mismatch"] is True


def test_619312_twlt_not_plain_two_way():
    spec = ss.load("619-312")
    if spec is None:
        return
    kinds = ss.highway_kinds(spec)
    assert "twlt" in kinds
    assert ss.highway_kind_match(spec, "twlt")["mismatch"] is False
    assert ss.highway_kind_match(spec, "two_way_undivided")["mismatch"] is True


def test_no_placed_road_is_not_a_mismatch():
    spec = ss.load("619-311")
    assert ss.highway_kind_match(spec, "")["mismatch"] is False
