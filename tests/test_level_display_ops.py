"""Unit tests for set_level_display / hide_levels_recipe helpers (no live ORD)."""
from __future__ import annotations

import wztc_ops as ops


def test_quote_level_name_spaces():
    assert ops._quote_level_name("Draft_Sight_Achieved") == '"Draft_Sight_Achieved"'
    assert ops._quote_level_name('Draft_Horiz Alignment Point') == '"Draft_Horiz Alignment Point"'
    assert ops._quote_level_name('"Already"') == '"Already"'


def test_level_truthy():
    assert ops._level_truthy("True") is True
    assert ops._level_truthy("False") is False
    assert ops._level_truthy("") is None
    assert ops._level_truthy(None) is None


def test_set_level_display_empty_name():
    r = ops.set_level_display("", False)
    assert r["status"] == "ERROR"
    assert "required" in r["note"].lower()


def test_hide_levels_recipe_requires_args():
    r = ops.hide_levels_recipe()
    assert r["status"] == "ERROR"
    assert "name_contains" in r["note"] or "levels" in r["note"]
