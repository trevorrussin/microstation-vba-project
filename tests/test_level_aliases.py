"""list_levels: HDM category letters + feature aliases."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))
import wztc_ops as ops


def test_drainage_is_category_d():
    letter, alias = ops._level_search_category("drainage")
    assert letter == "D"
    assert alias == "drainage"


def test_traffic_is_category_t():
    letter, alias = ops._level_search_category("traffic")
    assert letter == "T"


def test_utilities_is_category_u():
    letter, _ = ops._level_search_category("utilities")
    assert letter == "U"


def test_bridge_is_category_b():
    letter, _ = ops._level_search_category("bridge")
    assert letter == "B"


def test_feature_code_category_letter():
    assert ops._level_category_letter("DCB_P") == "D"
    assert ops._level_category_letter("TWZCD_P") == "T"
    assert ops._level_category_letter("O_Details_Notes_P") == "O"
    assert ops._level_category_letter("Draft_Alignment Text") is None
    assert ops._level_category_letter("Default") is None


def test_catch_basin_still_feature_alias():
    needles, hit = ops._level_search_needles("catch basin")
    assert hit == "catch basin"
    assert "DCB" in needles


def test_raw_prefix_still_works():
    needles, hit = ops._level_search_needles("DCB")
    assert needles[0] == "DCB"
    assert hit is None


def test_categories_file_loads():
    cats = ops._load_level_categories()
    assert cats["drainage"] == "D"
    assert cats["signing"] == "S"
    assert cats["pavement"] == "P"


def test_of_does_not_match_right_of_way():
    letter, _ = ops._level_search_category("of")
    assert letter is None
