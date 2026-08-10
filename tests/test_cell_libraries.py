"""Tests for cell-library listing (no MicroStation required if folder exists)."""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops  # noqa: E402


def test_list_cell_libraries_ny_plan_folder():
    folder = wztc_ops.DEFAULT_CELL_LIB_DIR
    if not os.path.isdir(folder):
        return  # skip when pwworking not mounted
    r = wztc_ops.list_cell_libraries()
    assert r["status"] == "OK"
    assert r["count"] >= 1
    names = {row["name"].lower() for row in r["libraries"]}
    assert "ny_plan_wztc.cel" in names
    assert "ny_plan_striping.cel" in names


def test_list_cell_libraries_filter():
    folder = wztc_ops.DEFAULT_CELL_LIB_DIR
    if not os.path.isdir(folder):
        return
    r = wztc_ops.list_cell_libraries(name_contains="util")
    assert r["status"] == "OK"
    assert r["count"] >= 1
    assert all("util" in row["name"].lower() for row in r["libraries"])


def test_list_cell_libraries_missing_dir():
    r = wztc_ops.list_cell_libraries(lib_dir=r"c:\no\such\cell\dir\zzz")
    assert r["status"] == "ERROR"
    assert r["libraries"] == []


def test_find_cell_requires_query():
    r = wztc_ops.find_cell("")
    assert r["status"] == "ERROR"
    assert r["matches"] == []
