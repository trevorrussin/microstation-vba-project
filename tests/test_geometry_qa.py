"""Offline Shapely geometry QA (sheet_geometry_qa)."""
from __future__ import annotations

from sheet_geometry_qa import check_compiled_geometry


def test_valid_hatch_and_symbols_outside_pass():
    hatch = [{
        "kind": "hatch",
        "boundary": [(0, 0), (100, 0), (100, 12), (0, 12)],
        "workAreaLengthFt": 100,
    }]
    symbols = [
        {"kind": "protectiveVehicle", "id": "pv", "x": -50, "y": 6, "stationFt": 80},
        {"kind": "arrowPanel", "id": "ap", "x": -10, "y": 6, "stationFt": 0},
    ]
    assert check_compiled_geometry(symbols, hatch, []) == []


def test_pv_inside_hatch_fails():
    hatch = [{
        "kind": "hatch",
        "boundary": [(0, 0), (100, 0), (100, 12), (0, 12)],
        "workAreaLengthFt": 100,
    }]
    symbols = [
        {"kind": "protectiveVehicle", "id": "pv", "x": 50, "y": 6, "stationFt": 80},
    ]
    fails = check_compiled_geometry(symbols, hatch, [])
    assert any("inside the work-area hatch" in f for f in fails)


def test_invalid_bowtie_hatch_fails():
    bow = [{
        "kind": "hatch",
        "boundary": [(0, 0), (10, 10), (10, 0), (0, 10)],
        "workAreaLengthFt": 10,
    }]
    fails = check_compiled_geometry([], bow, [])
    assert any("invalid" in f or "self-intersect" in f for f in fails)


def test_altgroup_colocated_symbols_ok():
    hatch = [{
        "kind": "hatch",
        "boundary": [(0, 0), (100, 0), (100, 12), (0, 12)],
        "workAreaLengthFt": 100,
    }]
    alts = [
        {"kind": "protectiveVehicle", "id": "pv", "x": -50, "y": 6,
         "stationFt": 100, "altGroup": "alt1"},
        {"kind": "arrowPanel", "id": "ap", "x": -50, "y": 6,
         "stationFt": 100, "altGroup": "alt1"},
    ]
    assert check_compiled_geometry(alts, hatch, []) == []


def test_stacked_ap_pv_different_stations_fails():
    hatch = [{
        "kind": "hatch",
        "boundary": [(0, 0), (100, 0), (100, 12), (0, 12)],
        "workAreaLengthFt": 100,
    }]
    stack = [
        {"kind": "protectiveVehicle", "id": "pv", "x": -50, "y": 0, "stationFt": 80},
        {"kind": "arrowPanel", "id": "ap", "x": -49, "y": 0, "stationFt": 0},
    ]
    fails = check_compiled_geometry(stack, hatch, [])
    assert any("apart" in f for f in fails)
