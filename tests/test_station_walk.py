"""Offline station_walk / compare_station_tables checks."""
from __future__ import annotations

import sheet_spec
from sheet_rules import compare_station_tables


def test_619311_station_walk_monotonic_upstream():
    spec = sheet_spec.load("619-311")
    assert spec is not None
    resolved = sheet_spec.resolve(spec, 45, 12, "8 ft", "RURAL")
    walk = [w for w in sheet_spec.station_walk(spec, resolved) if w["alignIdx"] == 1]
    main = [w for w in walk if w.get("rowNum") is not None]
    assert len(main) >= 4
    stas = [w["stationFt"] for w in main]
    assert stas == sorted(stas)
    assert stas[-1] > stas[0]


def test_compare_station_tables_sorted_multiset():
    walk = [
        {"rowNum": 1, "stationFt": 80.0, "item": "A"},
        {"rowNum": 2, "stationFt": 440.0, "item": "B"},
        {"rowNum": 3, "stationFt": 1000.0, "item": "C"},
    ]
    vba = [
        {"cumulativeStationFt": 1000.0, "label": "C"},
        {"cumulativeStationFt": 80.0, "label": "A"},
        {"cumulativeStationFt": 440.0, "label": "B"},
    ]
    assert compare_station_tables(vba, walk, tol_ft=0.5) == []


def test_compare_station_tables_detects_drift():
    walk = [
        {"rowNum": 1, "stationFt": 80.0, "item": "A"},
        {"rowNum": 2, "stationFt": 440.0, "item": "B"},
    ]
    vba = [
        {"cumulativeStationFt": 80.0, "label": "A"},
        {"cumulativeStationFt": 450.0, "label": "B"},
    ]
    fails = compare_station_tables(vba, walk, tol_ft=0.5)
    assert fails
