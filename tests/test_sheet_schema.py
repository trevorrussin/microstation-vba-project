"""Offline Pydantic schema gate for Data/sheet-specs."""
from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import ValidationError

import sheet_schema
import sheet_spec

SPEC_DIR = Path(__file__).resolve().parent.parent / "Data" / "sheet-specs"


def test_all_619_specs_pass_schema():
    paths = sorted(SPEC_DIR.glob("619-*.json"))
    assert paths, "no sheet specs found"
    for path in paths:
        sheet_spec.load_raw_path(path)  # raises SpecError on failure


def test_plan_sheet_requires_corridor():
    with pytest.raises(ValidationError):
        sheet_schema.validate_sheet_dict({
            "schemaVersion": "1.0",
            "sheet": {"number": "x", "title": "t"},
            "tableRoles": {},
            "tables": {"t": {"rows": [{}]}},
        })


def test_reference_library_ok_without_corridor():
    model = sheet_schema.validate_sheet_dict({
        "schemaVersion": "1.0",
        "sheet": {"number": "619-011", "title": "lib", "kind": "referenceLibrary"},
        "tableRoles": {"taper": "011-01"},
        "tables": {"011-01": {"rows": [{"speedMph": 45}]}},
    })
    assert model.sheet.kind == "referenceLibrary"
