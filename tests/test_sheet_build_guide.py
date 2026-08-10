"""Sheet build-guide companion (.build.md) loading."""
from __future__ import annotations

import sheet_resolve
import sheet_spec


def test_619311_build_guide_loads():
    guide = sheet_spec.load_build_guide("619-311")
    assert guide is not None
    assert guide["sheetNum"] == "619-311"
    assert "619-311.build.md" in guide["path"].replace("\\", "/")
    assert "Preferred call path" in guide["text"]
    assert "Vehicle Space" in guide["text"]
    assert "annotationStyle" in guide["text"]
    assert guide["charCount"] > 500


def test_build_guide_path_respects_json_pointer():
    spec = sheet_spec.load("619-311")
    assert spec is not None
    assert (spec.get("sheet") or {}).get("buildGuide") == "619-311.build.md"
    path = sheet_resolve.build_guide_path("619-311", spec)
    assert path is not None
    assert path.name == "619-311.build.md"


def test_missing_sheet_has_no_guide():
    assert sheet_spec.load_build_guide("619-999-does-not-exist") is None


def test_build_guide_path_rejects_traversal():
    fake = {
        "sheet": {
            "number": "619-311",
            "title": "x",
            "buildGuide": "../secrets.md",
        }
    }
    # Basename-only: looks for Data/sheet-specs/secrets.md — absent → None
    assert sheet_resolve.build_guide_path("619-311", fake) is None
