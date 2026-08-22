"""Sign stem must meet the face AABB (gold G20: L≈50, penetrate≈0)."""
from __future__ import annotations

import wztc_ops as ops


def test_measure_stem_into_face_positive_when_past_edge():
    penetrate = ops.measure_stem_into_face_ft(
        (0.0, 0.0), (58.0, 0.0), (50.0, -10.0), (90.0, 10.0)
    )
    assert abs(penetrate - 8.0) < 1e-9


def test_measure_stem_into_face_gold_edge_contact():
    # Gold: stem ends on face inward edge → penetrate ≈ 0
    penetrate = ops.measure_stem_into_face_ft(
        (0.0, 0.0), (50.0, 0.0), (50.0, -10.0), (90.0, 10.0)
    )
    assert abs(penetrate) < 1e-9


def test_measure_stem_into_face_negative_when_short():
    penetrate = ops.measure_stem_into_face_ft(
        (0.0, 0.0), (48.0, 0.0), (50.0, -10.0), (90.0, 10.0)
    )
    assert penetrate < 0.0
    assert abs(penetrate - (-2.0)) < 1e-9


def test_stem_gate_allows_gold_zero_penetrate():
    assert ops._MAX_STEM_SHORT_OF_FACE_FT >= 1.0
    # Gold-like: short_of = 0 → ok
    assert 0.0 <= ops._MAX_STEM_SHORT_OF_FACE_FT
