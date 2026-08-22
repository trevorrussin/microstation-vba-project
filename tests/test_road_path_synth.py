"""Path synthesis: 'build me a curved highway' must produce real geometry."""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import corridor_path as cp  # noqa: E402
import road_path_synth as rps  # noqa: E402


def _total_deflection_deg(verts) -> float:
    a0 = math.atan2(verts[1][1] - verts[0][1], verts[1][0] - verts[0][0])
    a1 = math.atan2(verts[-1][1] - verts[-2][1], verts[-1][0] - verts[-2][0])
    return (math.degrees(a1 - a0) + 180.0) % 360.0 - 180.0


@pytest.mark.parametrize("kw", [
    dict(length_ft=1000, bends=0),
    dict(length_ft=2000, kind="c_curve"),
    dict(length_ft=2000, bends=2),
    dict(length_ft=3000, kind="l_bend"),
    dict(length_ft=5000, bends=4, radius_ft=800),
])
def test_actual_length_matches_request(kw):
    """A 2000 ft request must produce 2000 ft of road, not 2000 ft of chord."""
    r = rps.synthesize_path(start_x=1000.0, start_y=2000.0, **kw)
    assert abs(r["actualLengthFt"] - r["requestedLengthFt"]) < 0.01 * r["requestedLengthFt"]
    assert abs(cp.polyline_length(r["vertices"]) - r["actualLengthFt"]) < 1.0


def test_starts_at_requested_point_and_bearing():
    r = rps.synthesize_path(length_ft=1000, bends=0, start_x=5000.0, start_y=6000.0,
                            bearing_deg=90.0)
    v = r["vertices"]
    assert abs(v[0][0] - 5000.0) < 1e-6 and abs(v[0][1] - 6000.0) < 1e-6
    # Due north: X holds, Y climbs by the full length.
    assert abs(v[-1][0] - 5000.0) < 0.5
    assert abs(v[-1][1] - 7000.0) < 1.0


def test_shape_deflections():
    """Each kind must actually turn the amount its name implies."""
    straight = rps.synthesize_path(length_ft=1000, bends=0)
    assert abs(_total_deflection_deg(straight["vertices"])) < 0.5

    c = rps.synthesize_path(length_ft=2000, kind="c_curve")
    assert 30.0 < abs(_total_deflection_deg(c["vertices"])) < 60.0

    # Reverse-S: equal and opposite bends, so it exits on the entry heading.
    s = rps.synthesize_path(length_ft=2000, bends=2)
    assert abs(_total_deflection_deg(s["vertices"])) < 1.0

    # L-bend is a corner, not a drift.
    l = rps.synthesize_path(length_ft=3000, kind="l_bend")
    assert abs(abs(_total_deflection_deg(l["vertices"])) - 90.0) < 1.0
    assert l["radiusFt"] < 0.25 * 3000, "L-bend radius must read as a corner"


def test_s_curve_actually_reverses():
    """Net-zero deflection is not enough — it must bend both ways."""
    s = rps.synthesize_path(length_ft=2000, bends=2)
    v = s["vertices"]
    cross = []
    for i in range(1, len(v) - 1):
        ax, ay = v[i][0] - v[i - 1][0], v[i][1] - v[i - 1][1]
        bx, by = v[i + 1][0] - v[i][0], v[i + 1][1] - v[i][1]
        cross.append(ax * by - ay * bx)
    assert max(cross) > 1e-6 and min(cross) < -1e-6


def test_bend_count_infers_shape():
    """The engineer says 'two bends'; they should not also have to say 'S'."""
    assert rps.infer_kind(0) == "straight"
    assert rps.infer_kind(1) == "c_curve"
    assert rps.infer_kind(2) == "s_curve"
    assert rps.infer_kind(4) == "n_bend"
    # An explicit kind wins over the count.
    assert rps.infer_kind(2, "l_bend") == "l_bend"
    # Engineer wording maps to a shape.
    assert rps.infer_kind(None, "curved") == "c_curve"
    assert rps.infer_kind(None, "reverse-S") == "s_curve"


def test_assumptions_are_reported_not_silent():
    """Anything defaulted must come back so the agent can state it."""
    vague = rps.synthesize_path(length_ft=2000, kind="curved")
    assert "radiusFt" in vague["assumedDefaults"]
    # Engineer-supplied radius is not reported as an assumption.
    told = rps.synthesize_path(length_ft=2000, bends=2, radius_ft=900)
    assert "radiusFt" not in told["assumedDefaults"]
    assert abs(told["radiusFt"] - 900.0) < 1e-6


def test_vertex_budget_capped():
    """VBA rejected a ~5k-vertex PLACE_POLYLINE (live 2026-08-13)."""
    r = rps.synthesize_path(length_ft=50000, bends=3, step_ft=1.0)
    assert r["vertexCount"] <= cp.MAX_PATH_VERTS
    assert abs(r["actualLengthFt"] - 50000.0) < 0.02 * 50000.0


def test_tiny_radius_clamped():
    r = rps.synthesize_path(length_ft=2000, bends=2, radius_ft=1.0)
    assert r["radiusFt"] >= rps._MIN_RADIUS_FT
    assert "radiusClampedFt" in r["assumedDefaults"]


def test_rejects_nonpositive_length():
    with pytest.raises(ValueError):
        rps.synthesize_path(length_ft=0)


def test_description_is_readable():
    r = rps.synthesize_path(length_ft=2000, bends=2, start_x=100000, start_y=300000)
    d = r["description"]
    assert "reverse-S" in d and "2000 ft" in d and "east" in d and "two" in d
