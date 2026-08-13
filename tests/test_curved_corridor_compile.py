"""Curved corridor: hatch + dim tips follow path; geometry helpers."""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import alignment_geometry as ag  # noqa: E402
import sheet_compile as sc  # noqa: E402


def _l_bend_path():
    # 200 ft east then 200 ft north — work bay spans the corner.
    return [[0.0, 0.0], [200.0, 0.0], [200.0, 200.0]]


def test_segments_from_polyline_and_nearest():
    segs = ag.segments_from_polyline(_l_bend_path())
    assert abs(ag.total_length(segs) - 400.0) < 1e-6
    sta, dist = ag.nearest_station(segs, 200.0, 0.0)
    assert abs(sta - 200.0) < 1e-6
    assert dist < 1e-6
    x, y, tx, ty = ag.point_at_extended(segs, -50.0)
    assert abs(x - (-50.0)) < 1e-6 and abs(y) < 1e-6
    assert abs(tx - 1.0) < 1e-6


def test_sample_path_vertices_extends_past_end():
    segs = ag.segments_from_polyline([[0.0, 0.0], [100.0, 0.0]])
    verts = ag.sample_path_vertices(segs, 100.0, 150.0, step_ft=25.0)
    assert abs(verts[-1][0] - 150.0) < 1e-6
    assert abs(verts[0][0] - 100.0) < 1e-6


def test_compile_hatch_curved_boundary_not_parallelogram():
    path = _l_bend_path()
    segs = ag.segments_from_polyline(path)
    # Align1: from corner going west (away upstream along first leg).
    a1 = ag.segments_from_polyline([[200.0, 0.0], [0.0, 0.0]])
    # Align2: from north end going north (away downstream).
    a2 = ag.segments_from_polyline([[200.0, 200.0], [200.0, 400.0]])
    bay = ag.sample_path_vertices(segs, 200.0, 400.0, step_ft=25.0)
    # Minimal hatched workArea stub
    spec = {
        "corridor": {"zones": [{"id": "workArea", "hatched": True}]},
        "symbols": {"items": []},
    }
    prims = sc.compile_hatch(
        spec, {}, a1, a2, lane_width_ft=12.0, outward_sign=1.0,
        work_bay_vertices=bay)
    hatch = next(p for p in prims if p["kind"] == "hatch")
    assert hatch["curvedWorkBay"] is True
    assert hatch["workAreaLengthFt"] == pytest.approx(200.0, abs=0.1)
    # Curved boundary has more than 4 corners (densified).
    assert len(hatch["boundary"]) > 4
    # Straight chord hatch would be a 4-pt parallelogram.
    straight = sc.compile_hatch(
        spec, {}, a1, a2, lane_width_ft=12.0, outward_sign=1.0)
    assert len(straight[0]["boundary"]) == 4
    assert straight[0].get("curvedWorkBay") is False


def test_compile_plan_dim_tips_use_local_normals():
    # Align path: east then north. Tangents (and outward) differ across bend.
    segs = ag.segments_from_polyline(
        [[0.0, 0.0], [100.0, 0.0], [100.0, 100.0]])
    x0, y0, t0x, t0y = ag.station_to_xy(segs, 50.0)
    x1, y1, t1x, t1y = ag.station_to_xy(segs, 150.0)
    o0 = sc._outward_unit(t0x, t0y, -1.0)
    o1 = sc._outward_unit(t1x, t1y, -1.0)
    assert abs(t0x - t1x) + abs(t0y - t1y) > 0.5
    assert abs(o0[0] - o1[0]) + abs(o0[1] - o1[1]) > 0.5
    # New compile_plan rule: tip1 uses local outward at prev station,
    # tip2 at far — not a single reused far normal.
    tip1 = (x0 + o0[0] * sc.PERP_HALF_LEN_FT, y0 + o0[1] * sc.PERP_HALF_LEN_FT)
    tip2 = (x1 + o1[0] * sc.PERP_HALF_LEN_FT, y1 + o1[1] * sc.PERP_HALF_LEN_FT)
    assert math.hypot(tip1[0] - tip2[0], tip1[1] - tip2[1]) > 50.0
    # Old bug: both tips with o1 would put tip1 off the first-leg normal.
    bad_tip1 = (x0 + o1[0] * sc.PERP_HALF_LEN_FT, y0 + o1[1] * sc.PERP_HALF_LEN_FT)
    assert math.hypot(tip1[0] - bad_tip1[0], tip1[1] - bad_tip1[1]) > 1.0


def test_compile_plan_uses_tip_half_len_and_linear_size():
    """Real-road half_len must move dim tips; straight span stays Linear Size."""
    import sheet_spec as ss

    segs = ag.segments_from_polyline([[0.0, 0.0], [3000.0, 0.0]])
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    elems = (real.get("sheet") or {}).get("elements", "")
    prims = sc.compile_plan(
        real, resolved, 1, segs, outward_sign=1.0,
        tip_half_len_ft=20.0, sheet_elements=elems)
    dims = [p for p in prims if p["kind"] == "dimension"]
    assert dims
    # Single straight segment → Linear Size (not path-hugging).
    assert all(not p.get("curved") for p in dims)
    assert all(not p.get("path") for p in dims)
    t1 = dims[0]["tip1"]
    assert abs(t1[1] - 20.0) < 0.05
    # Table lengths for Urban 55 must appear on dims (not chord measures).
    texts = " ".join(str(p.get("text") or "") for p in dims)
    assert "495" in texts  # buffer Table 311-02
    assert "120" in texts  # roll ahead min Table 311-04


def test_compile_plan_curved_dims_hug_roadside():
    """Bent tip path → path-hugging dim with sheet table length text."""
    import sheet_spec as ss

    # Bend early so roll-ahead / buffer tip paths bow (Align1 sta0 = WA edge).
    segs = ag.segments_from_polyline(
        [[0.0, 0.0], [50.0, 0.0], [50.0, 2500.0]])
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    elems = (real.get("sheet") or {}).get("elements", "")
    prims = sc.compile_plan(
        real, resolved, 1, segs, outward_sign=1.0,
        tip_half_len_ft=20.0, sheet_elements=elems)
    curved = [p for p in prims if p["kind"] == "dimension" and p.get("curved")]
    assert curved, "expected path-hugging dims when tip path bows"
    texts = [str(p.get("text") or "") for p in curved]
    assert any("495" in t or "120" in t or "680" in t for t in texts), texts
    for p in curved:
        assert p.get("path") and len(p["path"]) >= 2
        # Path length ≈ sheet length; chord would be much shorter across bend.
        tip1, tip2 = p["tip1"], p["tip2"]
        chord = math.hypot(tip2[0] - tip1[0], tip2[1] - tip1[1])
        path_len = sc._path_length(p["path"])
        assert path_len > chord + 1.0


def test_align2_closed_out_flips_tan():
    """Align2 tips must use −tan so they share Align1's closed shoulder."""
    ox, oy = sc._outward_unit(1.0, 0.0, 1.0)
    assert abs(ox) < 1e-9 and abs(oy - 1.0) < 1e-9
    ox2, oy2 = sc._outward_unit(-1.0, 0.0, 1.0)
    assert abs(ox2) < 1e-9 and abs(oy2 - (-1.0)) < 1e-9


def test_dim_tip_path_hugs_l_bend():
    segs = ag.segments_from_polyline(
        [[0.0, 0.0], [100.0, 0.0], [100.0, 100.0]])
    path = sc._dim_tip_path(segs, 50.0, 150.0, -1.0, sc.PERP_HALF_LEN_FT, step_ft=10.0)
    assert len(path) >= 5
    tip1, tip2 = path[0], path[-1]
    sag = sc._path_sagitta(path, tip1, tip2)
    assert sag > 5.0  # must bow away from the chord around the corner


def test_compile_hatch_curved_boundary_dense():
    path = _l_bend_path()
    segs = ag.segments_from_polyline(path)
    a1 = ag.segments_from_polyline([[200.0, 0.0], [0.0, 0.0]])
    a2 = ag.segments_from_polyline([[200.0, 200.0], [200.0, 400.0]])
    bay = ag.sample_path_vertices(segs, 200.0, 400.0, step_ft=5.0)
    spec = {
        "corridor": {"zones": [{"id": "workArea", "hatched": True}]},
        "symbols": {"items": []},
    }
    prims = sc.compile_hatch(
        spec, {}, a1, a2, lane_width_ft=12.0, outward_sign=1.0,
        work_bay_vertices=bay)
    hatch = next(p for p in prims if p["kind"] == "hatch")
    assert hatch["curvedWorkBay"] is True
    assert len(hatch["boundary"]) >= 20
