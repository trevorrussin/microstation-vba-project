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


def test_overlay_dim_stays_off_pavement_on_real_road():
    """SHOULDER TAPER (Overlay) must not flip across the travel lanes.

    619-311's annotationStyle sets overlayDimSide="opposite", which is a
    printed-sheet convention — on the schematic the other side is blank
    paper. On a real road the alignment is the closed-lane edge, so an
    unconditional flip drives the overlay dim through the pavement. Live
    C-curve build 2026-08-13: SHOULDER TAPER tipped at align-20 (mid-lane)
    while every other dim tipped at align+20 (on the EOP).

    Contract: when a real-road tip_half_len_ft is locked, EVERY dimension —
    overlay included — tips on the same side as the main column.
    """
    import sheet_spec as ss

    segs = ag.segments_from_polyline([[0.0, 0.0], [3000.0, 0.0]])
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    assert (real.get("annotationStyle") or {}).get("overlayDimSide") == "opposite"
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    # get_sheet_requirements' pipe-list (Data/sheet-registry.tsv); the spec
    # JSON has no sheet.elements key, and "" filters the overlay row out.
    elems = "MergingTaper|ShoulderTaper|DownstreamTaper"
    prims = sc.compile_plan(
        real, resolved, 1, segs, outward_sign=1.0,
        tip_half_len_ft=20.0, sheet_elements=elems)

    overlays = [p for p in prims
                if p["kind"] == "dimension" and (p.get("specRef") or {}).get("overlay")]
    mains = [p for p in prims
             if p["kind"] == "dimension" and not (p.get("specRef") or {}).get("overlay")]
    assert overlays, "expected at least one overlay dim (SHOULDER TAPER)"
    assert mains

    # outward_sign=+1 on a due-east align puts tips at +Y. Every tip, overlay
    # or not, must be on that same side — never negative (across the road).
    for p in mains + overlays:
        for tip in (p["tip1"], p["tip2"]):
            assert tip[1] > 0.0, f"dim tip crossed the alignment: {p.get('text')} {tip}"

    # Overlay tips land on the same offset as the main column (the EOP)...
    main_tip_y = mains[0]["tip1"][1]
    for p in overlays:
        assert abs(p["tip1"][1] - main_tip_y) < 0.05
    # ...and its dim line clears the main column so they do not overlap.
    assert overlays[0]["offset"][1] > mains[0]["offset"][1] + 1.0


def test_overlay_dim_still_flips_on_schematic_build():
    """Without a real-road tip_half_len_ft, overlayDimSide='opposite' stands."""
    import sheet_spec as ss

    segs = ag.segments_from_polyline([[0.0, 0.0], [3000.0, 0.0]])
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    elems = "MergingTaper|ShoulderTaper|DownstreamTaper"
    prims = sc.compile_plan(
        real, resolved, 1, segs, outward_sign=1.0, sheet_elements=elems)
    overlays = [p for p in prims
                if p["kind"] == "dimension" and (p.get("specRef") or {}).get("overlay")]
    assert overlays
    # Schematic keeps the printed-sheet flip: overlay tips on the far side.
    assert overlays[0]["tip1"][1] < 0.0


def test_feature_labels_carry_tangent_angle():
    segs = ag.segments_from_polyline(
        [[0.0, 0.0], [50.0, 0.0], [50.0, 2500.0]])
    import sheet_spec as ss
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    elems = (real.get("sheet") or {}).get("elements", "")
    prims = sc.compile_plan(
        real, resolved, 1, segs, outward_sign=1.0,
        tip_half_len_ft=20.0, sheet_elements=elems)
    labels = [p for p in prims if p["kind"] == "label" and p.get("text") != "ARROW PANEL"]
    assert labels
    assert all("angleDeg" in p for p in labels)
    assert all(-90.0 <= float(p["angleDeg"]) <= 90.0 for p in labels)
    # Northbound tangent on the bent leg → ~90° (not flipped; 90 is kept).
    # Northbound tangent on the bent leg → ~90°.
    northish = [p for p in labels if abs(float(p["angleDeg"]) - 90.0) < 15.0]
    assert northish, [p.get("angleDeg") for p in labels]


def test_protective_vehicle_angle_is_tangent_plus_180():
    segs = ag.segments_from_polyline([[0.0, 0.0], [500.0, 0.0]])
    import sheet_spec as ss
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    prims = sc.compile_symbols(real, resolved, 1, segs, outward_sign=1.0)
    pvs = [p for p in prims if p.get("kind") == "protectiveVehicle"]
    assert pvs
    for p in pvs:
        assert abs((p["angleDeg"] % 360.0) - 180.0) < 1e-6


def test_signpost_angle_follows_travel_tangent():
    """TWZSGN_P uses travel, not raw Align1 tan (which points upstream)."""
    # Align1 tan away-upstream = west on an eastbound approach.
    assert abs(sc._post_angle_deg(1, -1.0, 0.0) - 0.0) < 1e-6
    # Align2 tan away-downstream = north after a 90° left bend.
    assert abs(sc._post_angle_deg(2, 0.0, 1.0) - 90.0) < 1e-6
    tx, ty = sc._travel_unit(1, -1.0, 0.0)
    assert abs(tx - 1.0) < 1e-9 and abs(ty) < 1e-9


def test_label_angle_flips_when_upside_down():
    """Keep tangent; fold past ±90 so lettering is not inverted."""
    assert abs(sc._text_angle_deg(1.0, 0.0) - 0.0) < 1e-6
    assert abs(sc._text_angle_deg(0.0, 1.0) - 90.0) < 1e-6
    assert abs(sc._text_angle_deg(0.0, -1.0) - (-90.0)) < 1e-6
    assert abs(sc._text_angle_deg(-1.0, 0.0) - 0.0) < 1e-6  # 180 → 0
    a = sc._text_angle_deg(-0.1, 1.0)  # slightly past +90 CCW
    assert -90.0 <= a < 90.0
    assert a < 0.0


def test_split_reverse_s_parts_sum_to_sheet():
    """S-curve 680' lane taper → real SizeArrow + Arc Size pieces that sum to 680."""
    import lane_highway as lh

    outer = [
        [0.0, 0.0], [2200.0, 0.0], [2800.0, 600.0],
        [3400.0, 0.0], [5600.0, 0.0],
    ]
    segs = lh._prepare_path_segments(
        lh.resolve_edge_path(0, 0, 0, 0, outer), fillet_radius_ft=150.0)
    align = lh._continuous_offset_polyline(segs, 38.0, "right", step_ft=25.0)
    aseg = ag.segments_from_polyline(align)
    total = ag.total_length(aseg)
    sta1 = total * 0.5
    sta0 = sta1 - 680.0
    path = sc._dim_tip_path(aseg, sta0, sta1, 1.0, 20.0, step_ft=10.0)
    kind = sc.classify_dim_path(path)
    assert kind == "compound", kind
    runs = sc.split_dim_path_runs(path)
    assert len(runs) >= 2, runs
    kinds = {k for k, _ in runs}
    assert "straight" in kinds and "arc" in kinds
    raw = [sc._path_length(pts) for _, pts in runs]
    parts = sc._apportion_sheet_lengths(raw, 680.0)
    assert abs(sum(parts) - 680.0) < 0.15
    assert all(p > 0.0 for p in parts)


def test_short_highway_curve_classifies_as_arc():
    """50' downstream on R=3000 must be one Arc Size, not a SizeArrow chord."""
    r = 3000.0
    pts = []
    span = 50.0 / r
    for i in range(8):
        a = span * i / 7.0
        pts.append((r * math.sin(a), r * (1.0 - math.cos(a))))
    assert sc.classify_dim_path(pts) == "arc"


def test_c_curve_long_span_stays_one_arc():
    """Constant-R C-curve 680' must not split into fake straight crumbs."""
    r = 3000.0
    span = 680.0 / r
    pts = []
    n = 70
    for i in range(n + 1):
        a = span * i / n
        pts.append((r * math.sin(a), r * (1.0 - math.cos(a))))
    assert sc.classify_dim_path(pts) == "arc"
    runs = sc.split_dim_path_runs(pts)
    long_st = [p for k, p in runs if k == "straight" and sc._path_length(p) >= 80]
    assert not long_st


def test_compile_compound_emits_part_lengths():
    import lane_highway as lh
    import sheet_spec as ss

    outer = [
        [0.0, 0.0], [2200.0, 0.0], [2800.0, 600.0],
        [3400.0, 0.0], [5600.0, 0.0],
    ]
    segs = lh._prepare_path_segments(
        lh.resolve_edge_path(0, 0, 0, 0, outer), fillet_radius_ft=150.0)
    align = lh._continuous_offset_polyline(segs, 38.0, "right", step_ft=25.0)
    aseg = ag.segments_from_polyline(align)
    total = ag.total_length(aseg)
    sta_wa = total * 0.5
    a1_verts = []
    sta = sta_wa
    while sta >= 0 and (sta_wa - sta) <= 2500.0:
        x, y, _, _ = ag.point_at_extended(aseg, sta)
        a1_verts.append([x, y])
        sta -= 10.0
    a1 = ag.segments_from_polyline(a1_verts)
    real = ss.load("619-311")
    if real is None:
        pytest.skip("619-311 spec missing")
    resolved = ss.resolve(real, 55, 12, ">= 8 ft", "URBAN", None, None)
    elems = (real.get("sheet") or {}).get("elements", "")
    prims = sc.compile_plan(
        real, resolved, 1, a1, outward_sign=1.0,
        tip_half_len_ft=20.0, sheet_elements=elems)
    lane = [p for p in prims
            if p["kind"] == "dimension"
            and (p.get("specRef") or {}).get("zone") == "laneTaper"]
    assert lane, "expected laneTaper dims"
    n = int((lane[0].get("specRef") or {}).get("partCount") or 1)
    assert n >= 2, [(p.get("specRef"), p.get("pathKind")) for p in lane]
    parts = [float((p.get("specRef") or {}).get("partLengthFt") or 0) for p in lane]
    sheet = float((lane[0].get("specRef") or {}).get("sheetLengthFt") or 0)
    assert abs(sum(parts) - sheet) < 0.2
    assert abs(sum(parts) - float((lane[0].get("specRef") or {}).get("partsSumFt") or 0)) < 0.2


def test_arc_dim_line_radius_inside_of_curve_stays_off_pavement():
    """Closed shoulder on the inside of a bend: dim at r-pad, not r+pad.

    r+pad is through the travel lanes (live S-curve roll/downstream 2026-08-14).
    """
    import wztc_ops as ops

    cx, cy, r = 0.0, 0.0, 100.0
    mid = (100.0, 0.0)
    r_in = ops.arc_dim_line_radius(cx, cy, r, mid, 85.0, 0.0, pad=15.0)
    assert r_in < r, r_in
    assert abs(r_in - 85.0) < 0.05
    r_out = ops.arc_dim_line_radius(cx, cy, r, mid, 115.0, 0.0, pad=15.0)
    assert r_out > r, r_out
    assert abs(r_out - 115.0) < 0.05

