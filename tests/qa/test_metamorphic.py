"""Metamorphic tests: relations that must hold without knowing the right answer.

If the geometry helpers or the operation validator break one of these (translate, rotate, mirror twice, reverse, undo an
operation), a bug exists even though no hand-written expected value says so. Seeded random inputs; see CLAUDE.md QA rule 6/15.
"""
from __future__ import annotations

import copy
import json
import math
import os
import random
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa import geom, operations  # noqa: E402

LIVE = [json.loads(ln) for ln in open(os.path.join(ROOT, "tests", "qa", "fixtures", "oplog_live_2026-10-23.jsonl"), encoding="utf-8")]
RNG = random.Random(20261024)
EPS = 1e-6


def poly(n=None):
    n = n or RNG.randint(2, 7)
    return [(RNG.uniform(-500, 500), RNG.uniform(-500, 500)) for _ in range(n)]


def pt():
    return (RNG.uniform(-600, 600), RNG.uniform(-600, 600))


def near(a, b, tol=EPS):
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


# ------------------------------------------------------------------ geometry helpers
def test_rotate_then_rotate_back_is_identity_and_preserves_length():
    for _ in range(200):
        p, o, ang = pt(), pt(), RNG.uniform(-720, 720)
        q = geom.rotate(geom.rotate(p, o, ang), o, -ang)
        assert geom.dist(p, q) < 1e-7
        assert near(geom.dist(geom.rotate(p, o, ang), o), geom.dist(p, o))


def test_mirror_twice_is_identity_and_mirror_preserves_distance_to_axis():
    for _ in range(200):
        p, a, b = pt(), pt(), pt()
        if geom.dist(a, b) < 1:
            continue
        m = geom.mirror(p, a, b)
        assert geom.dist(geom.mirror(m, a, b), p) < 1e-7
        assert near(abs(geom.project_on_line(p, a, b)[1]), abs(geom.project_on_line(m, a, b)[1]))


def test_translation_preserves_length_and_nearest_distance():
    for _ in range(100):
        v, p = poly(), pt()
        dx, dy = RNG.uniform(-1e4, 1e4), RNG.uniform(-1e4, 1e4)
        v2 = [(x + dx, y + dy) for x, y in v]
        assert near(geom.length(v), geom.length(v2), 1e-9)
        try:
            d1 = geom.nearest_on_polyline(p, v)
        except ValueError:
            continue
        d2 = geom.nearest_on_polyline((p[0] + dx, p[1] + dy), v2)
        assert near(d1[0], d2[0], 1e-7) and near(d1[1], d2[1], 1e-7) and d1[2] == d2[2]


def test_reversing_a_polyline_keeps_length_and_distance_and_flips_the_side():
    for _ in range(100):
        v, p = poly(), pt()
        try:
            d1 = geom.nearest_on_polyline(p, v)
        except ValueError:
            continue
        d2 = geom.nearest_on_polyline(p, v[::-1])
        assert near(geom.length(v), geom.length(v[::-1]), 1e-9)
        assert near(d1[0], d2[0], 1e-7)
        at_vertex = abs(d1[0] - min(geom.dist(p, q) for q in v)) < 1e-6  # nearest point is a vertex: the side is ambiguous
        if d1[0] > 1e-6 and not at_vertex:
            assert d1[1] * d2[1] < 0, "left/right must swap when the line direction is reversed"


def test_uniform_scale_scales_distance_by_the_same_factor():
    for _ in range(100):
        v, p, s = poly(), pt(), RNG.uniform(0.2, 5)
        try:
            d1 = geom.nearest_on_polyline(p, v)[0]
        except ValueError:
            continue
        d2 = geom.nearest_on_polyline((p[0] * s, p[1] * s), [(x * s, y * s) for x, y in v])[0]
        assert near(d1 * s, d2, 1e-7)


def test_signed_offset_changes_sign_when_the_reference_direction_is_swapped():
    for _ in range(200):
        p, a, b = pt(), pt(), pt()
        if geom.dist(a, b) < 1:
            continue
        s1 = geom.project_on_line(p, a, b)[1]
        s2 = geom.project_on_line(p, b, a)[1]
        assert near(s1, -s2, 1e-7)


def test_quadrant_bearing_east_and_west_are_mirror_images_about_north():
    for deg in (1, 17, 45, 89):
        e = geom.bearing_to_azimuth(f"N {deg} 00 00 E")
        w = geom.bearing_to_azimuth(f"N {deg} 00 00 W")
        assert near(e + w, 360.0)
        se = geom.bearing_to_azimuth(f"S {deg} 00 00 E")
        sw = geom.bearing_to_azimuth(f"S {deg} 00 00 W")
        assert near(se + sw, 360.0)
        assert near(e + se, 180.0)


def test_box_to_polyline_distance_is_translation_invariant():
    for _ in range(60):
        v = poly()
        x, y = RNG.uniform(-300, 300), RNG.uniform(-300, 300)
        box = (x, y, x + RNG.uniform(1, 30), y + RNG.uniform(1, 30))
        dx, dy = RNG.uniform(-900, 900), RNG.uniform(-900, 900)
        d1 = geom.dist_box_polyline(box, v)
        d2 = geom.dist_box_polyline((box[0] + dx, box[1] + dy, box[2] + dx, box[3] + dy), [(a + dx, b + dy) for a, b in v])
        assert near(d1, d2, 1e-7)


# ------------------------------------------------------------------ operation validator
_X = ("originX", "x", "x1", "x2")
_Y = ("originY", "y", "y1", "y2")
_RANGE_X, _RANGE_Y = ("rangeLowX", "rangeHighX"), ("rangeLowY", "rangeHighY")


def translate_record(rec, dx, dy):
    """The same operation performed somewhere else in the file: every absolute coordinate moves, every delta stays."""
    r = copy.deepcopy(rec)
    for k in _X:
        if k in r["params"]:
            r["params"][k] = float(r["params"][k]) + dx
    for k in _Y:
        if k in r["params"]:
            r["params"][k] = float(r["params"][k]) + dy
    if "verticesTSV" in r["params"]:
        pts = [[float(c) for c in s.split(",")] for s in r["params"]["verticesTSV"].split("|")]
        r["params"]["verticesTSV"] = "|".join(f"{p[0] + dx},{p[1] + dy},{p[2] if len(p) > 2 else 0}" for p in pts)
    for side in ("before", "after"):
        for el in r[side].values():
            el["props"] = el.get("props") or {}
            el["verts"] = [[v[0] + dx, v[1] + dy] + list(v[2:]) for v in el.get("verts") or []]
            for k in _RANGE_X:
                if el["props"].get(k) not in (None, ""):
                    el["props"][k] = f"{float(el['props'][k]) + dx:.2f}"
            for k in _RANGE_Y:
                if el["props"].get(k) not in (None, ""):
                    el["props"][k] = f"{float(el['props'][k]) + dy:.2f}"
    return r


def test_validator_verdict_is_unchanged_when_the_whole_operation_is_translated():
    """Moving an operation to another place in the file must not change whether it is right. Run over every real record."""
    checked = 0
    for rec in LIVE:
        base = operations.validate(rec)["status"]
        for dx, dy in ((1000.0, -2500.0), (-60000.0, 12345.0)):
            moved = operations.validate(translate_record(rec, dx, dy))["status"]
            assert moved == base, (rec["op"], dx, dy, base, moved)
            checked += 1
    assert checked == 2 * len(LIVE)


def test_move_then_inverse_move_cancels_in_the_validator():
    """A MOVE by (dx,dy) followed by a MOVE by (-dx,-dy) leaves the element where it began: both records validate, and
    the second one starts where the first ended."""
    rec = next(r for r in LIVE if r["op"] == "MOVE_ELEMENT")
    eid = rec["params"]["elementId"]
    start = rec["before"][eid]["verts"]
    end = rec["after"][eid]["verts"]
    back = copy.deepcopy(rec)
    back["params"]["deltaX"], back["params"]["deltaY"] = -rec["params"]["deltaX"], -rec["params"]["deltaY"]
    back["before"], back["after"] = copy.deepcopy(rec["after"]), copy.deepcopy(rec["before"])
    assert operations.validate(rec)["status"] == "PASS"
    assert operations.validate(back)["status"] == "PASS"
    ok, _ = geom.same_pts(back["after"][eid]["verts"], start)
    assert ok and not geom.same_pts(start, end)[0]


def test_copy_parallel_distance_sign_decides_the_side_and_the_record_must_agree():
    """Flip the distance sign without flipping the geometry: the validator must notice (same relation, opposite side)."""
    rec = next(r for r in LIVE if r["op"] == "COPY_PARALLEL")
    flipped = copy.deepcopy(rec)
    flipped["params"]["distance"] = -rec["params"]["distance"]
    assert operations.validate(rec)["status"] == "PASS"
    assert operations.validate(flipped)["status"] != "PASS"
