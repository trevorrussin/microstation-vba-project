"""Operation contracts: did the edit do what the operation said it would?

Each contract takes one op record from outputs/qa/oplog.jsonl (mcp-server/op_recorder.py): the op, its exact
parameters, and the operand's geometry/properties BEFORE and the operand + new elements AFTER, read back from
MicroStation. The contract recomputes the postcondition from those two independent reads and compares it with the
parameters. It never trusts the agent's description or the op's own response.

status: PASS | FAIL | UNCERTAIN (data missing / not measurable)
"""
from __future__ import annotations

import math

from qa import geom
from qa.geom import TOL, ANG_TOL

# MicroStation's ROTATE_ELEMENT matrix convention vs the standard CCW-positive angle (calibrated live, see tests)
ROT_SIGN = 1.0


def _r(status, observed, expected="", evidence=None):
    return {"status": status, "observed": observed, "expected": expected, "evidence": evidence or []}


def _f(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _verts(rec, side, eid):
    return ((rec.get(side) or {}).get(str(eid)) or {}).get("verts")


def _props(rec, side, eid):
    return ((rec.get(side) or {}).get(str(eid)) or {}).get("props")


def _unc(what):
    return _r("UNCERTAIN", f"cannot verify: {what}", "before/after geometry read back")


def _operand(rec):
    return str(rec["params"].get("elementId") or "")


def _new(rec):
    return [str(i) for i in rec.get("new_ids") or []]


def _unchanged(rec, eid):
    b, a = _verts(rec, "before", eid), _verts(rec, "after", eid)
    if b is None or a is None:
        return None
    return geom.same_pts(b, a)[0]


# ---------------------------------------------------------------- transforms of the operand
def move(rec):
    p, eid = rec["params"], _operand(rec)
    b, a = _verts(rec, "before", eid), _verts(rec, "after", eid)
    if not b or not a:
        return _unc("operand geometry before/after not readable")
    dx, dy, dz = _f(p.get("deltaX")), _f(p.get("deltaY")), _f(p.get("deltaZ"))
    want = [[v[0] + dx, v[1] + dy, v[2] + dz] for v in b]
    ok, detail = geom.same_pts(want, a)
    got = (a[0][0] - b[0][0], a[0][1] - b[0][1])
    obs = f"moved ({got[0]:.3f}, {got[1]:.3f}) ft; commanded ({dx:g}, {dy:g}); {detail}"
    return _r("PASS" if ok else "FAIL", obs, f"every vertex displaced by ({dx:g}, {dy:g}) ft, shape preserved")


def rotate(rec):
    p, eid = rec["params"], _operand(rec)
    b, a = _verts(rec, "before", eid), _verts(rec, "after", eid)
    if not b or not a:
        return _unc("operand geometry before/after not readable")
    o = (_f(p.get("originX")), _f(p.get("originY")))
    ang = _f(p.get("angleDeg"))
    want = [list(geom.rotate(v, o, ROT_SIGN * ang)) + [v[2]] for v in b]
    ok, detail = geom.same_pts(want, a)
    return _r("PASS" if ok else "FAIL", f"rotate {ang:g} deg about ({o[0]:g}, {o[1]:g}): {detail}",
              "every vertex rotated counterclockwise (+) by the angle about the origin")


def scale(rec):
    p, eid = rec["params"], _operand(rec)
    b, a = _verts(rec, "before", eid), _verts(rec, "after", eid)
    if not b or not a:
        return _unc("operand geometry before/after not readable")
    o = (_f(p.get("originX")), _f(p.get("originY")))
    s = _f(p.get("scaleFactor"), 1.0)
    want = [[o[0] + s * (v[0] - o[0]), o[1] + s * (v[1] - o[1]), v[2]] for v in b]
    ok, detail = geom.same_pts(want, a)
    return _r("PASS" if ok else "FAIL", f"scale x{s:g} about ({o[0]:g}, {o[1]:g}): {detail}", "every vertex scaled about the origin")


def mirror(rec):
    p, eid = rec["params"], _operand(rec)
    b, a = _verts(rec, "before", eid), _verts(rec, "after", eid)
    if not b or not a:
        return _unc("operand geometry before/after not readable")
    p1, p2 = (_f(p.get("x1")), _f(p.get("y1"))), (_f(p.get("x2")), _f(p.get("y2")))
    if geom.dist(p1, p2) < 1e-9:
        return _r("FAIL", "mirror axis has zero length", "two distinct axis points")
    want = [list(geom.mirror(v, p1, p2)) + [v[2]] for v in b]
    ok, detail = geom.same_pts(want, a)
    return _r("PASS" if ok else "FAIL", f"mirror about axis: {detail}", "every vertex reflected across the axis")


def extend(rec):
    p, eid = rec["params"], _operand(rec)
    new = _new(rec)
    after_id = new[0] if new else eid  # EXTEND_LINE recreates the line under a new id
    b, a = _verts(rec, "before", eid), _verts(rec, "after", after_id)
    if not b or not a or len(b) < 2 or len(a) < 2:
        return _unc("line endpoints before/after not readable")
    want_len = _f(p.get("newLength"))
    got_len = geom.dist(a[0], a[-1])
    start_ok = geom.dist(b[0], a[0]) <= TOL
    dir_ok = geom.ang_diff(geom.heading_deg(b[0], b[-1]), geom.heading_deg(a[0], a[-1])) <= ANG_TOL if got_len > TOL else False
    ok = start_ok and dir_ok and abs(got_len - want_len) <= TOL
    return _r("PASS" if ok else "FAIL",
              f"length {got_len:.3f} ft (commanded {want_len:g}); start fixed: {start_ok}; direction kept: {dir_ok}",
              "same start point and direction, length equal to the commanded length")


# ---------------------------------------------------------------- copies
def copy(rec):
    p, eid = rec["params"], _operand(rec)
    new = _new(rec)
    b = _verts(rec, "before", eid)
    if not new or b is None or _verts(rec, "after", new[0]) is None:
        return _unc("copy geometry not readable")
    dx, dy, dz = _f(p.get("deltaX")), _f(p.get("deltaY")), _f(p.get("deltaZ"))
    want = [[v[0] + dx, v[1] + dy, v[2] + dz] for v in b]
    ok, detail = geom.same_pts(want, _verts(rec, "after", new[0]))
    orig = _unchanged(rec, eid)
    ok = ok and orig is not False
    return _r("PASS" if ok else "FAIL", f"copy offset ({dx:g}, {dy:g}): {detail}; original unchanged: {orig}",
              "copy = original translated by the delta; original untouched")


def copy_parallel(rec):
    p, eid = rec["params"], _operand(rec)
    new = _new(rec)
    b = _verts(rec, "before", eid)
    n = _verts(rec, "after", new[0]) if new else None
    if not b or not n or len(b) < 2 or len(n) < 2:
        return _unc("line endpoints before/after not readable")
    if len(b) > 2:
        return _r("UNCERTAIN", "offset of a multi-vertex element is not measurable by the line contract", "LINE operand")
    want = _f(p.get("distance"))
    o0, o1 = b[0], b[-1]
    (_, d0, t0), (_, d1, t1) = (geom.project_on_line(n[0], o0, o1), geom.project_on_line(n[-1], o0, o1))
    L = geom.dist(o0, o1)
    parallel = abs(d0 - d1) <= TOL
    dist_ok = abs(d0 - want) <= TOL  # signed: + = left of start->end
    extent_ok = abs(t0) <= 0.05 and abs(t1 - L) <= 0.05
    orig = _unchanged(rec, eid)
    ok = parallel and dist_ok and extent_ok and orig is not False
    return _r("PASS" if ok else "FAIL",
              f"measured offset {d0:.3f} ft ({'left' if d0 >= 0 else 'right'}); commanded {want:g}; parallel: {parallel}; "
              f"same extent: {extent_ok}; original unchanged: {orig}",
              "copy parallel to the original, perpendicular distance = commanded (+ = left of start->end), same extent")


def array(rec):
    p, eid = rec["params"], _operand(rec)
    new = _new(rec)
    b = _verts(rec, "before", eid)
    n = int(_f(p.get("count")))
    if b is None or any(_verts(rec, "after", i) is None for i in new):
        return _unc("array geometry not readable")
    sx, sy = _f(p.get("spacingX")), _f(p.get("spacingY"))
    if len(new) != n:
        return _r("FAIL", f"{len(new)} copies created, {n} commanded", f"{n} copies")
    bad = []
    for k, i in enumerate(new, start=1):
        want = [[v[0] + k * sx, v[1] + k * sy, v[2]] for v in b]
        ok, detail = geom.same_pts(want, _verts(rec, "after", i))
        if not ok:
            bad.append(f"copy {k} (id {i}): {detail}")
    return _r("FAIL" if bad else "PASS", f"{len(new)} copies at i*({sx:g}, {sy:g}); {len(bad)} wrong",
              "copy i is the original translated by i*(spacingX, spacingY)", bad)


def break_line(rec):
    p, eid = rec["params"], _operand(rec)
    new = _new(rec)
    b = _verts(rec, "before", eid)
    if b is None or len(new) != 2:
        return _unc("break pieces not readable")
    c1, c2 = _verts(rec, "after", new[0]), _verts(rec, "after", new[1])
    if not c1 or not c2:
        return _unc("break pieces not readable")
    pt = (_f(p.get("x")), _f(p.get("y")))
    foot, off, _ = geom.project_on_line(pt, b[0], b[-1])
    joint_ok = geom.dist(c1[-1], c2[0]) <= TOL
    ends_ok = geom.dist(c1[0], b[0]) <= TOL and geom.dist(c2[-1], b[-1]) <= TOL
    on_line = abs(geom.project_on_line(c1[-1], b[0], b[-1])[1]) <= TOL
    at_pt = geom.dist(c1[-1], pt) <= TOL
    len_ok = abs(geom.dist(c1[0], c1[-1]) + geom.dist(c2[0], c2[-1]) - geom.dist(b[0], b[-1])) <= TOL
    ev = []
    if not on_line:
        ev.append(f"break point is {abs(off):.3f} ft off the original line, so the two pieces form a kink")
    ok = joint_ok and ends_ok and on_line and at_pt and len_ok
    return _r("PASS" if ok else "FAIL",
              f"pieces meet: {joint_ok}; ends kept: {ends_ok}; break point on original line: {on_line}; "
              f"at requested point: {at_pt}; length preserved: {len_ok}",
              "two collinear pieces joined at the requested point on the line, total length unchanged", ev)


# ---------------------------------------------------------------- properties
def level(rec):
    p, eid = rec["params"], _operand(rec)
    a = _props(rec, "after", eid)
    if not isinstance(a, dict):
        return _unc("element properties not readable after the edit")
    want = str(p.get("level") or "")
    ok = (a.get("level") or "").lower() == want.lower() and _unchanged(rec, eid) is not False
    return _r("PASS" if ok else "FAIL", f"level now {a.get('level')!r} (commanded {want!r}); geometry unchanged: {_unchanged(rec, eid)}",
              "element on the commanded level, geometry untouched")


def symbology(rec):
    p, eid = rec["params"], _operand(rec)
    a = _props(rec, "after", eid)
    if not isinstance(a, dict):
        return _unc("element properties not readable after the edit")
    bad = []
    if "color" in p and str(int(_f(p["color"]))) != str(a.get("color")):
        bad.append(f"color {a.get('color')} (commanded {p['color']})")
    if "weight" in p and str(int(_f(p["weight"]))) != str(a.get("weight")):
        bad.append(f"weight {a.get('weight')} (commanded {p['weight']})")
    if _unchanged(rec, eid) is False:
        bad.append("geometry changed")
    return _r("FAIL" if bad else "PASS", "; ".join(bad) or f"color {a.get('color')} weight {a.get('weight')} as commanded",
              "commanded color/weight set, geometry untouched", bad)


def edit_text(rec):
    p, eid = rec["params"], _operand(rec)
    a, b = _props(rec, "after", eid), _props(rec, "before", eid)
    if not isinstance(a, dict) or not isinstance(b, dict):
        return _unc("text properties before/after not readable")
    bad = []
    if (a.get("text") or "").strip() != str(p.get("newText") or "").strip():
        bad.append(f"text is {a.get('text')!r}, commanded {p.get('newText')!r}")
    for k in ("textHeight", "textWidth", "textFont"):
        if str(a.get(k)) != str(b.get(k)):
            bad.append(f"{k} changed {b.get(k)} -> {a.get(k)}")
    for k in ("rangeLowX", "rangeLowY"):
        if abs(_f(a.get(k)) - _f(b.get(k))) > 0.5:  # edit may change the box width, not the anchor
            bad.append(f"text moved ({k} {b.get(k)} -> {a.get(k)})")
    return _r("FAIL" if bad else "PASS", "; ".join(bad) or "text replaced, style and anchor unchanged",
              "new text content, same height/width/font, same anchor", bad)


def delete(rec):
    eid = _operand(rec)
    a = ((rec.get("after") or {}).get(eid) or {}).get("props", "unreadable")
    if a == "unreadable":
        return _unc("could not confirm the element is gone")
    return _r("PASS" if a is None else "FAIL", "element removed" if a is None else "element still exists", "element no longer in the model")


# ---------------------------------------------------------------- primitive placements
def place_polyline(rec):
    new = _new(rec)
    got = _verts(rec, "after", new[0]) if new else None
    if not got:
        return _unc("placed geometry not readable")
    want = []
    for part in str(rec["params"].get("verticesTSV") or "").split("|"):
        xs = part.split(",")
        want.append([_f(xs[0]), _f(xs[1]), _f(xs[2]) if len(xs) > 2 else 0.0])
    ok, detail = geom.same_pts(want, got)
    return _r("PASS" if ok else "FAIL", f"placed polyline: {detail}", "vertices equal the requested vertices")


def place_circle(rec):
    p = rec["params"]
    new = _new(rec)
    got = _verts(rec, "after", new[0]) if new else None
    if not got:
        return _unc("placed geometry not readable")
    c, r = (_f(p.get("cx")), _f(p.get("cy"))), _f(p.get("radius"))
    worst = max(abs(geom.dist(v, c) - r) for v in got)
    x0, y0, x1, y1 = geom.bbox(got)
    centre_err = geom.dist(((x0 + x1) / 2, (y0 + y1) / 2), c)
    ok = worst <= 0.05 and centre_err <= 0.05
    return _r("PASS" if ok else "FAIL", f"radius error {worst:.3f} ft, centre error {centre_err:.3f} ft",
              f"circle centred ({c[0]:g}, {c[1]:g}) radius {r:g}")


def place_arc(rec):
    p = rec["params"]
    new = _new(rec)
    got = _verts(rec, "after", new[0]) if new else None
    if not got or len(got) < 3:
        return _unc("placed geometry not readable")
    worst = 0.0
    for k in (1, 2, 3):
        q = (_f(p.get(f"x{k}")), _f(p.get(f"y{k}")))
        worst = max(worst, min(geom.dist(q, v) for v in got))
    ok = worst <= 0.5  # densified arc vertices: nearest-vertex distance, generous but catches wrong points
    return _r("PASS" if ok else "FAIL", f"farthest of the 3 defining points from the arc: {worst:.3f} ft",
              "the three defining points lie on the placed arc")


def place_block(rec):
    p = rec["params"]
    new = _new(rec)
    got = _verts(rec, "after", new[0]) if new else None
    if not got:
        return _unc("placed geometry not readable")
    want = (min(_f(p.get("x1")), _f(p.get("x2"))), min(_f(p.get("y1")), _f(p.get("y2"))),
            max(_f(p.get("x1")), _f(p.get("x2"))), max(_f(p.get("y1")), _f(p.get("y2"))))
    have = geom.bbox(got)
    err = max(abs(a - b) for a, b in zip(want, have))
    return _r("PASS" if err <= TOL else "FAIL", f"corner error {err:.3f} ft", "rectangle spans the two requested corners")


def place_text(rec):
    p = rec["params"]
    new = _new(rec)
    a = _props(rec, "after", new[0]) if new else None
    if not isinstance(a, dict):
        return _unc("placed text not readable")
    bad = []
    if (a.get("text") or "").strip() != str(p.get("text") or "").strip():
        bad.append(f"text is {a.get('text')!r}, commanded {p.get('text')!r}")
    x, y = _f(p.get("x")), _f(p.get("y"))
    lo = (_f(a.get("rangeLowX")), _f(a.get("rangeLowY")))
    hi = (_f(a.get("rangeHighX")), _f(a.get("rangeHighY")))
    if not (lo[0] - 1.0 <= x <= hi[0] + 1.0 and lo[1] - 1.0 <= y <= hi[1] + 1.0):
        bad.append(f"requested anchor ({x:g}, {y:g}) is not at the text box ({lo[0]:.1f},{lo[1]:.1f})-({hi[0]:.1f},{hi[1]:.1f})")
    want_ang = _f(p.get("angleDeg"))
    got_ang = _f(a.get("textRotDeg"))
    if a.get("textRotDeg") not in (None, "") and geom.ang_diff(want_ang, got_ang) > 0.5:
        bad.append(f"rotation {got_ang:g} deg (commanded {want_ang:g})")
    return _r("FAIL" if bad else "PASS", "; ".join(bad) or "text content, anchor and rotation as commanded",
              "requested text at the requested anchor and angle", bad)


CONTRACTS = {
    "MOVE_ELEMENT": move, "COPY_ELEMENT": copy, "COPY_PARALLEL": copy_parallel, "ROTATE_ELEMENT": rotate,
    "SCALE_ELEMENT": scale, "MIRROR_ELEMENT": mirror, "ARRAY_ELEMENT": array, "EXTEND_LINE": extend,
    "BREAK_LINE": break_line, "CHANGE_ELEMENT_LEVEL": level, "CHANGE_ELEMENT_SYMBOLOGY": symbology,
    "EDIT_TEXT_ELEMENT": edit_text, "DELETE_ELEMENT": delete, "PLACE_POLYLINE": place_polyline,
    "PLACE_CIRCLE": place_circle, "PLACE_ARC": place_arc, "PLACE_BLOCK": place_block, "PLACE_TEXT_LABEL": place_text,
}


def validate(rec):
    fn = CONTRACTS.get(rec.get("op"))
    if fn is None:
        return _r("UNCERTAIN", f"no contract for {rec.get('op')}", "a validator for this operation")
    try:
        return fn(rec)
    except Exception as ex:  # a broken validator must never read as PASS
        return _r("UNCERTAIN", f"validator error for {rec.get('op')}: {ex}", "validator runs")


def numbers_in_params(rec) -> list[float]:
    """Absolute numeric parameters of an op (for the intent cross-check against the engineer's request)."""
    out = []
    for k, v in (rec.get("params") or {}).items():
        if k in ("elementId", "ownElementOnly", "reason", "verticesTSV", "level", "newText", "text"):
            continue
        try:
            out.append(abs(float(v)))
        except (TypeError, ValueError):
            pass
    return out
