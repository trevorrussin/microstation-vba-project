"""Operation-level acceptance: did the requested EDIT do what was asked (not just: is the result well-formed CAD)?

operations(task, ctx)  OP-01       every recorded edit op is re-measured from before/after geometry read back by
                                   mcp-server/op_recorder.py and compared with the op's own parameters (qa/operations.py);
                                   numbers/directions in the engineer's request must match an op parameter;
                                   changes made outside a recorded op cannot pass without an independent expectation.
relations(task, ctx)   REQ-GEOM-01 geometric relations declared with `qa_gate expect --rel ...` are measured from the
                                   saved geometry (offset between two lines, element length, angle between lines).
"""
from __future__ import annotations

import re

from qa import geom
from qa import operations as opcontracts

MAX_IDS = 25


def _res(status, observed="", expected="", evidence=None, fix=""):
    return {"status": status, "observed": observed, "expected": expected, "evidence": (evidence or [])[:MAX_IDS], "fix": fix}


_NUM = re.compile(r"(\d+(?:\.\d+)?)\s*(?:ft|feet|foot|')", re.I)
_DEG = re.compile(r"(\d+(?:\.\d+)?)\s*(?:deg|degrees|°)", re.I)


_SCALE = re.compile(r"\d+(?:\.\d+)?\s*(?:in|inch|\")\s*=\s*\d+(?:\.\d+)?\s*(?:ft|feet|foot|')", re.I)


def _desc_numbers(desc: str) -> list[float]:
    desc = _SCALE.sub(" ", desc or "")  # a drawing scale (1in=40ft) names a scale, not a distance an operation must use
    return [float(x) for x in _NUM.findall(desc)] + [float(x) for x in _DEG.findall(desc)]


MAGNITUDE_OPS = {"MOVE_ELEMENT", "COPY_ELEMENT", "COPY_PARALLEL", "ROTATE_ELEMENT", "SCALE_ELEMENT", "EXTEND_LINE", "ARRAY_ELEMENT"}


def intent_lines(desc: str) -> list[str]:
    """What the gate will hold the operations to, derived from the request text (printed at `qa_gate start` so a wrong reading is caught first)."""
    d = (desc or "").lower()
    nums = _desc_numbers(desc)
    comp = [w for w in ("east", "west", "north", "south") if re.search(rf"\b{w}\b", d)]
    side = [w for w in ("left", "right") if re.search(rf"\b{w}\b", d)]
    m = re.search(r"offset[^0-9+\-]{0,25}([+-]\d+(?:\.\d+)?)", d)
    out = [f"quantities in the request: {', '.join(f'{n:g}' for n in nums) or 'NONE'} (each must be used by an operation)",
           f"directions: {', '.join(comp + side) or 'none stated'}"
           + (f"; ORD signed offset {m.group(1)} ({'left' if float(m.group(1)) < 0 else 'right'} of the baseline; copy_parallel sign is opposite)" if m else "")]
    if not nums and not re.search(r"\d", desc or ""):
        out.append("NO number in the request: a move/offset/rotate/extend/copy with a size the agent chose is UNCERTAIN (INTENT) unless you sign it off")
    out.append("anything geometric you need guaranteed beyond this: qa_gate expect --rel offset|offset_curve|distance|side|vclear ...")
    return out


def _direction_notes(desc: str, recs) -> list[str]:
    """Compass words must agree with the sign of a single MOVE; left/right with a single COPY_PARALLEL."""
    out = []
    d = (desc or "").lower()
    comp = [w for w in ("east", "west", "north", "south") if re.search(rf"\b{w}\b", d)]
    moves = [r for r in recs if r["op"] == "MOVE_ELEMENT"]
    if len(comp) == 1 and len(moves) == 1:
        p = moves[0]["params"]
        dx, dy = float(p.get("deltaX") or 0), float(p.get("deltaY") or 0)
        ok = {"east": dx > 0, "west": dx < 0, "north": dy > 0, "south": dy < 0}[comp[0]]
        if not ok:
            out.append(f"request says {comp[0]} but the move was dx={dx:g}, dy={dy:g}")
    side = [w for w in ("left", "right") if re.search(rf"\b{w}\b", d)]
    offs = [r for r in recs if r["op"] == "COPY_PARALLEL"]
    if len(side) == 1 and len(offs) == 1:
        dist = float(offs[0]["params"].get("distance") or 0)
        if (side[0] == "left") != (dist > 0):
            out.append(f"request says {side[0]} but the offset distance was {dist:g} (+ = left of start->end)")
    # ORD convention (tutorial playbook offsets-tapers-geometry S2): offset NEGATIVE = left, POSITIVE = right.
    # copy_parallel is the opposite (+ = left), so a signed ORD offset in the request flips the tool's sign.
    m = re.search(r"offset[^0-9+\-]{0,25}([+-]\d+(?:\.\d+)?)", d)
    if m and len(offs) == 1:
        ord_off = float(m.group(1))
        dist = float(offs[0]["params"].get("distance") or 0)
        if (ord_off < 0) != (dist > 0):
            out.append(f"request gives the ORD offset {ord_off:g} ({'left' if ord_off < 0 else 'right'} of the baseline; ORD negative = left) "
                       f"but copy_parallel distance {dist:g} is {'left' if dist > 0 else 'right'} (+ = left)")
    return out


def operations(task, ctx):
    recs = ctx.get("oplog")
    if recs is None:
        return _res("NA", "no operation log position recorded for this task (started before operation recording)")
    ch = ctx.get("changes") or {}
    changed = {str(r["elementId"]) for r in ch.get("created", []) + ch.get("modified", [])} | {str(i) for i in ch.get("deleted", {})}
    if not changed:
        changed = {str(e["elementId"]) for e in ctx.get("elements") or []}
    contract = [r for r in recs if r.get("contract")]
    other = [r for r in recs if not r.get("contract")]
    recorded = set()
    for r in recs:
        recorded |= {str(k) for k in (r.get("before") or {})} | {str(k) for k in (r.get("after") or {})} | {str(i) for i in r.get("new_ids") or []}
    ev, fails, uncs, passed, nfail = [], [], [], 0, 0
    for r in contract:
        v = opcontracts.validate(r)
        line = f"{r['op']} {r['params'].get('elementId', '')} -> {v['status']}: {v['observed']}"
        if v["status"] == "PASS":
            passed += 1
        elif v["status"] == "FAIL":
            nfail += 1
            fails.append(line)
            fails.extend(f"    {e}" for e in v.get("evidence") or [])
        else:
            uncs.append(line)
        ev.append(line)
    # intent cross-check: every number in the request must be used by some operation
    notes = []
    if contract:
        used = [n for r in contract for n in opcontracts.numbers_in_params(r)]
        miss = [n for n in _desc_numbers(task.get("desc", "")) if not any(abs(n - u) <= 0.01 for u in used)]
        if miss:
            notes.append("request mentions " + ", ".join(f"{m:g}" for m in miss) + " but no operation used that value")
        notes += _direction_notes(task.get("desc", ""), contract)
        mags = [r for r in contract if r["op"] in MAGNITUDE_OPS]
        if mags and (task.get("desc") or "").strip() and not re.search(r"\d", task["desc"]):
            so = ctx.get("signoff") or {}
            if "INTENT" not in (so.get("items") or []):
                notes.append("the request gives no distance, angle or count, so the agent chose the size of: "
                             + ", ".join(sorted({r["op"] for r in mags})))
                ctx.setdefault("needs_signoff", []).append(
                    {"id": "INTENT", "why": "an edit's size was chosen by the agent, not given in the request", "kind": "signoff"})
    rel_ids = {str(e[k]) for e in task.get("expect") or [] if e.get("rel") for k in ("a", "b") if e.get(k)}
    unrecorded = sorted(changed - recorded - rel_ids)  # a declared relation (REQ-GEOM-01) is an independent check of that element
    if fails:
        return _res("FAIL", f"{nfail} of {len(contract)} recorded edit operations did not do what they were commanded to do",
                    "each edit re-measured from before/after geometry matches its parameters", fails + notes + [x for x in ev if "-> PASS" in x],
                    "Undo or redo the failing operation (undo_last_op), then re-run. Do not edit the expectation to fit the result.")
    if notes:
        return _res("UNCERTAIN", "the operations do not match the engineer's request: " + "; ".join(notes),
                    "numbers and directions in the request appear in the operations", ev + notes,
                    "Confirm with the engineer what distance/direction was meant, then redo the operation or correct the request.")
    if uncs:
        return _res("UNCERTAIN", f"{len(uncs)} recorded edit operations could not be measured", "before/after geometry readable", uncs,
                    "Re-run the edit with MicroStation open, or have the engineer review it.")
    if unrecorded:
        so = ctx.get("signoff") or {}
        if "OP-01" in (so.get("items") or []):
            res = _res("PASS", f"{len(unrecorded)} changed elements were not made by a recorded operation; the engineer signed them off",
                       evidence=[f"signed by {so.get('by')} {so.get('at')}: {so.get('reason', '')}"] + [f"element {i}" for i in unrecorded])
            res["signed_off_from"] = {"status": "UNCERTAIN", "observed": f"{len(unrecorded)} changed elements were not made by a recorded operation"}
            return res
        ctx.setdefault("needs_signoff", []).append(
            {"id": "OP-01", "why": f"{len(unrecorded)} changed elements were not made by a recorded operation (direct script edit)", "kind": "signoff"})
        return _res("UNCERTAIN", f"{len(unrecorded)} changed elements were not made by a recorded operation, so there is no independent "
                                 f"check that the edit is right",
                    "an edit made through the recorded tools, or an independent geometry expectation (qa_gate expect --rel ...)",
                    [f"element {i}" for i in unrecorded],
                    "Make the edit with the bridge tools (move_element, copy_parallel, place_cell, ...), declare a measurable relation "
                    "(`qa_gate expect --rel offset --a ID --b ID --value 12`), or have the engineer review it and run `qa_gate signoff --items OP-01`.")
    if other and not contract:
        return _res("WARN", f"{len(other)} operations have no geometric contract (composite builders): not verified here",
                    "a contract per operation", sorted({o["op"] for o in other}))
    if not contract:
        return _res("NA", "no edit operations recorded")
    extra = []
    if other:
        extra.append("operations without a contract: " + ", ".join(sorted({o['op'] for o in other})))
    if extra:
        return _res("WARN", f"{passed} of {len(contract)} recorded edit operations verified from before/after geometry", "", ev + extra)
    return _res("PASS", f"{passed} of {len(contract)} recorded edit operations verified from before/after geometry", evidence=ev)


# ---------------------------------------------------------------- declared relations (REQ-GEOM-01)
REL_KINDS = ("offset", "length", "angle", "offset_curve", "distance", "side", "vclear")


def _line_ends(vs):
    return (vs[0], vs[-1]) if vs and len(vs) >= 2 else None


def _xy(v):
    return (float(v[0]), float(v[1]))


def _measure(rel, a, b):
    if rel == "length":
        return geom.length(a), "ft"
    la, lb = _line_ends(a), _line_ends(b or [])
    if not la or not lb:
        return None, ""
    if rel == "angle":
        d = geom.ang_diff(geom.heading_deg(*la), geom.heading_deg(*lb))
        return min(d, 180 - d), "deg"
    if rel == "offset":
        (_, d0, _t), (_, d1, _t2) = geom.project_on_line(lb[0], *la), geom.project_on_line(lb[1], *la)
        if abs(d0 - d1) > 0.05:
            return None, "not parallel"
        return abs(d0), "ft"
    return None, ""


def _rel_offset_curve(e, g, boxes):
    """b is a constant perpendicular offset from the (possibly curved) reference a: every sample along b is `value` from a, on one side."""
    ref, off = [_xy(v) for v in g.get(str(e["a"])) or []], [_xy(v) for v in g.get(str(e["b"])) or []]
    if len(ref) < 2 or len(off) < 2:
        return "UNC", "reference or offset geometry not readable"
    want, tol = float(e["value"]), float(e.get("tol") or 0.05)
    samples = [geom.nearest_on_polyline(p, ref)[1] for p in geom.sample_polyline(off)]
    sides = {s > 0 for s in samples if abs(s) > 1e-6}
    worst = max(abs(abs(s) - abs(want)) for s in samples)
    side = "left" if samples[len(samples) // 2] > 0 else "right"
    want_side = e.get("side")
    ok = worst <= tol and len(sides) <= 1 and (not want_side or want_side == side)
    return ("OK" if ok else "BAD"), (f"offset_curve: {len(samples)} samples along b are {min(abs(s) for s in samples):.3f}-"
                                      f"{max(abs(s) for s in samples):.3f} ft from the reference ({side}); expected {abs(want):g} +/- {tol:g}"
                                      + (f", {want_side}" if want_side else "") + ("; crosses the reference" if len(sides) > 1 else ""))


def _box_or_points(eid, g, boxes):
    if str(eid) in boxes:
        return boxes[str(eid)]
    pts = g.get(str(eid)) or []
    if pts:
        xs, ys = [float(p[0]) for p in pts], [float(p[1]) for p in pts]
        return (min(xs), min(ys), max(xs), max(ys))
    return None


def _rel_distance(e, g, boxes):
    """Minimum distance from a's range box (text, cell) or path to the polyline b, within value +/- tol, or >= min / <= max."""
    ref = [_xy(v) for v in g.get(str(e["b"])) or []]
    box = _box_or_points(e["a"], g, boxes)
    if len(ref) < 2 or box is None:
        return "UNC", "geometry of a or b not readable"
    d = geom.dist_box_polyline(box, ref)
    parts, ok = [], True
    if e.get("value") is not None:
        tol = float(e.get("tol") or 0.05)
        ok &= abs(d - float(e["value"])) <= tol
        parts.append(f"{float(e['value']):g} +/- {tol:g}")
    if e.get("min") is not None:
        ok &= d >= float(e["min"]) - 1e-9
        parts.append(f">= {float(e['min']):g}")
    if e.get("max") is not None:
        ok &= d <= float(e["max"]) + 1e-9
        parts.append(f"<= {float(e['max']):g}")
    return ("OK" if ok else "BAD"), f"distance {d:.3f} ft; required {' and '.join(parts) or 'none given'}"


def _rel_side(e, g, boxes):
    """a lies left/right of the directed polyline b, or above/below it in the drawing (profile annotation)."""
    ref = [_xy(v) for v in g.get(str(e["b"])) or []]
    box = _box_or_points(e["a"], g, boxes)
    if len(ref) < 2 or box is None:
        return "UNC", "geometry of a or b not readable"
    want = str(e.get("target") or e.get("side") or "").lower()
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    if want in ("above", "below"):
        y = geom.y_at_x(ref, cx)
        if y is None:
            return "UNC", f"b does not span x={cx:.1f}"
        got = "above" if cy > y else "below"
        return ("OK" if got == want else "BAD"), f"a's centre is {got} b (y {cy:.2f} vs {y:.2f}); required {want}"
    if want in ("left", "right"):
        _, signed, _ = geom.nearest_on_polyline((cx, cy), ref)
        got = "left" if signed > 0 else "right"
        return ("OK" if got == want else "BAD"), f"a is {got} of b ({abs(signed):.3f} ft); required {want}"
    return "UNC", "side needs target above|below|left|right"


def _rel_vclear(e, g, boxes):
    """Vertical clearance in a profile: (y_upper - y_lower) / vertical exaggeration at each station x must be >= min (ft)."""
    up, low = [_xy(v) for v in g.get(str(e["a"])) or []], [_xy(v) for v in g.get(str(e["b"])) or []]
    if len(up) < 2 or len(low) < 2:
        return "UNC", "geometry of the soffit (a) or the profile (b) not readable"
    xs = e.get("x") if isinstance(e.get("x"), list) else [e.get("x")]
    xs = [float(x) for x in xs if x is not None]
    if not xs or e.get("min") is None:
        return "UNC", "vclear needs x (station position) and min (the required clearance, ft; from the engineer/HDM, never assumed)"
    ve, need = float(e.get("ve") or 1.0), float(e["min"])
    rows, ok = [], True
    for x in xs:
        ya, yb = geom.y_at_x(up, x), geom.y_at_x(low, x)
        if ya is None or yb is None:
            return "UNC", f"a or b does not span x={x:g}"
        c = (ya - yb) / ve
        ok &= c >= need - 1e-9
        rows.append(f"x={x:g}: {c:.2f} ft")
    return ("OK" if ok else "BAD"), f"vertical clearance (VE {ve:g}) {', '.join(rows)}; required >= {need:g} ft"


def relations(task, ctx):
    rows = [e for e in task.get("expect") or [] if e.get("rel")]
    if not rows:
        return _res("NA", "no geometric relations declared")
    gm = ctx.get("geom") or {}
    boxes = ctx.get("geom_boxes") or {}
    bad, unc, ok = [], [], []
    for e in rows:
        rel = e["rel"]
        if rel in ("offset_curve", "distance", "side", "vclear"):
            fn = {"offset_curve": _rel_offset_curve, "distance": _rel_distance, "side": _rel_side, "vclear": _rel_vclear}[rel]
            try:
                st, text = fn(e, gm, boxes)
            except (KeyError, TypeError, ValueError) as ex:
                st, text = "UNC", f"relation row incomplete: {ex}"
            line = f"{e['id']}: {text}"
            {"OK": ok, "BAD": bad, "UNC": unc}[st].append(line)
            continue
        a, b = gm.get(str(e.get("a"))), gm.get(str(e.get("b"))) if e.get("b") else None
        if not a or (rel != "length" and not b):
            unc.append(f"{e['id']}: geometry of {e.get('a')}/{e.get('b')} not readable")
            continue
        val, unit = _measure(rel, a, b)
        if val is None:
            bad.append(f"{e['id']}: {rel} not measurable ({unit or 'need two lines'})")
            continue
        want, tol = float(e["value"]), float(e.get("tol") or 0.05)
        line = f"{e['id']}: {rel} measured {val:.3f} {unit}, expected {want:g} +/- {tol:g}"
        (ok if abs(val - want) <= tol else bad).append(line)
    if bad:
        return _res("FAIL", f"{len(bad)} of {len(rows)} declared geometric relations not met", "", bad,
                    "Fix the geometry, not the expectation.")
    if unc:
        return _res("UNCERTAIN", f"{len(unc)} relations not measurable", "", unc, "Check the element ids and re-run.")
    return _res("PASS", f"{len(ok)} declared geometric relations measured from saved geometry and met", evidence=ok)
