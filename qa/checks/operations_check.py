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


def _desc_numbers(desc: str) -> list[float]:
    return [float(x) for x in _NUM.findall(desc or "")] + [float(x) for x in _DEG.findall(desc or "")]


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
    unrecorded = sorted(changed - recorded)
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
    if unrecorded and not contract and not other:
        has_expect = bool(task.get("expect"))
        status = "WARN" if has_expect else "UNCERTAIN"
        return _res(status, f"{len(unrecorded)} changed elements were not made by a recorded operation, so there is no independent "
                            f"check that the edit is right" + ("" if not has_expect else " (only the task's own expectations apply)"),
                    "an edit made through the recorded tools, or an independent geometry expectation (qa_gate expect --rel ...)",
                    [f"element {i}" for i in unrecorded],
                    "Make the edit with the bridge tools (move_element, copy_parallel, ...), or declare a measurable relation: "
                    "`qa_gate expect --rel offset --a ID --b ID --value 12`.")
    if other and not contract:
        return _res("WARN", f"{len(other)} operations have no geometric contract (composite builders): not verified here",
                    "a contract per operation", sorted({o["op"] for o in other}))
    if not contract:
        return _res("NA", "no edit operations recorded")
    extra = []
    if unrecorded:
        extra.append(f"{len(unrecorded)} other changed elements were not made by a recorded op: {', '.join(unrecorded[:8])}")
    if other:
        extra.append("operations without a contract: " + ", ".join(sorted({o['op'] for o in other})))
    if extra:
        return _res("WARN", f"{passed} of {len(contract)} recorded edit operations verified from before/after geometry", "", ev + extra)
    return _res("PASS", f"{passed} of {len(contract)} recorded edit operations verified from before/after geometry", evidence=ev)


# ---------------------------------------------------------------- declared relations (REQ-GEOM-01)
def _line_ends(vs):
    return (vs[0], vs[-1]) if vs and len(vs) >= 2 else None


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


def relations(task, ctx):
    rows = [e for e in task.get("expect") or [] if e.get("rel")]
    if not rows:
        return _res("NA", "no geometric relations declared")
    gm = ctx.get("geom") or {}
    bad, unc, ok = [], [], []
    for e in rows:
        a, b = gm.get(str(e.get("a"))), gm.get(str(e.get("b"))) if e.get("b") else None
        if not a or (e["rel"] != "length" and not b):
            unc.append(f"{e['id']}: geometry of {e.get('a')}/{e.get('b')} not readable")
            continue
        val, unit = _measure(e["rel"], a, b)
        if val is None:
            bad.append(f"{e['id']}: {e['rel']} not measurable ({unit or 'need two lines'})")
            continue
        want, tol = float(e["value"]), float(e.get("tol") or 0.05)
        line = f"{e['id']}: {e['rel']} measured {val:.3f} {unit}, expected {want:g} +/- {tol:g}"
        (ok if abs(val - want) <= tol else bad).append(line)
    if bad:
        return _res("FAIL", f"{len(bad)} of {len(rows)} declared geometric relations not met", "", bad,
                    "Fix the geometry, not the expectation.")
    if unc:
        return _res("UNCERTAIN", f"{len(unc)} relations not measurable", "", unc, "Check the element ids and re-run.")
    return _res("PASS", f"{len(ok)} declared geometric relations measured from saved geometry and met", evidence=ok)
