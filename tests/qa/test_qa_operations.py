"""Operation-contract regression corpus.

fixtures/oplog_live_2026-10-23.jsonl = 17 real records captured live from MicroStation through the bridge
(op, exact params, geometry before/after read back). Good records must PASS; the deliberate off-line BREAK_LINE
must FAIL; and every good record is MUTATED (wrong distance, wrong side, wrong angle, ...) and must then FAIL:
false passes are measured, not assumed to be zero.
"""
from __future__ import annotations

import copy
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa import operations  # noqa: E402
from qa.checks import operations_check as oc  # noqa: E402
from tests.qa.test_qa_checks import ctx  # noqa: E402

CORPUS = [json.loads(ln) for ln in open(os.path.join(ROOT, "tests", "qa", "fixtures", "oplog_live_2026-10-23.jsonl"), encoding="utf-8")]


def _by(op):
    return [r for r in CORPUS if r["op"] == op]


def test_corpus_good_records_pass_and_bad_break_fails():
    for r in CORPUS:
        st = operations.validate(r)["status"]
        want = "FAIL" if (r["op"] == "BREAK_LINE") else "PASS"
        assert st == want, (r["op"], r["params"], operations.validate(r))


def _num_mutations(rec):
    """Yield (label, mutated copy) for each numeric parameter that defines the operation's geometry."""
    keys = {"MOVE_ELEMENT": ["deltaX", "deltaY"], "COPY_ELEMENT": ["deltaX", "deltaY"], "COPY_PARALLEL": ["distance"],
            "ROTATE_ELEMENT": ["angleDeg", "originX"], "SCALE_ELEMENT": ["scaleFactor"], "MIRROR_ELEMENT": ["y1", "y2"],
            "ARRAY_ELEMENT": ["spacingY", "count"], "EXTEND_LINE": ["newLength"], "PLACE_CIRCLE": ["radius"],
            "PLACE_BLOCK": ["x2"], "CHANGE_ELEMENT_SYMBOLOGY": ["color", "weight"]}.get(rec["op"], [])
    for k in keys:
        if k in rec["params"]:
            m = copy.deepcopy(rec)
            m["params"][k] = float(m["params"][k]) + 1.0
            yield f"{rec['op']}.{k}+1", m
    if rec["op"] == "COPY_PARALLEL":  # wrong side
        m = copy.deepcopy(rec)
        m["params"]["distance"] = -float(m["params"]["distance"])
        yield "COPY_PARALLEL wrong side", m
    if rec["op"] == "ROTATE_ELEMENT":  # opposite direction
        m = copy.deepcopy(rec)
        m["params"]["angleDeg"] = -float(m["params"]["angleDeg"])
        yield "ROTATE_ELEMENT wrong direction", m


def test_mutated_records_never_false_pass():
    tried, false_pass = 0, []
    for r in CORPUS:
        if operations.validate(r)["status"] != "PASS":
            continue
        for label, m in _num_mutations(r):
            tried += 1
            if operations.validate(m)["status"] == "PASS":
                false_pass.append(label)
    assert tried >= 15, tried
    assert not false_pass, f"false passes: {false_pass}"


def test_geometry_mutations_fail():
    for r in CORPUS:
        if operations.validate(r)["status"] != "PASS" or r["op"] not in ("MOVE_ELEMENT", "COPY_PARALLEL", "EXTEND_LINE", "COPY_ELEMENT"):
            continue
        m = copy.deepcopy(r)
        for side in m["after"].values():  # the CAD result differs from what was commanded
            if side.get("verts"):
                side["verts"] = [[v[0] + 0.5, v[1], v[2]] for v in side["verts"]]
        assert operations.validate(m)["status"] == "FAIL", r["op"]


def test_text_edit_that_changes_style_fails():
    r = copy.deepcopy(_by("EDIT_TEXT_ELEMENT")[0])
    eid = r["params"]["elementId"]
    r["after"][eid]["props"]["textHeight"] = "0.003"
    assert operations.validate(r)["status"] == "FAIL"


def test_missing_readback_is_uncertain_never_pass():
    r = copy.deepcopy(_by("MOVE_ELEMENT")[0])
    r["after"] = {}
    assert operations.validate(r)["status"] == "UNCERTAIN"
    assert operations.validate({"op": "FENCE_MOVE_CONTENTS", "params": {}})["status"] == "UNCERTAIN"


# ---------------------------------------------------------------- the OP-01 gate row
def _t(desc="", expect=None):
    return {"desc": desc, "expect": expect or []}


def _ctx(recs, changed=()):
    c = ctx([])
    c["oplog"] = recs
    c["changes"] = {"created": [{"elementId": i} for i in changed], "modified": [], "deleted": {}, "modified_fields": {}}
    return c


def test_op01_passes_when_every_edit_measures_right():
    recs = _by("MOVE_ELEMENT") + _by("COPY_PARALLEL")
    r = oc.operations(_t("move line 3 ft and offset 12 ft"), _ctx(recs))
    assert r["status"] == "PASS", r


def test_op01_fails_on_a_wrong_edit():
    bad = copy.deepcopy(_by("COPY_PARALLEL")[0])
    bad["params"]["distance"] = -12.0
    assert oc.operations(_t(), _ctx([bad]))["status"] == "FAIL"


def test_op01_request_number_not_used_by_any_operation_is_uncertain():
    r = oc.operations(_t("offset the line 21 ft"), _ctx(_by("COPY_PARALLEL")))
    assert r["status"] == "UNCERTAIN" and "21" in r["observed"]


def test_op01_request_direction_contradicting_the_operation_is_uncertain():
    r = oc.operations(_t("offset the line 12 ft to the right"), _ctx(_by("COPY_PARALLEL")))
    assert r["status"] == "UNCERTAIN"
    r = oc.operations(_t("move it 3 ft west"), _ctx(_by("MOVE_ELEMENT")))
    assert r["status"] == "UNCERTAIN"


def test_op01_change_without_recorded_operation_needs_independent_expectation():
    assert oc.operations(_t(), _ctx([], changed=["9"]))["status"] == "UNCERTAIN"
    assert oc.operations(_t(expect=[{"id": "E1", "count": ">=1"}]), _ctx([], changed=["9"]))["status"] == "WARN"


def test_op01_composite_ops_without_a_contract_are_listed_not_passed():
    other = [{"op": "PLACE_SIGN", "params": {}, "contract": False, "new_ids": ["5"], "before": {}, "after": {}}]
    assert oc.operations(_t(), _ctx(other, changed=["5"]))["status"] == "WARN"


def test_op01_not_applicable_without_a_log_position():
    c = ctx([])
    assert oc.operations(_t(), c)["status"] == "NA"


# ---------------------------------------------------------------- declared relations (REQ-GEOM-01)
def _rel(rel, value, a, b=None, tol=0.05):
    return {"id": "G1", "rel": rel, "a": "1", "b": "2" if b is not None else None, "value": value, "tol": tol}


def test_relation_offset_measured_from_geometry():
    geom = {"1": [[0, 0, 0], [40, 0, 0]], "2": [[0, 12, 0], [40, 12, 0]]}
    c = ctx([])
    c["geom"] = geom
    assert oc.relations({"expect": [_rel("offset", 12, 1, 2)]}, c)["status"] == "PASS"
    assert oc.relations({"expect": [_rel("offset", 21, 1, 2)]}, c)["status"] == "FAIL"
    c["geom"] = {"1": geom["1"], "2": [[0, 12, 0], [40, 15, 0]]}  # not parallel -> not an offset
    assert oc.relations({"expect": [_rel("offset", 12, 1, 2)]}, c)["status"] == "FAIL"


def test_relation_length_and_missing_geometry():
    c = ctx([])
    c["geom"] = {"1": [[0, 0, 0], [30, 40, 0]]}
    assert oc.relations({"expect": [_rel("length", 50, 1)]}, c)["status"] == "PASS"
    c["geom"] = {}
    assert oc.relations({"expect": [_rel("length", 50, 1)]}, c)["status"] == "UNCERTAIN"
    assert oc.relations({"expect": []}, c)["status"] == "NA"


def test_ord_signed_offset_flips_the_copy_parallel_sign():
    cp = _by("COPY_PARALLEL")  # distance +12 = LEFT
    assert oc.operations(_t("offset the line 12 ft, ORD offset -12"), _ctx(cp))["status"] == "PASS"       # -12 ORD = left
    r = oc.operations(_t("offset the line 12 ft, ORD offset +12"), _ctx(cp))                              # +12 ORD = right
    assert r["status"] == "UNCERTAIN" and "ORD offset 12" in r["observed"]
