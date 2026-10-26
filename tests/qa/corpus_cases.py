"""Regression corpus cases (see qa/corpus.py). Each case is a defect the gate must block or legitimate work it must pass.

Operation cases reuse real records captured live from MicroStation (fixtures/oplog_live_2026-10-23.jsonl) and mutate
them; geometry/annotation cases are built from the same element-record shape the gate reads back.
Add a case for every engineer correction (CLAUDE.md QA gate rule 6).
"""
from __future__ import annotations

import copy
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa.corpus import Case  # noqa: E402
from qa.checks import process  # noqa: E402
from tests.qa.test_qa_checks import el, ctx, _txt  # noqa: E402

LIVE = [json.loads(ln) for ln in open(os.path.join(ROOT, "tests", "qa", "fixtures", "oplog_live_2026-10-23.jsonl"), encoding="utf-8")]


def rec(op, n=0):
    return copy.deepcopy([r for r in LIVE if r["op"] == op][n])


def op_case(cid, should_block, why, recs, desc, expect=None, signoff=None, family="operation"):
    ids = set()
    for r in recs:
        ids |= set(r["after"]) | set(r["before"]) | {str(i) for i in r["new_ids"]}
    c = ctx([])
    c["oplog"] = recs
    c["changes"] = {"created": [{"elementId": i} for i in ids], "modified": [], "deleted": {}, "modified_fields": {}}
    if signoff:
        c["signoff"] = signoff
    return Case(cid, should_block, why, {"desc": desc, "expect": expect or []}, c, family)


def geom_case(cid, should_block, why, expect, geom, boxes=None, family="relation"):
    c = ctx([])
    c["geom"], c["geom_boxes"] = geom, boxes or {}
    return Case(cid, should_block, why, {"desc": "", "expect": expect}, c, family)


# reference geometry: a curved baseline (quarter-ish arc as a polyline) and offsets of it
import math  # noqa: E402

ARC = [(100 * math.sin(math.radians(a)), 100 - 100 * math.cos(math.radians(a))) for a in range(0, 61, 5)]


def offset_arc(d):  # parallel to ARC at distance d to the LEFT of travel (left normal of a curve turning left = toward the centre)
    out = []
    for a in range(0, 61, 5):
        r = 100 - d
        out.append((r * math.sin(math.radians(a)), 100 - r * math.cos(math.radians(a))))
    return out


DEFAULT_LVL = {"Default": {"color": "0", "weight": "0", "style": "0", "fromLibrary": "True"}}
GRADE = [(0.0, 10.0), (100.0, 14.0)]  # profile grade line, y = 10 + 0.04 x


def cases() -> list[Case]:
    out: list[Case] = []
    # ---------------------------------------------------------------- operations (recorded edits re-measured)
    mv = rec("MOVE_ELEMENT")
    out.append(op_case("OP-move-ok", False, "move 3 ft east, 2 ft north as asked", [mv], "move the line 3 ft east and 2 ft north"))
    bad = copy.deepcopy(mv)
    bad["after"]["4447"]["verts"] = [[v[0] + 5, v[1] + 2, v[2]] for v in mv["before"]["4447"]["verts"]]
    out.append(op_case("OP-move-wrong-distance", True, "geometry moved 5 ft but the op says 3 ft", [bad], "move the line 3 ft east and 2 ft north"))
    out.append(op_case("OP-move-wrong-direction", True, "request says west, the move went east", [mv], "move the line 3 ft west"))
    out.append(op_case("OP-move-number-unused", True, "request says 12 ft, the op moved 3 ft", [mv], "move the line 12 ft east"))
    out.append(op_case("OP-move-size-chosen-by-agent", True, "request names no distance, the agent chose one", [mv], "move the line east"))
    out.append(op_case("OP-move-size-signed-off", False, "no distance in the request, engineer signed INTENT", [mv], "move the line east",
                       signoff={"items": ["INTENT"], "by": "eng", "at": "now"}))
    cp = rec("COPY_PARALLEL")
    out.append(op_case("OP-offset-ok", False, "offset 12 ft left", [cp], "offset the line 12 ft to the left"))
    out.append(op_case("OP-offset-wrong-side", True, "request says right, offset went left", [cp], "offset the line 12 ft to the right"))
    wrong = copy.deepcopy(cp)
    wrong["params"]["distance"] = -12.0
    out.append(op_case("OP-offset-measured-side-differs", True, "op says -12 (right) but the copy sits 12 ft left", [wrong], "offset the line 12 ft"))
    ex = rec("EXTEND_LINE")
    short = copy.deepcopy(ex)
    short["after"]["4455"]["verts"] = [[4800.0, 700.0, 0.0], [4850.0, 700.0, 0.0]]
    out.append(op_case("OP-extend-wrong-length", True, "extended to 50 ft instead of 55", [short], "extend the line to 55 ft"))
    out.append(op_case("OP-extend-ok", False, "extended to 55 ft", [ex], "extend the line to 55 ft"))
    br = rec("BREAK_LINE")
    out.append(op_case("OP-break-off-line", True, "break point is off the line so the pieces kink", [br], "break the line at 4920"))
    rot = rec("ROTATE_ELEMENT")
    rbad = copy.deepcopy(rot)
    rbad["params"]["angleDeg"] = -90
    out.append(op_case("OP-rotate-wrong-direction", True, "commanded -90 but the geometry turned +90", [rbad], "rotate the line 90 degrees"))
    out.append(op_case("OP-rotate-ok", False, "rotated 90 degrees about its start", [rot], "rotate the line 90 degrees"))
    txt = rec("EDIT_TEXT_ELEMENT")
    tbad = copy.deepcopy(txt)
    tbad["after"]["4460"]["props"]["textHeight"] = "0.003"
    out.append(op_case("OP-edit-text-changes-style", True, "text edit shrank the text height", [tbad], "edit the label"))
    c = op_case("OP-unrecorded-direct-edit", True, "element changed by a direct script, no recorded operation",
                [], "fix the label", expect=[{"id": "E1", "count": ">=1"}])
    c.ctx["changes"]["created"] = [{"elementId": "9"}]
    out.append(c)
    c = op_case("OP-unrecorded-but-relation-declared", False, "direct edit covered by an independent --rel check", [], "offset the line 12 ft",
                expect=[{"id": "R1", "rel": "offset", "a": "8", "b": "9", "value": 12}])
    c.ctx["changes"]["created"] = [{"elementId": "9"}]
    c.ctx["geom"] = {"8": [[0.0, 0.0, 0.0], [40.0, 0.0, 0.0]], "9": [[0.0, 12.0, 0.0], [40.0, 12.0, 0.0]]}
    out.append(c)
    c = op_case("OP-unrecorded-signed-off", False, "direct edit signed off by the engineer", [], "fix the label",
                signoff={"items": ["OP-01"], "by": "eng", "at": "now"})
    c.ctx["changes"]["created"] = [{"elementId": "9"}]
    out.append(c)
    c = op_case("OP-unrecorded-direct-edit-b", True, "element created by a direct script, no operation and no relation", [], "add the line")
    c.ctx["changes"]["created"] = [{"elementId": "9"}]
    out.append(c)

    # ---------------------------------------------------------------- change scope
    line = el(1, typ="LINE", level="Default", color="0", weight="0", box=(0, 0, 10, 10))
    base = ctx([line])
    base["changes"] = {"created": [], "modified": [], "deleted": {}, "modified_fields": {}}
    out.append(Case("CHG-no-change", False, "nothing but the new element changed", {}, copy.deepcopy(base), "change-scope"))
    c = copy.deepcopy(base)
    c["changes"] = {"created": [], "modified": [line], "deleted": {}, "modified_fields": {"1": ["vertices"]}}
    out.append(Case("CHG-vertex-edit-undeclared", True, "an existing line's vertices changed inside an unchanged range box, not declared", {}, c, "change-scope"))
    out.append(Case("CHG-vertex-edit-declared", False, "the engineer asked to edit existing geometry", {"may_modify": True}, copy.deepcopy(c), "change-scope"))
    c = copy.deepcopy(base)
    c["changes"] = {"created": [], "modified": [], "deleted": {"7": {"type": "LINE", "level": "Default"}}, "modified_fields": {}}
    out.append(Case("CHG-silent-delete", True, "an existing element was deleted without being asked", {}, c, "change-scope"))
    c = copy.deepcopy(base)
    c["model_mismatch"] = "task started in a|Design but the active model is a|Default"
    out.append(Case("CHG-different-model", True, "the active model is not the model the snapshot was taken in", {}, c, "change-scope"))

    # ---------------------------------------------------------------- evidence integrity
    c = ctx([line])
    c["state_rows"] = [line]
    import qa.checks.com_spot as cs  # noqa: E402
    c["_com"] = {"1": {"low": (0.0, 0.0), "high": (10.0, 10.0), "level": "Default"}}
    out.append(Case("EVID-bridge-and-com-agree", False, "bridge read-back matches MicroStation COM", {}, c, "evidence",
                    checks=[("COM-01", lambda t, x: _com(cs, t, x))]))
    c = copy.deepcopy(c)
    c["_com"] = {"1": {"low": (0.0, 0.0), "high": (10.0, 40.0), "level": "Default"}}
    out.append(Case("EVID-bridge-disagrees-with-com", True, "bridge reports a 10 ft box, COM says 40 ft", {}, c, "evidence",
                    checks=[("COM-01", lambda t, x: _com(cs, t, x))]))

    # ---------------------------------------------------------------- CAD conventions
    out.append(Case("CAD-wrong-symbology", True, "WZTC element drawn with color 3", {}, ctx([el(1, color="3")]), "convention"))
    out.append(Case("CAD-correct-symbology", False, "WZTC element drawn per the project table", {}, ctx([el(1)]), "convention"))
    out.append(Case("CAD-non-library-level", True, "element on a level that is not in the NYSDOT library", {},
                    ctx([el(1, level="MY_LEVEL", lvlFromLibrary="False")]), "convention"))
    out.append(Case("CAD-outside-extent", True, "element 5000 ft from the work area", {}, ctx([el(1, box=(5000, 5000, 5010, 5001))]), "convention"))
    out.append(Case("CAD-stacked-duplicate", True, "the same line placed twice", {}, ctx([el(1), el(2)]), "convention"))
    out.append(Case("CAD-scale-bar-wrong-scale", True, "1:240 scale bar in a 1:480 model", {},
                    ctx([el(1, typ="CELL", level="NSB_P", cellName="240_1", lvlFromLibrary="True")], model_state={"annotationScaleFactor": "480"}), "convention"))
    out.append(Case("CAD-scale-bar-right-scale", False, "1:480 scale bar in a 1:480 model", {},
                    ctx([el(1, typ="CELL", level="", cellName="480_Vert", childLevels="NSB_P|NSB_P_T")], model_state={"annotationScaleFactor": "480"},
                        level_library={"NSB_P", "NSB_P_T"}), "convention"))

    # ---------------------------------------------------------------- annotation
    notes = [_txt(4081, 0.007, 0.007, (4677, 409, 4797, 416), text="STA 66+00 El. 33.0")]
    good = _txt(900, 0.007, 0.007, (4700, 430, 4760, 437), text="MIN. VC")
    out.append(Case("TXT-clear-and-matching", False, "label matches the sheet style and sits clear", {}, ctx([good], reference_texts=notes, extent=[4600, 400, 4900, 500], level_defs=DEFAULT_LVL), "annotation"))
    small = _txt(901, 0.003, 0.010, (4700, 430, 4760, 433), text="RW")
    out.append(Case("TXT-wrong-style", True, "label 0.003x0.010 beside notes 0.007x0.007", {}, ctx([small], reference_texts=notes, extent=[4600, 400, 4900, 500]), "annotation"))
    bold_notes = [_txt(4082, 0.00521, 0.004, (4677, 409, 4797, 416), text="270", bold=True)]
    out.append(Case("TXT-regular-beside-bold", True, "regular-weight station label beside bold axis text", {},
                    ctx([_txt(904, 0.00521, 0.004, (4700, 430, 4760, 437), text="66+00", bold=False)], reference_texts=bold_notes, extent=[4600, 400, 4900, 500]), "annotation"))
    out.append(Case("TXT-bold-beside-bold", False, "bold station label beside bold axis text", {},
                    ctx([_txt(905, 0.00521, 0.004, (4700, 430, 4760, 437), text="66+00", bold=True)], reference_texts=bold_notes, extent=[4600, 400, 4900, 500], level_defs=DEFAULT_LVL), "annotation"))
    ord_a = _txt(66005, 0.00521, 0.004, (4119, 706, 4171, 766), text="CURVE: CREST")
    out.append(Case("TXT-existing-nodes-overlap", True, "two existing ORD annotation nodes overlap", {},
                    ctx([], model_texts=[ord_a, _txt(66027, 0.00521, 0.004, (4123, 675, 4145, 727), text="VHP")]), "annotation"))
    out.append(Case("TXT-existing-nodes-clear", False, "the same two nodes after moving one clear", {},
                    ctx([], model_texts=[ord_a, _txt(66027, 0.00521, 0.004, (4180, 675, 4202, 727), text="VHP")], level_defs=DEFAULT_LVL), "annotation"))
    over = _txt(902, 0.007, 0.007, (4700, 410, 4760, 415), text="OVER")
    out.append(Case("TXT-overlaps-text", True, "label sits on an existing note", {}, ctx([over], reference_texts=notes, extent=[4600, 400, 4900, 500]), "annotation"))
    ln = el(10, typ="LINE", level="Default", color="0", weight="0", box=(0, 0, 100, 50))
    on = _txt(903, 0.007, 0.007, (48, 22, 72, 30), text="RW")
    c = ctx([ln, on], reference_texts=[_txt(1, 0.007, 0.007, (500, 500, 520, 507))], line_paths={"10": [(0.0, 0.0), (100.0, 50.0)]})
    out.append(Case("TXT-on-line", True, "label sits on a diagonal line (range box would hide it)", {}, c, "annotation"))
    off = _txt(904, 0.007, 0.007, (40, 40, 70, 47), text="RW")
    c = ctx([ln, off], reference_texts=[_txt(1, 0.007, 0.007, (500, 500, 520, 507))], line_paths={"10": [(0.0, 0.0), (100.0, 50.0)]}, level_defs=DEFAULT_LVL)
    out.append(Case("TXT-off-line", False, "label clear of the diagonal line", {}, c, "annotation"))

    # ---------------------------------------------------------------- reference-relative geometry
    out.append(geom_case("REL-curve-offset-ok", False, "curved line is a constant 12 ft from the curved reference",
                         [{"id": "R", "rel": "offset_curve", "a": "1", "b": "2", "value": 12, "tol": 0.05, "side": "left"}],
                         {"1": ARC, "2": offset_arc(12)}))
    out.append(geom_case("REL-curve-offset-diverges", True, "line is 12 ft off at the start and 14 ft off at the end",
                         [{"id": "R", "rel": "offset_curve", "a": "1", "b": "2", "value": 12, "tol": 0.05}],
                         {"1": ARC, "2": [offset_arc(12)[0]] + [p for p in offset_arc(14)[1:]]}))
    out.append(geom_case("REL-straight-translation-of-a-curve", True, "agent translated the curve 12 ft instead of offsetting it",
                         [{"id": "R", "rel": "offset_curve", "a": "1", "b": "2", "value": 12, "tol": 0.05}],
                         {"1": ARC, "2": [(x + 12, y) for x, y in ARC]}))
    out.append(geom_case("REL-wrong-side", True, "offset is on the right, the request says left",
                         [{"id": "R", "rel": "offset_curve", "a": "1", "b": "2", "value": 12, "tol": 0.05, "side": "left"}],
                         {"1": ARC, "2": offset_arc(-12)}))
    out.append(geom_case("REL-wrong-reference", True, "offset measured against the wrong baseline element",
                         [{"id": "R", "rel": "offset_curve", "a": "3", "b": "2", "value": 12, "tol": 0.05}],
                         {"3": [(x, y + 30) for x, y in ARC], "2": offset_arc(12)}))
    box_below = (40.0, 8.0, 60.0, 11.0)   # grade at x=50 is y=12
    box_above = (40.0, 12.5, 60.0, 15.0)
    out.append(geom_case("REL-label-above-grade", False, "label above the profile grade as required",
                         [{"id": "R", "rel": "side", "a": "t", "b": "g", "target": "above"}], {"g": GRADE}, {"t": box_above}))
    out.append(geom_case("REL-label-below-grade", True, "label below the profile grade, required above",
                         [{"id": "R", "rel": "side", "a": "t", "b": "g", "target": "above"}], {"g": GRADE}, {"t": box_below}))
    out.append(geom_case("REL-label-gap-ok", False, "label 0.5 ft off the grade as required",
                         [{"id": "R", "rel": "distance", "a": "t", "b": "g", "value": 0.5, "tol": 0.1}], {"g": GRADE}, {"t": (40.0, 12.9, 60.0, 14.0)}))
    out.append(geom_case("REL-label-too-close", True, "label 0.1 ft off the grade, at least 0.5 required",
                         [{"id": "R", "rel": "distance", "a": "t", "b": "g", "min": 0.5}], {"g": GRADE}, {"t": (40.0, 12.3, 60.0, 14.0)}))
    soffit = [(0.0, 20.0), (100.0, 20.0)]
    out.append(geom_case("REL-clearance-met", False, "16 ft clearance at VE 1 against a 14.5 ft minimum",
                         [{"id": "R", "rel": "vclear", "a": "s", "b": "g", "x": [50], "ve": 1, "min": 5.0}], {"s": soffit, "g": GRADE}))
    out.append(geom_case("REL-clearance-below-minimum", True, "soffit only 8 ft over the grade, 14.5 ft required",
                         [{"id": "R", "rel": "vclear", "a": "s", "b": "g", "x": [50], "ve": 1, "min": 14.5}], {"s": soffit, "g": GRADE}))
    out.append(geom_case("REL-clearance-no-minimum-given", True, "vclear with no required minimum cannot be decided",
                         [{"id": "R", "rel": "vclear", "a": "s", "b": "g", "x": [50], "ve": 1}], {"s": soffit, "g": GRADE}))
    out.append(geom_case("REL-clearance-ve-matters", True, "10 drawing units over the grade at VE 5 is only 2 ft, 14.5 required",
                         [{"id": "R", "rel": "vclear", "a": "s", "b": "g", "x": [50], "ve": 5, "min": 14.5}], {"s": [(0.0, 22.0), (100.0, 22.0)], "g": GRADE}))

    # ---------------------------------------------------------------- coverage
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    c["results"] = {"CAD-LVL-01": {"status": "PASS"}}
    out.append(Case("COV-mandatory-unverified", True, "text task where most mandatory checks never passed", {}, c, "coverage", checks=[("none", lambda t, x: {"status": "PASS"})]))
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    c["results"] = {k: {"status": "PASS"} for k in ("CAD-LVL-01", "CAD-SYM-01", "CAD-LOC-01", "CAD-DUP-01", "CAD-DEG-01", "CAD-TXT-02", "CAD-TXT-03", "CAD-TXT-04", "CAD-TXT-05")}
    out.append(Case("COV-mandatory-all-verified", False, "every mandatory text requirement verified", {}, c, "coverage", checks=[("none", lambda t, x: {"status": "PASS"})]))

    # ---------------------------------------------------------------- process (event log of the work, PROC-01)
    def proc(cid, should_block, why, *evs):
        c = ctx([])
        c["events"] = [{"kind": k if isinstance(k, str) else k[0], "at": i, **({} if isinstance(k, str) else k[1])} for i, k in enumerate(evs)]
        out.append(Case(cid, should_block, why, {}, c, "process", checks=[("PROC-01", process.discipline)]))

    esc = ("escalate", {"reason": "library levels refuse agent changes; Level Display step handed to the engineer"})
    proc("PROC-circles-edit-before-task", True, "2026-10-24 circles run: edits ran before any QA task existed",
         "attempt_fail", "attempt_fail", "task_open")
    proc("PROC-circles-edit-before-reading", True, "circles run: task opened but the Level Display playbook was never read",
         "task_open", "attempt_fail", "attempt_ok")
    proc("PROC-circles-trial-and-error", True, "circles run: five failed level/display edits in a row, no replan",
         "task_open", "read", "attempt_fail", "attempt_fail", "attempt_fail", "attempt_fail", "attempt_fail")
    proc("PROC-replan-without-new-reading", True, "replan issued with nothing read since the failures",
         "task_open", "read", "attempt_fail", "attempt_fail", ("replan", {}))
    proc("PROC-blocked-without-escalation", True, "outcome blocked with no escalation reason for the engineer",
         "task_open", "read", "attempt_fail", ("outcome", {"value": "blocked"}))
    proc("PROC-pass-without-verifying-run", True, "outcome pass claimed with no passing verification",
         "task_open", "read", "attempt_ok", ("outcome", {"value": "pass"}))
    proc("PROC-circles-bounded-handoff", False, "read the playbook, two attempts, escalate the GUI step with a reason",
         "task_open", "search", "read", "attempt_fail", "attempt_fail", esc, ("outcome", {"value": "blocked"}))
    proc("PROC-replan-after-reading", False, "failed twice, read the source, replanned, then succeeded and verified",
         "task_open", "read", "attempt_fail", "attempt_fail", "read", ("replan", {}), "attempt_ok", "verify_pass", ("outcome", {"value": "pass"}))
    lad = ("ladder_open", {"gating": ["R1", "R2", "R3", "R4"]})
    rg = lambda r, st: ("rung", {"rung": r, "status": st})  # noqa: E731
    proc("PROC-ladder-script-before-cheap-checks", True, "circles run: scripted edit while Level Display / View Options rungs had no result",
         "task_open", lad, "read", rg("R1", "negative"), "attempt_fail")
    proc("PROC-ladder-no-rungs-recorded", True, "scripted edit with the diagnostic ladder untouched",
         "task_open", lad, "read", "attempt_ok", "verify_pass", ("outcome", {"value": "pass"}))
    proc("PROC-ladder-cheap-checks-first", False, "read-only identification, GUI rungs handed to the engineer, then a scripted edit",
         "task_open", lad, "read", rg("R1", "negative"), rg("R2", "handed_off"), rg("R3", "handed_off"), rg("R4", "na"), "attempt_ok", "verify_pass")
    proc("PROC-ladder-cause-found-early", False, "Level Display rung positive: the diagnosis ended, later rungs not needed",
         "task_open", lad, "read", rg("R1", "negative"), rg("R2", "positive"), "attempt_ok", "verify_pass")
    proc("PROC-clean-pass", False, "read, one successful edit, verified, pass",
         "task_open", "read", "attempt_ok", "verify_pass", ("outcome", {"value": "pass"}))
    return out


def _com(cs, task, c):
    """COM-01 against the case's canned COM answer instead of a live MicroStation."""
    orig = cs.read_com
    cs.read_com = lambda ids: {k: v for k, v in c["_com"].items() if k in ids}
    try:
        return cs.verify(task, c)
    finally:
        cs.read_com = orig
