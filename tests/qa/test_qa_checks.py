"""Known-bad / known-good fixtures for QA checkers. A checker that passes a bad fixture is broken.

Run: python -m pytest tests/qa -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

import pytest
import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa.checks import cad_rules, task_expect, curb, calc  # noqa: E402
from engcalc import hydraulics  # noqa: E402

RULES = yaml.safe_load(open(os.path.join(ROOT, "qa", "project_rules.yaml"), encoding="utf-8"))


def el(eid, level="TWZCD_P", typ="LineString", color="6", weight="2", z=0.0, box=(0, 0, 10, 1), **kw):
    r = {"elementId": str(eid), "type": typ, "level": level, "color": color, "weight": weight, "lineStyle": "0",
         "lvlColor": "6", "lvlWeight": "2", "lvlStyle": "0", "lvlFromLibrary": "True", "cellName": "",
         "textHeight": "", "elClass": "0", "zLow": str(z), "zHigh": str(z),
         "rangeLowX": str(box[0]), "rangeLowY": str(box[1]), "rangeHighX": str(box[2]), "rangeHighY": str(box[3])}
    r.update(kw)
    return r


def ctx(els, **kw):
    c = {"elements": els, "rules": RULES, "is2d": True, "extent": [-100, -100, 100, 100]}
    c.update(kw)
    return c


# ---------------- CAD rules: good passes, each bad fails
def test_clean_drawing_passes_all():
    c = ctx([el(1), el(2, box=(0, 5, 10, 6))])
    for fn in (cad_rules.new_elements, cad_rules.levels, cad_rules.symbology, cad_rules.elevation,
               cad_rules.location, cad_rules.duplicates, cad_rules.degenerate, cad_rules.cells):
        assert fn({}, c)["status"] == "PASS", fn.__name__


def test_zero_new_elements_is_uncertain_not_pass():
    assert cad_rules.new_elements({}, ctx([]))["status"] == "UNCERTAIN"


def test_wrong_color_on_wztc_level_fails():
    r = cad_rules.symbology({}, ctx([el(1, color="3")]))
    assert r["status"] == "FAIL" and "1:" in r["evidence"][0]


def test_dimension_on_default_must_be_yellow():
    assert cad_rules.symbology({}, ctx([el(1, level="Default", typ="Dimension", color="0", weight="0")]))["status"] == "FAIL"
    assert cad_rules.symbology({}, ctx([el(1, level="Default", typ="Dimension", color="2", weight="0")]))["status"] == "PASS"


def test_non_library_level_fails():
    r = cad_rules.levels({}, ctx([el(1, level="MyLevel", lvlFromLibrary="False")]))
    assert r["status"] == "FAIL"


def test_level_library_snapshot_used_when_flag_blank():
    lib = {"TWZCD_P", "Default"}
    assert cad_rules.levels({}, ctx([el(1, lvlFromLibrary="")], level_library=lib))["status"] == "PASS"
    assert cad_rules.levels({}, ctx([el(1, level="MyLevel", lvlFromLibrary="")], level_library=lib))["status"] == "FAIL"


def test_unreadable_library_flag_is_uncertain():
    assert cad_rules.levels({}, ctx([el(1, level="XYZ", lvlFromLibrary="")]))["status"] == "UNCERTAIN"


def test_nonzero_z_fails():
    assert cad_rules.elevation({}, ctx([el(1, z=12.5)]))["status"] == "FAIL"


def test_stray_element_fails_location():
    assert cad_rules.location({}, ctx([el(1, box=(5000, 5000, 5010, 5001))]))["status"] == "FAIL"


def test_missing_extent_is_uncertain():
    assert cad_rules.location({}, ctx([el(1)], extent=None))["status"] == "UNCERTAIN"


def test_stacked_duplicate_fails():
    assert cad_rules.duplicates({}, ctx([el(1), el(2)]))["status"] == "FAIL"


def test_zero_length_line_fails():
    assert cad_rules.degenerate({}, ctx([el(1, box=(3, 3, 3, 3))]))["status"] == "FAIL"


def test_unnamed_cell_fails():
    assert cad_rules.cells({}, ctx([el(1, typ="CellHeader", cellName="")]))["status"] == "FAIL"


def test_bylevel_tokens_pass_symbology():
    e = el(1, level="DCB_P", color="4294967295", weight="4294967295", lvlColor="3", lvlWeight="2")
    assert cad_rules.symbology({}, ctx([e]))["status"] == "PASS"


def test_bylevel_mismatch_fails():
    e = el(1, level="DCB_P", color="1", weight="2", lvlColor="3", lvlWeight="2")
    assert cad_rules.symbology({}, ctx([e]))["status"] == "FAIL"


def test_text_height_off_scale_fails():
    t = el(1, typ="Text", textHeight="2.5", box=(0, 0, 4, 1))
    assert cad_rules.text_heights({}, ctx([t], text_heights=[3.0, 4.5]))["status"] == "FAIL"
    assert cad_rules.text_heights({}, ctx([t], text_heights=[2.5]))["status"] == "PASS"


def test_text_on_line_segment_fails_txt04():
    """FM-013: label on a diagonal TOW line must FAIL (segment check, not range AABB)."""
    line = el(10, typ="Line", box=(0, 0, 100, 50))  # range AABB huge
    text = el(11, typ="Text", box=(48, 22, 72, 30), text="RW label")  # sits on y=0.5x
    paths = {"10": [(0.0, 0.0), (100.0, 50.0)]}
    r = cad_rules.text_line_clearance({}, ctx([line, text], line_paths=paths))
    assert r["status"] == "FAIL" and "11" in r["evidence"][0]


def test_text_clear_of_line_segment_passes_txt04():
    line = el(10, typ="Line", box=(0, 0, 100, 50))
    text = el(11, typ="Text", box=(40, 40, 70, 48), text="RW label")  # above diagonal
    paths = {"10": [(0.0, 0.0), (100.0, 50.0)]}
    assert cad_rules.text_line_clearance({}, ctx([line, text], line_paths=paths))["status"] == "PASS"


def test_text_line_clearance_uncertain_without_paths():
    line = el(10, typ="Line", box=(0, 0, 100, 50))
    text = el(11, typ="Text", box=(48, 22, 72, 30))
    assert cad_rules.text_line_clearance({}, ctx([line, text]))["status"] == "UNCERTAIN"


# ---------------- task expectations
def test_expect_counts():
    task = {"expect": [{"id": "E1", "desc": "drums", "where": {"level": "TWZCD"}, "count": ">=2"}]}
    assert task_expect.evaluate(task, ctx([el(1)]))["status"] == "FAIL"
    assert task_expect.evaluate(task, ctx([el(1), el(2, box=(0, 3, 9, 4))]))["status"] == "PASS"


def test_no_expectations_is_uncertain():
    assert task_expect.evaluate({}, ctx([el(1)]))["status"] == "UNCERTAIN"


# ---------------- curb identity (Loc10 incident)
def test_location_parsing():
    assert curb.locs_from_name("HWP24QC1_Location_03-04_53-55_SGD.dwg") == {3, 4, 53, 54, 55}
    assert curb.locs_from_name(r"C:\x\T4318AG25712_3-4,53-55.dwg") == {3, 4, 53, 54, 55}
    assert curb.locs_from_name("T4318A_G25712_58-60.dwg") == {58, 59, 60}
    assert curb.locs_from_name("HWP24QC1_Location_10_SGD.dwg") == {10}


def test_loc10_wrong_content_fails_identity(monkeypatch):
    fake = {"status": "PASSED", "xrefs": [{"name": "Survey R", "path": r"C:\repos\x\Screenshots\T4318AG25712_3-4,53-55.dwg"},
                                          {"name": "HWP24QC1_Location_10", "path": "Screenshots/HWP24QC1_Location_10.dwg"}],
            "plan_xrefs": [{"name": "HWP24QC1_Location_10", "path": "Screenshots/HWP24QC1_Location_10.dwg"}]}
    task = {"files": ["Screenshots/HWP24QC1_Location_10_SGD.dwg"]}
    r = curb.identity(task, {"_curb_verify": fake})
    assert r["status"] == "FAIL" and "3, 4, 53" in r["observed"]


def test_matching_sheet_passes_identity():
    fake = {"xrefs": [{"name": "T4318AG25712_20", "path": "Screenshots/T4318AG25712_20.dwg"}],
            "plan_xrefs": [{"name": "HWP24QC1_Location_20", "path": "Screenshots/HWP24QC1_Location_19-20.dwg"}]}
    r = curb.identity({"files": ["Screenshots/HWP24QC1_Location_20_SGD.dwg"]}, {"_curb_verify": fake})
    assert r["status"] == "PASS"


def test_verify_gate_maps_lines_not_feet_failure():
    # the 2026-10-01 bug: early-bound Length -> 0.01 ft per line -> Depressed "0.04" with ramps present
    fake = {"status": "BLOCKED", "fails": ["ramps_present_depressed_zero:ramps=1"],
            "steel_lf": {"Straight": 0.06, "Depressed": 0.04, "Corner": 0.08, "total": 0.18}}
    assert curb.verify_gate({"files": ["a.dwg"]}, {"_curb_verify": fake})["status"] == "FAIL"


# ---------------- engineering calcs
def test_manning_18in_hand_value():
    q = hydraulics.manning_full_pipe_q(18, 0.005, 0.013)["value"]
    assert abs(q - 7.428) < 0.01


def test_rational_and_required_size():
    q = hydraulics.rational_q_cfs(0.9, 5.0, 2.0)["value"]
    assert abs(q - 9.0) < 1e-9
    assert hydraulics.required_diameter_in(q, 0.005, 0.013)["value"] == 21.0


def test_calc_reproduction_catches_wrong_result():
    good = {"calcs": [{"id": "Q1", "fn": "hydraulics.manning_full_pipe_q", "inputs": {"diameter_in": 18, "slope_ft_per_ft": 0.005, "n": 0.013},
                       "result": 7.428, "units": "cfs", "criterion": "hdm/chapt_08.pdf p94"}]}
    assert calc.verify(good, {})["status"] == "PASS"
    bad = json.loads(json.dumps(good)); bad["calcs"][0]["result"] = 9.1
    assert calc.verify(bad, {})["status"] == "FAIL"
    nocite = json.loads(json.dumps(good)); nocite["calcs"][0]["criterion"] = ""
    assert calc.verify(nocite, {})["status"] == "FAIL"


# ---------------- stop hook blocks an open failing task
def test_stop_hook_blocks_failing_task(tmp_path, monkeypatch):
    import importlib
    hook = importlib.import_module("scripts.qa_stop_hook") if False else None  # noqa: F841
    tasks = tmp_path / "outputs" / "qa" / "tasks" / "t1"
    tasks.mkdir(parents=True)
    import time
    (tasks / "task.json").write_text(json.dumps({"id": "t1", "type": "cad.drafting", "status": "in_progress",
                                                 "created": time.time(), "last_status": "FAIL", "last_run": time.time()}))
    (tasks / "latest.json").write_text(json.dumps({"results": [{"id": "CAD-SYM-01", "status": "FAIL", "observed": "2 off", "blocking": True}]}))
    src = open(os.path.join(ROOT, "scripts", "qa_stop_hook.py"), encoding="utf-8").read()
    script = tmp_path / "scripts" / "qa_stop_hook.py"
    script.parent.mkdir()
    script.write_text(src, encoding="utf-8")
    out = subprocess.run([sys.executable, str(script)], input="{}", capture_output=True, text=True).stdout
    d = json.loads(out)
    assert d["decision"] == "block" and "CAD-SYM-01" in d["reason"]
    # PASS task -> no block, only a close reminder (still block so agent closes it)
    (tasks / "task.json").write_text(json.dumps({"id": "t1", "type": "cad.drafting", "status": "closed",
                                                 "created": time.time(), "last_status": "PASS"}))
    out2 = subprocess.run([sys.executable, str(script)], input="{}", capture_output=True, text=True).stdout
    assert out2.strip() == ""


def test_catalogs_parse_and_reference_real_checks():
    import importlib
    crit = yaml.safe_load(open(os.path.join(ROOT, "qa", "criteria.yaml"), encoding="utf-8"))
    fms = yaml.safe_load(open(os.path.join(ROOT, "qa", "failure_modes.yaml"), encoding="utf-8"))
    assert len(crit) >= 10 and len(fms) >= 5
    for r in crit:
        mod, fn = r["check"].split(".")
        assert hasattr(importlib.import_module(f"qa.checks.{mod}"), fn), r["check"]


def test_default_bylevel_resolves_to_level_symbology():
    e = el(1, level="Default", typ="LINE", color="-1", weight="-1", lvlColor="0", lvlWeight="0")
    assert cad_rules.symbology({}, ctx([e]))["status"] == "PASS"
    e2 = el(2, level="Default", typ="LINE", color="-1", weight="-1", lvlColor="3", lvlWeight="0")
    assert cad_rules.symbology({}, ctx([e2]))["status"] == "FAIL"


def test_civil_helper_levels_ignored():
    pts = [el(1, level="_CIVIL_CONSTRUCTION", typ="LINE", box=(5, 5, 5, 5)),
           el(2, level="_CIVIL_CONSTRUCTION", typ="LINE", box=(5, 5, 5, 5))]
    c = ctx(pts)
    assert cad_rules.duplicates({}, c)["status"] == "PASS"
    assert cad_rules.degenerate({}, c)["status"] == "PASS"


# ---------------- text style vs sheet (FM-012, incident 2026-10-23: RW labels 0.003x0.010 beside 0.007x0.007 notes)
def _txt(eid, h, w, box, font="Arial", text="MIN. VC = 13.36'"):
    return el(eid, level="Default", typ="TEXT", color="-1", weight="-1", box=box, textHeight=str(h),
              textWidth=str(w), textFont=font, textRotDeg="0", text=text)


SHEET_NOTES = [_txt(4081, 0.007, 0.007, (4677, 409, 4797, 416), text="STA 66+00 El. 33.0")]


def test_text_wrong_size_and_aspect_fails():
    c = ctx([_txt(900, 0.003, 0.010, (4700, 250, 4760, 253))], reference_texts=SHEET_NOTES)
    assert cad_rules.text_style({}, c)["status"] == "FAIL"


def test_text_matching_sheet_style_passes():
    c = ctx([_txt(900, 0.007, 0.007, (4700, 300, 4760, 307))], reference_texts=SHEET_NOTES)
    assert cad_rules.text_style({}, c)["status"] == "PASS"


def test_text_wrong_font_fails():
    c = ctx([_txt(900, 0.007, 0.007, (4700, 300, 4760, 307), font="Simplex")], reference_texts=SHEET_NOTES)
    assert cad_rules.text_style({}, c)["status"] == "FAIL"


def test_text_unreadable_style_is_uncertain_not_pass():
    c = ctx([_txt(900, "", "", (4700, 300, 4760, 307), font="")], reference_texts=SHEET_NOTES)
    assert cad_rules.text_style({}, c)["status"] == "UNCERTAIN"


def test_text_without_any_style_authority_is_uncertain():
    c = ctx([_txt(900, 0.007, 0.007, (4700, 300, 4760, 307))], reference_texts=[])
    assert cad_rules.text_style({}, c)["status"] == "UNCERTAIN"


def test_text_overlapping_sheet_note_fails():
    c = ctx([_txt(900, 0.007, 0.007, (4690, 410, 4750, 415))], reference_texts=SHEET_NOTES)
    assert cad_rules.text_overlap({}, c)["status"] == "FAIL"
    c = ctx([_txt(900, 0.007, 0.007, (4690, 440, 4750, 447))], reference_texts=SHEET_NOTES)
    assert cad_rules.text_overlap({}, c)["status"] == "PASS"


# ---------------- cell headers have no level of their own: judge the cell's content levels (live 2026-10-23)
def test_cell_header_with_library_child_levels_passes_level_and_symbology():
    cell = el(7, level="", typ="CELL", color="0", weight="0", cellName="480_Vert", childLevels="NSB_P|NSB_P_T")
    c = ctx([cell], level_library={"NSB_P", "NSB_P_T"})
    assert cad_rules.levels({}, c)["status"] == "PASS"
    assert cad_rules.symbology({}, c)["status"] == "PASS"
    assert cad_rules.cells({}, c)["status"] == "PASS"


def test_cell_with_non_library_child_level_fails():
    cell = el(7, level="", typ="CELL", color="0", weight="0", cellName="X", childLevels="NSB_P|MADE_UP")
    assert cad_rules.levels({}, ctx([cell], level_library={"NSB_P"}))["status"] == "FAIL"


# ---------------- tutorial-derived rules
def _bar(cell, **kw):
    return el(8, level="", typ="CELL", color="0", weight="0", cellName=cell, childLevels="NSB_P|NSB_P_T", **kw)


def test_scale_bar_must_match_annotation_scale():
    assert cad_rules.scale_bar_scale({}, ctx([_bar("480_Vert")], model_state={"annotationScaleFactor": "480"}))["status"] == "PASS"
    assert cad_rules.scale_bar_scale({}, ctx([_bar("240_1")], model_state={"annotationScaleFactor": "480"}))["status"] == "FAIL"
    assert cad_rules.scale_bar_scale({}, ctx([_bar("480_Vert")], model_state={}))["status"] == "UNCERTAIN"
    assert cad_rules.scale_bar_scale({}, ctx([el(1)], model_state={}))["status"] == "NA"


def test_annotation_in_sheet_model_warns():
    t = el(5, level="Default", typ="TEXT", color="0", weight="0")
    assert cad_rules.annotation_model({}, ctx([t], model_state={"modelType": "1", "modelName": "X [Sheet]"}))["status"] == "WARN"
    assert cad_rules.annotation_model({}, ctx([t], model_state={"modelType": "3"}))["status"] == "NA"


# ---------------- quadrant bearings (as-built layout playbook)
def test_bearing_conversion():
    from qa import geom
    assert abs(geom.bearing_to_azimuth("N29°54'00\"E") - 29.9) < 1e-9
    assert abs(geom.bearing_to_azimuth("S 57 08 20 E") - (180 - 57.138889)) < 1e-4
    assert abs(geom.bearing_to_azimuth("N01:33:00:W") - (360 - 1.55)) < 1e-9
    assert abs(geom.azimuth_to_math_deg(90.0)) < 1e-9 and abs(geom.azimuth_to_math_deg(0.0) - 90.0) < 1e-9
    for bad in ("X10 10 10 E", "N95 00 00 E", "29.9"):
        try:
            geom.bearing_to_azimuth(bad)
        except ValueError:
            continue
        raise AssertionError(bad)
