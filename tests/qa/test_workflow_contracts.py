"""Workflow contract tests: the agent must ask (WAITING_FOR_INPUT) instead of inventing designer inputs, and the mutation boundary must refuse
substitute operations. These run against the real contract file and the real guard, not a mock of them.

Scenarios (profile-design): no inputs, no alignment, no units, bad tables, units/stationing mismatch, READY, synthetic-test limits,
terrain-drape / pass-through substitutes refused, disabled creation op refused, verification that detects a wrong profile.
"""
from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

yaml = pytest.importorskip("yaml")
import workflow_contract as wc  # noqa: E402

WF = "profile-design"


@pytest.fixture(autouse=True)
def _state(tmp_path, monkeypatch):
    monkeypatch.setattr(wc, "STATE", tmp_path / "workflow")


def rows(station0=1000.0):
    return [{"station": station0, "elevation": 100.0, "curve_length": 0},
            {"station": station0 + 250, "elevation": 106.0, "curve_length": 100},
            {"station": station0 + 500, "elevation": 102.0, "curve_length": 80},
            {"station": station0 + 755, "elevation": 108.0, "curve_length": 0}]


def good(**kw):
    d = {"alignment": "ALG", "units": "m", "pvis": rows()}
    d.update(kw)
    return d


def reader_for(start=594.36, length=6995.15, found=True):
    def r(cmd, args):
        if cmd == "get_alignment":
            if not found:
                raise RuntimeError("cif_read get_alignment failed: InvalidOperationException: alignment not found (name='X' idx=-1)")
            return [{"key": "startStation", "value": str(start)}, {"key": "length", "value": str(length)}]
        raise AssertionError(cmd)
    return r


# ---------------------------------------------------------------- asking, not inventing
def test_no_inputs_waits_and_lists_everything():
    out = wc.preflight(WF, {})
    assert out["status"] == wc.WAITING
    assert "alignment" in out["missing_inputs"]
    assert any("units" in m for m in out["missing_inputs"])
    assert any("pvis" in m for m in out["missing_inputs"])
    assert "invent_pvi_elevations" in out["prohibited_actions"]


def test_shipped_template_is_waiting_not_ready():
    """Reset-the-conversation regression: the template as shipped (nothing filled in) must never preflight READY."""
    tpl = wc.load_inputs(wc.INPUTS / "profile-design.template.yaml")
    assert wc.preflight(WF, tpl)["status"] == wc.WAITING


def test_pvis_without_alignment_asks_for_alignment_only():
    out = wc.preflight(WF, {"units": "m", "pvis": rows()})
    assert out["status"] == wc.WAITING and out["missing_inputs"] == ["alignment"]


def test_alignment_without_pvis_asks_for_pvis():
    out = wc.preflight(WF, {"alignment": "ALG", "units": "m"})
    assert out["status"] == wc.WAITING and any("pvis" in m for m in out["missing_inputs"])


def test_units_are_never_defaulted():
    out = wc.preflight(WF, {"alignment": "ALG", "pvis": rows()})
    assert out["status"] == wc.WAITING and any("units" in m for m in out["missing_inputs"])
    assert wc.preflight(WF, {"alignment": "ALG", "units": "yards", "pvis": rows()})["status"] == wc.BLOCKED


def test_partial_row_names_the_missing_field():
    r = rows()
    r[1]["elevation"] = None
    out = wc.preflight(WF, good(pvis=r))
    assert out["status"] == wc.WAITING and "pvis[2].elevation" in out["missing_inputs"]


# ---------------------------------------------------------------- invalid tables
@pytest.mark.parametrize("mutate,needle", [
    (lambda r: r[2].update(station=r[1]["station"]), "strictly increasing"),
    (lambda r: r[0].update(curve_length=10), "0 on the first and last"),
    (lambda r: r[1].update(curve_length=600), "does not fit"),
    (lambda r: r[1].update(curve_length=-5), "negative"),
])
def test_bad_tables_are_blocked(mutate, needle):
    r = rows()
    mutate(r)
    out = wc.preflight(WF, good(pvis=r))
    assert out["status"] == wc.BLOCKED and any(needle in p for p in out["problems"]), out


def test_overlapping_curves_blocked():
    r = [{"station": 0, "elevation": 0, "curve_length": 0}, {"station": 100, "elevation": 5, "curve_length": 80},
         {"station": 150, "elevation": 2, "curve_length": 80}, {"station": 400, "elevation": 9, "curve_length": 0}]
    out = wc.preflight(WF, good(pvis=r))
    assert out["status"] == wc.BLOCKED and any("overlap" in p for p in out["problems"]), out


# ---------------------------------------------------------------- READY, live checks, units
def test_complete_inputs_are_ready_with_live_check_and_arm_a_token():
    out = wc.preflight_and_arm(WF, good(alignment="ALG"), reader_for(start=900, length=2000))
    assert out["status"] == wc.READY and out["alignment_stationing_m"][0] == "900"
    assert wc._ready_token(WF)["pvi_args_hash"]


def test_offline_is_unverified_not_ready():
    assert wc.preflight(WF, good())["status"] == wc.UNVERIFIED


def test_stations_outside_alignment_blocked_with_unit_hint():
    out = wc.preflight(WF, good(), reader_for(start=5000, length=800))
    assert out["status"] == wc.BLOCKED and "units" in out["problems"][0]


def test_alignment_not_in_model_blocked():
    out = wc.preflight(WF, good(), reader_for(found=False))
    assert out["status"] == wc.BLOCKED and "not in the open model" in out["problems"][0]


def test_feet_are_converted_to_metres():
    pv = wc.normalised_pvis({"units": "ft", "pvis": [{"station": 1000, "elevation": 100, "curve_length": 0}]})
    assert pv[0]["station"] == pytest.approx(1000 * 1200 / 3937)


# ---------------------------------------------------------------- synthetic tests only on throwaway objects
def test_synthetic_only_on_test_alignments():
    assert wc.preflight(WF, good(synthetic_test=True))["status"] == wc.BLOCKED
    assert wc.preflight(WF, good(alignment="WZTC_TEST_ALG", synthetic_test=True))["status"] == wc.UNVERIFIED


# ---------------------------------------------------------------- mutation boundary
@pytest.mark.parametrize("cmd", ["create_profile", "create_profile_vpi", "profile_from_terrain"])
def test_substitute_profile_ops_refused_on_engineer_alignments(cmd):
    with pytest.raises(wc.WorkflowRefused) as e:
        wc.guard_cif_exec(cmd, {"name": "727.21.101 1960's CL1"})
    assert "PVI" in str(e.value) or "design profile" in str(e.value)


def test_substitute_ops_allowed_on_throwaway_test_objects():
    wc.guard_cif_exec("create_profile", {"name": "WZTC_TEST_ALG"})


def test_create_from_pvis_refused_while_disabled_by_contract():
    with pytest.raises(wc.WorkflowRefused) as e:
        wc.guard_cif_exec("create_profile_from_pvis", {"name": "ALG", "pvis": "1,2,0;3,4,0"})
    assert "disabled by contract" in str(e.value)


def _enabled(monkeypatch):
    real = wc.load_contract

    def lc(wf):
        c = real(wf)
        c["guarded_ops"]["create_profile_from_pvis"]["enabled"] = True
        return c
    monkeypatch.setattr(wc, "load_contract", lc)


def test_enabled_op_needs_a_ready_token(monkeypatch):
    _enabled(monkeypatch)
    with pytest.raises(wc.WorkflowRefused) as e:
        wc.guard_cif_exec("create_profile_from_pvis", {"name": "ALG", "pvis": "1,2,0;3,4,0"})
    assert "not READY" in str(e.value)


def test_enabled_op_refuses_numbers_that_were_not_preflighted(monkeypatch):
    _enabled(monkeypatch)
    inp = good()
    out = wc.preflight_and_arm(WF, inp, reader_for(start=900, length=2000))
    assert out["status"] == wc.READY
    wc.guard_cif_exec("create_profile_from_pvis", {"name": "ALG", "pvis": out["pvis_arg_m"]})          # same numbers: allowed
    with pytest.raises(wc.WorkflowRefused) as e:
        wc.guard_cif_exec("create_profile_from_pvis", {"name": "ALG", "pvis": out["pvis_arg_m"].replace("106.0", "107.0")})
    assert "differ" in str(e.value)


def test_cif_exec_calls_the_guard():
    """The guard sits in the shared mutation function, not only in an MCP tool."""
    src = open(os.path.join(ROOT, "mcp-server", "wztc_ops.py"), encoding="utf-8").read()
    i = src.index("def cif_exec(")
    assert "workflow_contract.guard_cif_exec(cmd, args)" in src[i:i + 900]


def test_cif_rows_reader_takes_the_dict_the_contract_passes():
    """workflow_contract calls reader(cmd, {args}); a keyword-only reader made every live MCP preflight/verify UNVERIFIED."""
    src = open(os.path.join(ROOT, "mcp-server", "wztc_ops.py"), encoding="utf-8").read()
    assert "def _cif_rows(cmd: str, args: dict | None = None, **kw)" in src


# ---------------------------------------------------------------- maths and verification
def test_expected_profile_geometry():
    pv = wc.normalised_pvis({"units": "m", "pvis": [{"station": 0, "elevation": 100, "curve_length": 0},
                                                      {"station": 100, "elevation": 102, "curve_length": 40},
                                                      {"station": 200, "elevation": 100, "curve_length": 0}]})
    assert wc.expected_y(pv, 0) == pytest.approx(100.0)
    assert wc.expected_y(pv, 50) == pytest.approx(101.0)                       # tangent, 2 % grade
    assert wc.expected_y(pv, 100) == pytest.approx(102.0 + (-0.02 - 0.02) * 40 / 8)   # curve midpoint: PVI + (g2-g1)*L/8
    assert wc.expected_y(pv, 200) == pytest.approx(100.0)
    assert wc.expected_y(pv, 201) is None


def _profile_reader(inputs, offset=0.0, defined=True):
    pv = wc.normalised_pvis(inputs)

    def r(cmd, args):
        if cmd == "sample_profile":
            x = args["x0"]
            y = wc.expected_y(pv, x)
            return [{"x": str(x), "y": str(y + offset), "defined": "Y" if defined else "N"}]
        raise AssertionError(cmd)
    return r


def test_verify_passes_a_correct_profile_and_fails_a_wrong_one():
    inp = good()
    assert wc.verify_profile(inp, _profile_reader(inp), 3)["status"] == "PASS"
    bad = wc.verify_profile(inp, _profile_reader(inp, offset=0.5), 3)
    assert bad["status"] == "FAIL" and bad["max_dev_m"] == pytest.approx(0.5)
    assert wc.verify_profile(inp, _profile_reader(inp, defined=False), 3)["status"] == "FAIL"


def test_find_profile_by_station_range():
    inp = good()
    pv = wc.normalised_pvis(inp)

    def r(cmd, args):
        assert cmd == "list_profiles"
        return [{"alignment": "OTHER", "x0": "0", "x1": "100"}, {"alignment": "ALG", "x0": "0", "x1": "100"},
                {"alignment": "ALG", "x0": str(pv[0]["station"]), "x1": str(pv[-1]["station"])}]
    assert wc.find_profile_idx(inp, r) == 1


# ---------------------------------------------------------------- contract consistency
def test_contract_steps_exist_in_the_playbook():
    import re
    c = wc.load_contract(WF)
    pb = open(os.path.join(ROOT, "Data", "tutorial-playbooks", "profile-cutting-workflow.md"), encoding="utf-8").read()
    ids = set(re.findall(r"^\s+- id: (N\d+)", pb, re.M))
    for s in c["manual_steps"]:
        assert s["playbook"] in ids, f"contract step {s['id']} cites playbook step {s['playbook']} that does not exist"
    assert set(c["prohibited_ops"]) >= {"create_profile", "create_profile_vpi", "profile_from_terrain"}


# ---------------------------------------------------------------- full verification: each new check can fail
def _full_reader(inputs, offset_after=None, slope_err=0.0, extra_same_range=0, active=True, drop_original=False):
    """Fake OpenRoads for verify_profile_full. offset_after = station after which elevations jump by 0.05 m."""
    pv = wc.normalised_pvis(inputs)
    s0, s1 = pv[0]["station"], pv[-1]["station"]

    def r(cmd, args):
        if cmd == "sample_profile":
            rows = []
            for x in (args["x"] if "x" in args else [args["x0"]]):
                y = wc.expected_y(pv, x)
                if y is None:
                    rows.append({"x": str(x), "y": "", "slope": "", "defined": "N"})
                    continue
                if offset_after is not None and x > offset_after:
                    y += 0.05
                rows.append({"x": str(x), "y": str(y), "slope": str(wc.expected_slope(pv, x) + slope_err), "defined": "Y"})
            return rows
        if cmd == "list_profiles":
            out = [{"alignment": "ALG", "x0": "0", "x1": "100"}] if not drop_original else []
            out += [{"alignment": "ALG", "x0": str(s0), "x1": str(s1)}] * (1 + extra_same_range)
            return out
        if cmd == "active_profile":
            return [{"key": "activeIsNull", "value": "N" if active else "Y"}, {"key": "activeRange", "value": f"{s0}..{s1}" if active else ""}]
        raise AssertionError(cmd)
    return r


def _full(inp, **kw):
    return wc.verify_profile_full(inp, _full_reader(inp, **kw), 1, baseline=[[0.0, 100.0]])


def test_full_verify_passes_a_correct_profile():
    inp = good()
    out = _full(inp)
    assert out["status"] == "PASS", out["checks"]
    assert set(out["checks"]) >= {"elevation_samples", "dense_elevation", "grade", "endpoints", "continuity", "single_profile", "active_profile", "originals_unchanged"}


def test_full_verify_catches_each_defect():
    inp = good()
    pv = wc.normalised_pvis(inp)
    mid = (pv[0]["station"] + pv[-1]["station"]) / 2
    assert _full(inp, slope_err=0.001)["checks"]["grade"]["status"] == "FAIL"
    j = _full(inp, offset_after=mid)
    assert j["status"] == "FAIL" and j["checks"]["dense_elevation"]["status"] == "FAIL"
    assert _full(inp, extra_same_range=1)["checks"]["single_profile"]["status"] == "FAIL"
    assert _full(inp, active=False)["checks"]["active_profile"]["status"] == "FAIL"
    assert _full(inp, drop_original=True)["checks"]["originals_unchanged"]["status"] == "FAIL"


def test_full_verify_flags_a_jump_at_a_break_and_wrong_endpoint():
    inp = good()
    pv = wc.normalised_pvis(inp)
    brk = pv[1]["station"]                         # PVI of the first curve
    out = _full(inp, offset_after=brk)
    assert out["checks"]["continuity"]["status"] == "FAIL" or out["checks"]["endpoints"]["status"] == "FAIL"
    assert out["checks"]["endpoints"]["status"] == "FAIL"        # last elevation is 0.05 m off


def test_full_verify_without_baseline_is_unverified_not_pass():
    inp = good()
    out = wc.verify_profile_full(inp, _full_reader(inp), 1, baseline=None)
    assert out["checks"]["originals_unchanged"]["status"] == "UNVERIFIED" and out["status"] == "UNVERIFIED"


def test_expected_slope_matches_numeric_derivative_of_expected_y():
    pv = wc.normalised_pvis(good())
    for x in (pv[0]["station"] + 1.0, pv[1]["station"] - 3.0, pv[1]["station"] + 2.0, pv[-1]["station"] - 1.0):
        h = 1e-4
        num = (wc.expected_y(pv, x + h) - wc.expected_y(pv, x - h)) / (2 * h)
        assert wc.expected_slope(pv, x) == pytest.approx(num, abs=1e-6)


def test_originals_check_allows_only_the_known_openroads_stub():
    inp = good()
    pv = wc.normalised_pvis(inp)
    ours = {"alignment": "ALG", "x0": str(pv[0]["station"]), "x1": str(pv[-1]["station"])}
    base = _full_reader(inp)

    def reader(extra):
        def r(cmd, args):
            if cmd == "list_profiles":
                return [{"alignment": "ALG", "x0": "0", "x1": "100"}] + extra + [ours]
            return base(cmd, args)
        return r
    # first profile on an alignment: ORD adds its 0..100 stub, baseline was empty -> PASS
    assert wc.verify_profile_full(inp, reader([]), 1, baseline=[])["checks"]["originals_unchanged"]["status"] == "PASS"
    # any other new profile -> FAIL
    other = [{"alignment": "ALG", "x0": "5", "x1": "50"}]
    assert wc.verify_profile_full(inp, reader(other), 1, baseline=[])["checks"]["originals_unchanged"]["status"] == "FAIL"


# ---------------------------------------------------------------- engineer-run sheet steps (N5-N7) are read back
def _sheet_reader(boundary="present", ve="5", vspace="5", hspace="100", dvert="True", dterr="False", align="ALG", start=4650.0, stop=6050.0,
                  drawing=True, sheet=True, attach=True, scale=1 / 480):
    d = [{"key": "profileBoundary", "value": boundary}, {"key": "name", "value": "P"}, {"key": "group", "value": "G"}, {"key": "assocAlignment", "value": align},
         {"key": "startDistance", "value": str(start * 10000)}, {"key": "stopDistance", "value": str(stop * 10000)}, {"key": "profileHeight", "value": "500000"},
         {"key": "topClearance", "value": "10000"}, {"key": "botClearance", "value": "10000"}, {"key": "displayVertical", "value": dvert},
         {"key": "displayTerrain", "value": dterr}, {"key": "verticalExaggeration", "value": ve}, {"key": "verticalAxisSpacing", "value": vspace},
         {"key": "horizontalAxisSpacing", "value": hspace}]
    models = ([{"model": "Name=X1_cph_pro-1 Drawing; ModelType=Drawing; Id=3"}] if drawing else []) + ([{"model": "Name=X1_cph_pro-1 Sheet; ModelType=Sheet; Id=4"}] if sheet else [])
    atts = [{"model": "X1_cph_pro-1 Sheet", "attachment": f"model=X1_cph_pro-1 Drawing; file=a.dgn; scale={scale}; clipped=False"}] if attach else []

    def r(cmd, args):
        return {"profile_boundary_data": d, "list_models": models, "list_attachments": atts}[cmd]
    return r


def _sheet_inputs():
    inp = good(alignment="ALG", profile_name="X1_cph_pro-1")
    inp["pvis"] = [{"station": 1500.0, "elevation": 10.0, "curve_length": 0}, {"station": 1700.0, "elevation": 12.0, "curve_length": 0}]
    inp["units"] = "m"
    return inp


def test_sheet_verify_passes_the_pdf_settings_and_each_defect_fails():
    inp = _sheet_inputs()
    ok = wc.verify_profile_sheet(inp, _sheet_reader())
    assert ok["status"] == "PASS", ok["checks"]
    assert ok["engineer_confirms"]
    for kw, check in [({"boundary": "null"}, "boundary_exists"), ({"ve": "10"}, "boundary_settings"), ({"vspace": "10"}, "boundary_settings"),
                      ({"dvert": "False"}, "boundary_settings"), ({"align": "OTHER"}, "boundary_settings"), ({"start": 5200.0}, "boundary_range"),
                      ({"stop": 5000.0}, "boundary_range"), ({"drawing": False}, "models"), ({"sheet": False}, "models"),
                      ({"attach": False}, "sheet_attachment"), ({"scale": 1 / 1200}, "sheet_attachment")]:
        out = wc.verify_profile_sheet(inp, _sheet_reader(**kw))
        assert out["status"] == "FAIL" and out["checks"][check]["status"] == "FAIL", (kw, out["checks"])


def test_sheet_verify_reads_the_boundary_from_the_drawing_model_and_honours_an_explicit_range():
    inp = _sheet_inputs()
    seen = []
    base = _sheet_reader()

    def r(cmd, args):
        seen.append((cmd, dict(args)))
        return base(cmd, args)
    wc.verify_profile_sheet(inp, r)
    assert [a for c, a in seen if c == "profile_boundary_data"] == [{"name": "ALG", "model": "X1_cph_pro-1 Drawing"}]
    inp["profile_range_m"] = [1400.0, 1700.0]          # not the PVI table's 1500..1700: boundary 4650..6050 ft = 1417..1844 m starts after 1400
    out = wc.verify_profile_sheet(inp, _sheet_reader())
    assert out["checks"]["boundary_range"]["profile_m"] == [1400.0, 1700.0] and out["checks"]["boundary_range"]["status"] == "FAIL"


def test_sheet_verify_without_profile_name_cannot_identify_models():
    inp = _sheet_inputs()
    inp.pop("profile_name")
    out = wc.verify_profile_sheet(inp, _sheet_reader())
    assert out["checks"]["models"]["status"] == "UNVERIFIED" and out["status"] == "UNVERIFIED"


def test_apostrophe_in_profile_name_is_flagged():
    inp = _sheet_inputs()
    inp["profile_name"] = "X1's"
    out = wc.verify_profile_sheet(inp, _sheet_reader())
    assert out["checks"]["models"]["status"] == "FAIL"


def test_preflight_blocks_an_alignment_that_starts_at_zero():
    out = wc.preflight(WF, good(), reader_for(start=0.0, length=2000))
    assert out["status"] == wc.BLOCKED and any("10+00" in p for p in out["problems"])
    ok = wc.preflight(WF, good(), reader_for(start=900.0, length=2000))
    assert not any("10+00" in p for p in ok["problems"])
