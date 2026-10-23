"""Completeness gate fixtures: change set (modified/deleted existing elements), standard-candidate dispositions,
requirement coverage derived from what was actually changed. A checker that passes a bad fixture is broken.
"""
from __future__ import annotations

import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa import changeset  # noqa: E402
from qa.checks import coverage  # noqa: E402
from tests.qa.test_qa_checks import el, ctx, _txt  # noqa: E402


def _rows(*specs):
    return [el(i, level="Default", typ=t, color="0", weight="0", box=box) for i, t, box in specs]


BEFORE = changeset.snapshot(_rows((1, "LINE", (0, 0, 10, 0)), (2, "LINE", (0, 5, 10, 5)), (3, "TEXT", (0, 9, 8, 10))))


def test_diff_finds_created_modified_deleted():
    after = _rows((1, "LINE", (0, 0, 10, 0)), (2, "LINE", (0, 7, 10, 7)), (9, "LINE", (1, 1, 2, 2)))
    d = changeset.diff(BEFORE, after)
    assert [r["elementId"] for r in d["created"]] == ["9"]
    assert [r["elementId"] for r in d["modified"]] == ["2"] and "rangeLowY" in d["modified_fields"]["2"]
    assert list(d["deleted"]) == ["3"]


def test_unchanged_model_has_empty_change_set():
    d = changeset.diff(BEFORE, _rows((1, "LINE", (0, 0, 10, 0)), (2, "LINE", (0, 5, 10, 5)), (3, "TEXT", (0, 9, 8, 10))))
    assert not d["created"] and not d["modified"] and not d["deleted"]


def _changes():
    after = _rows((1, "LINE", (0, 0, 10, 0)), (2, "LINE", (0, 7, 10, 7)), (9, "LINE", (1, 1, 2, 2)))
    return changeset.diff(BEFORE, after)


def test_undeclared_edit_or_delete_of_existing_elements_fails():
    assert coverage.changes({}, {"changes": _changes()})["status"] == "FAIL"


def test_declared_may_modify_passes():
    assert coverage.changes({"may_modify": True}, {"changes": _changes()})["status"] == "PASS"


def test_agent_touched_ids_are_not_collateral():
    c = {"changes": _changes(), "declared_ids": ["2", "3"]}
    assert coverage.changes({}, c)["status"] == "PASS"
    c["declared_ids"] = ["2"]  # element 3 deleted without being touched on purpose
    assert coverage.changes({}, c)["status"] == "FAIL"


def test_truncated_read_is_uncertain_not_pass():
    assert coverage.changes({}, {"changes": {"error": "truncated"}})["status"] == "UNCERTAIN"


def test_missing_snapshot_is_flagged_not_silent():
    assert coverage.changes({}, {})["status"] == "WARN"


def test_undispositioned_candidates_block():
    t = {"candidates": [{"cite": "HDM 20.9.6", "doc": "d", "applies": None, "reason": ""}]}
    assert coverage.dispositions(t, {})["status"] == "UNCERTAIN"
    t["candidates"][0].update(applies="no", reason="x")  # reason too short
    assert coverage.dispositions(t, {})["status"] == "UNCERTAIN"
    t["candidates"][0].update(applies="no", reason="not a text task, only linework")
    assert coverage.dispositions(t, {})["status"] == "PASS"


def test_text_change_brings_in_text_requirements_even_if_agent_never_mentions_text():
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    reqs, _ = coverage.applicable(c)
    assert {"R-TXT-SIZE", "R-TXT-STYLE", "R-TXT-OVL", "R-LVL"} <= set(reqs)


def test_linework_only_change_does_not_pull_text_requirements():
    reqs, _ = coverage.applicable(ctx([el(1, level="Default", typ="LINE", color="0", weight="0")]))
    assert "R-TXT-STYLE" not in reqs and "R-LVL" in reqs


def test_wztc_and_profile_kinds_derived_from_level_and_model():
    c = ctx([el(1)], model_name="727.21.101 1960s CL1 - Profile 1")
    reqs, _ = coverage.applicable(c)
    assert "R-WZ-GEOM" in reqs and "R-PROF-ANNOT" in reqs


def test_coverage_never_claims_full_compliance_when_items_unverified():
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    c["results"] = {"CAD-LVL-01": {"status": "PASS"}, "CAD-SYM-01": {"status": "PASS"}, "CAD-TXT-02": {"status": "PASS"}}
    r = coverage.report({}, c)
    assert r["status"] == "WARN" and "not full NYSDOT/DDC compliance" in r["observed"]
    rows = {x["requirement"]: x["status"] for x in c["coverage"]["requirements"]}
    assert rows["R-TXT-STYLE"] == "verified"
    assert rows["R-TXT-SIZE"] == "unverified"       # CAD-TXT-01 did not run -> not counted as verified
    assert rows["R-TXT-WORD"] == "engineer-review"  # manual requirement is listed, not hidden


def test_failed_check_is_reported_as_failed_not_verified():
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    c["results"] = {"CAD-TXT-02": {"status": "FAIL"}}
    coverage.report({}, c)
    rows = {x["requirement"]: x["status"] for x in c["coverage"]["requirements"]}
    assert rows["R-TXT-STYLE"] == "FAILED"


def test_applicable_candidate_with_no_registry_check_is_a_gap():
    c = ctx([_txt(900, 0.007, 0.007, (0, 0, 5, 1))])
    t = {"candidates": [{"cite": "Chapter 99 - 99.9.9 Something (p.1)", "doc": "nysdot/hdm/chapt_99.pdf", "applies": "yes", "reason": "applies to this"}]}
    coverage.report(t, c)
    assert any("99.9.9" in g for g in c["coverage"]["candidate_gaps"])


def test_registry_and_criteria_are_consistent():
    reg = coverage.registry()
    crit = {r["id"] for r in yaml.safe_load(open(os.path.join(ROOT, "qa", "criteria.yaml"), encoding="utf-8"))}
    for kind in reg["kinds"]:
        for r in kind["requirements"]:
            assert r["check"] == "manual" or r["check"] in crit, r["id"]
    fm = {r["id"] for r in yaml.safe_load(open(os.path.join(ROOT, "qa", "failure_modes.yaml"), encoding="utf-8"))}
    assert {"FM-013", "FM-014"} <= fm
