"""Regression corpus engine: measures how often the gate lets unacceptable work through (false PASS) and how often
it rejects acceptable work (false FAIL).

A case is a task + the evidence context the gate would have built (change set, op records, geometry, ...) and the
outcome a correct gate must give. `mini_gate` runs the same check functions `qa_gate run` runs over that context
(everything except the live pieces: vision, MicroStation reads); `evaluate` aggregates the counts.

Cases live in tests/qa/corpus_cases.py; `python scripts/qa_corpus_report.py` prints the table and writes
outputs/qa/corpus-report.md; tests/qa/test_qa_corpus.py requires 0 false passes and 0 false fails.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

BLOCKING = ("FAIL", "UNCERTAIN")


@dataclass
class Case:
    id: str
    should_block: bool          # True: a correct gate must NOT pass this
    why: str                    # the defect injected (should_block) or the legitimate work (should pass)
    task: dict
    ctx: dict
    family: str = ""            # grouping for the report
    checks: list = field(default_factory=list)  # (name, fn) - defaults to DEFAULT_CHECKS


def default_checks():
    from qa.checks import cad_rules, coverage, operations_check, com_spot
    return [
        ("CAD-CHG-01", coverage.changes), ("OP-01", operations_check.operations), ("REQ-GEOM-01", operations_check.relations),
        ("CAD-LVL-01", cad_rules.levels), ("CAD-SYM-01", cad_rules.symbology), ("CAD-LOC-01", cad_rules.location),
        ("CAD-DUP-01", cad_rules.duplicates), ("CAD-TXT-02", cad_rules.text_style), ("CAD-TXT-03", cad_rules.text_overlap), ("CAD-TXT-05", cad_rules.text_model_clear),
        ("CAD-TXT-04", cad_rules.text_line_clearance), ("CAD-SCALE-01", cad_rules.scale_bar_scale), ("COM-01", com_spot.verify),
    ]


def mini_gate(case: Case) -> tuple[str, dict]:
    """-> (overall PASS|FAIL|UNCERTAIN, {check: status}). A checker that raises is UNCERTAIN, as in qa_gate."""
    out = {}
    for name, fn in case.checks or default_checks():
        try:
            out[name] = fn(case.task, case.ctx)["status"]
        except Exception as ex:  # noqa: BLE001
            out[name] = "UNCERTAIN"
            out[name + "_error"] = str(ex)
    if "results" in case.ctx:  # coverage needs prior results
        from qa.checks import coverage
        out["COV-01"] = coverage.report(case.task, case.ctx)["status"]
    st = [v for k, v in out.items() if not k.endswith("_error")]
    blocked = [s for s in st if s in BLOCKING]
    overall = "PASS" if not blocked else ("FAIL" if "FAIL" in blocked else "UNCERTAIN")
    return overall, out


def evaluate(cases: list[Case]) -> dict:
    rows, fp, ff, unc_block = [], [], [], 0
    n_block = sum(1 for c in cases if c.should_block)
    n_ok = len(cases) - n_block
    for c in cases:
        overall, detail = mini_gate(c)
        passed = overall == "PASS"
        if c.should_block and passed:
            fp.append(c.id)
        if (not c.should_block) and not passed:
            ff.append(c.id)
        if c.should_block and overall == "UNCERTAIN":
            unc_block += 1
        why_blocked = ", ".join(k for k, v in detail.items() if v in BLOCKING and not k.endswith("_error"))
        rows.append({"id": c.id, "family": c.family, "expected": "BLOCK" if c.should_block else "PASS", "got": overall,
                     "by": why_blocked, "ok": (not passed) == c.should_block, "why": c.why})
    return {"rows": rows, "cases": len(cases), "should_block": n_block, "should_pass": n_ok,
            "false_pass": fp, "false_fail": ff,
            "false_pass_rate": len(fp) / n_block if n_block else 0.0,
            "false_fail_rate": len(ff) / n_ok if n_ok else 0.0,
            "uncertain_rate": unc_block / n_block if n_block else 0.0}


def markdown(rep: dict) -> str:
    lines = ["# QA gate regression corpus", "",
             f"{rep['cases']} cases: {rep['should_block']} defects the gate must block, {rep['should_pass']} legitimate edits it must pass.", "",
             f"- false-pass rate: **{rep['false_pass_rate']:.1%}** ({len(rep['false_pass'])} of {rep['should_block']})  {rep['false_pass'] or ''}",
             f"- false-fail rate: **{rep['false_fail_rate']:.1%}** ({len(rep['false_fail'])} of {rep['should_pass']})  {rep['false_fail'] or ''}",
             f"- blocked as UNCERTAIN (not decidable) rather than FAIL: {rep['uncertain_rate']:.1%} of the defects", "",
             "| case | family | expected | got | caught by | injected defect / legitimate work |", "|---|---|---|---|---|---|"]
    for r in rep["rows"]:
        lines.append(f"| {r['id']} | {r['family']} | {r['expected']} | {r['got']}{'' if r['ok'] else ' **WRONG**'} | {r['by']} | {r['why']} |")
    return "\n".join(lines) + "\n"
