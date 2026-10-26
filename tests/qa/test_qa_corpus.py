"""The regression corpus must hold: no defect passes (false-pass rate 0) and no legitimate edit is rejected."""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from qa import corpus  # noqa: E402
from tests.qa.corpus_cases import cases  # noqa: E402


def test_the_corpus_is_big_enough_to_mean_something():
    cs = cases()
    assert len(cs) >= 50
    assert sum(1 for c in cs if c.should_block) >= 30 and sum(1 for c in cs if not c.should_block) >= 12
    assert len({c.id for c in cs}) == len(cs)


def test_no_defect_passes_and_no_legitimate_edit_is_rejected():
    rep = corpus.evaluate(cases())
    wrong = [r for r in rep["rows"] if not r["ok"]]
    assert not rep["false_pass"], f"false passes (defects the gate let through): {rep['false_pass']}\n{wrong}"
    assert not rep["false_fail"], f"false fails (legitimate work rejected): {rep['false_fail']}\n{wrong}"


def test_each_defect_is_blocked_by_a_named_check():
    rep = corpus.evaluate(cases())
    assert all(r["by"] for r in rep["rows"] if r["expected"] == "BLOCK")  # a block with no named check would be an accident


def test_the_report_renders():
    md = corpus.markdown(corpus.evaluate(cases()))
    assert "false-pass rate" in md and "| OP-move-wrong-distance |" in md
