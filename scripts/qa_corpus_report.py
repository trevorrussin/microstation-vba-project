"""Print the QA gate regression-corpus table and write outputs/qa/corpus-report.md.

  python scripts/qa_corpus_report.py          # exit 1 if any defect passes or any legitimate edit is rejected
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from qa import corpus  # noqa: E402
from tests.qa.corpus_cases import cases  # noqa: E402


def main() -> int:
    rep = corpus.evaluate(cases())
    md = corpus.markdown(rep)
    out = os.path.join(ROOT, "outputs", "qa", "corpus-report.md")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(md)
    print(md)
    print(f"written: {out}")
    return 1 if rep["false_pass"] or rep["false_fail"] else 0


if __name__ == "__main__":
    sys.exit(main())
