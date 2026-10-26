"""Trajectory-quality report over all QA tasks: python scripts/qa_trajectory_report.py [--out outputs/qa/trajectory-report.md]"""
from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

from qa import trajectory  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "outputs", "qa", "trajectory-report.md"))
    a = ap.parse_args()
    rows = trajectory.load_all()
    md = trajectory.markdown(rows, trajectory.summarize(rows))
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    open(a.out, "w", encoding="utf-8").write(md)
    print(md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
