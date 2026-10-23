"""Measure standards retrieval quality (recall@k, MRR) on tests/qa/standards_eval.yaml.

  python scripts/standards_eval.py                 # full pipeline (BM25 + dense + rerank)
  python scripts/standards_eval.py --no-dense --no-rerank   # ablation
Keep an index/retrieval change only if recall@5 does not drop.
"""
from __future__ import annotations

import argparse
import os
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
import standards_index as si  # noqa: E402


def is_hit(h, row):
    if row["doc"].lower() not in h["doc"].lower():
        return False
    if row.get("section") and not str(h.get("section") or "").startswith(str(row["section"])) \
            and str(row["section"]) not in (h.get("path") or ""):
        return False
    if row.get("page"):
        a, _, b = str(h["pages"]).partition("-")
        lo, hi = int(a), int(b or a)
        if not (lo - 1 <= int(row["page"]) <= hi + 1):
            return False
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-k", type=int, default=5)
    ap.add_argument("--no-dense", action="store_true")
    ap.add_argument("--no-rerank", action="store_true")
    a = ap.parse_args()
    rows = yaml.safe_load(open(os.path.join(ROOT, "tests", "qa", "standards_eval.yaml"), encoding="utf-8"))
    hits, rr = 0, 0.0
    for row in rows:
        res = si.search(row["q"], agency=row.get("agency"), k=a.k, dense=not a.no_dense, rerank=not a.no_rerank, follow_refs=False)
        rank = next((i + 1 for i, h in enumerate(res) if is_hit(h, row)), None)
        hits += rank is not None
        rr += 1.0 / rank if rank else 0.0
        top = res[0]["cite"][:90] if res else "-"
        print(f"{'HIT ' + str(rank) if rank else 'MISS ':<6} {row['q'][:60]:<60} | top: {top}")
    n = len(rows)
    print(f"\nrecall@{a.k} = {hits}/{n} = {hits/n:.2f}   MRR = {rr/n:.2f}   "
          f"(dense={'off' if a.no_dense else 'on'}, rerank={'off' if a.no_rerank else 'on'})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
