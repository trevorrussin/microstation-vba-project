"""Dump 318 plan-side words in reading order for corridor reconstruction."""
from __future__ import annotations

import pathlib
from collections import defaultdict

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]
doc = fitz.open(str(ROOT / "Bridge/captures/619-318.pdf"))
p = doc[0]
words = [w for w in p.get_text("words") if w[0] < 450]

rows = defaultdict(list)
for w in words:
    rows[round(w[1] / 6)].append(w)

print("=== PLAN WORDS (x<450) by y ===")
for k in sorted(rows, reverse=True):
    toks = sorted(rows[k], key=lambda w: w[0])
    line = " ".join(f"{w[4]}@{w[0]:.0f}" for w in toks)
    print(f"y~{k*6:6.1f}: {line}")

print("\n=== NOTES column (x 800-1200, y<350) ===")
notes = [w for w in p.get_text("words") if 800 <= w[0] <= 1200 and w[1] < 350]
nrows = defaultdict(list)
for w in notes:
    nrows[round(w[1] / 4)].append(w)
for k in sorted(nrows):
    toks = sorted(nrows[k], key=lambda w: w[0])
    print(f"y~{k*4:6.1f}: {' '.join(w[4] for w in toks)}")
