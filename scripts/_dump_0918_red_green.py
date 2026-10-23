"""Dump 09.18 BTS page-3 red deletions + green replacements for callout item fixes."""
from __future__ import annotations

from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
page = fitz.open(PDF)[2]
words = page.get_text("words")

print("=== ALL green FreeTexts ===")
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "FreeText":
        continue
    stroke = (a.colors or {}).get("stroke")
    content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    if not content:
        continue
    r = a.rect
    print(f"\n[{i}] stroke={stroke} @({r.x0:.0f},{r.y0:.0f})")
    print(content[:500])

print("\n\n=== ALL red marks covering words ===")
for i, a in enumerate(page.annots() or []):
    stroke = (a.colors or {}).get("stroke") or []
    fill = (a.colors or {}).get("fill") or []
    is_red = False
    for c in (stroke, fill):
        if c and len(c) >= 3 and c[0] > 0.7 and c[1] < 0.35 and c[2] < 0.35:
            is_red = True
    if not is_red:
        continue
    r = a.rect
    cov = sorted(
        [
            w
            for w in words
            if w[0] < r.x1 and w[2] > r.x0 and w[1] < r.y1 and w[3] > r.y0
        ],
        key=lambda w: (round(w[1], 0), w[0]),
    )
    if not cov:
        continue
    txt = " ".join(w[4] for w in cov)
    # context: nearby words in same y band
    y0 = min(w[1] for w in cov) - 2
    y1 = max(w[3] for w in cov) + 2
    x0 = min(w[0] for w in cov) - 80
    x1 = max(w[2] for w in cov) + 80
    ctx = sorted(
        [
            w
            for w in words
            if w[0] < x1 and w[2] > x0 and w[1] < y1 and w[3] > y0
        ],
        key=lambda w: (round(w[1], 0), w[0]),
    )
    ctx_txt = " ".join(w[4] for w in ctx)
    print(f"\n[{i}] {a.type[1]} @({r.x0:.0f},{r.y0:.0f}) size={r.width:.0f}x{r.height:.0f}")
    print(f"  STRIKE: {txt!r}")
    print(f"  CONTEXT: {ctx_txt[:220]!r}")
