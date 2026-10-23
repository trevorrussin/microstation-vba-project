"""List non-digit red strikes on 09.18 BTS page 3."""
from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
page = fitz.open(PDF)[2]
words = page.get_text("words")
print("=== RED marks (skip pure item-digit overlays) ===")
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
    digits_only = all(ch.isdigit() or ch in " ." for ch in txt)
    if digits_only and len(txt) < 24:
        continue
    y0 = min(w[1] for w in cov) - 3
    y1 = max(w[3] for w in cov) + 3
    x0 = min(w[0] for w in cov) - 100
    x1 = max(w[2] for w in cov) + 100
    ctx = sorted(
        [
            w
            for w in words
            if w[0] < x1 and w[2] > x0 and w[1] < y1 and w[3] > y0
        ],
        key=lambda w: (round(w[1], 0), w[0]),
    )
    ctx_txt = " ".join(w[4] for w in ctx)
    print(f"[{i}] {a.type[1]} @({r.x0:.0f},{r.y0:.0f}) {r.width:.0f}x{r.height:.0f}")
    print(f"  STRIKE: {txt!r}")
    print(f"  CTX: {ctx_txt[:260]!r}")
