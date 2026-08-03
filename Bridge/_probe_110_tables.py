"""Precise cell dump + crops for 619-110 tables."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parent.parent
CAP = ROOT / "Bridge" / "captures"

doc = fitz.open(str(CAP / "619-110.pdf"))
pg = doc[0]
# Full high-res render then crop with pixmap subset
pix = pg.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), annots=False)
print("full", pix.width, pix.height)

# Save full
pix.save(str(CAP / "sheet_619110_hi.png"))

# Dump ALL words sorted, focusing on numbers near tables
words = list(pg.get_text("words"))
# Find table title positions
for w in words:
    if w[4] in ("110-01:", "110-02:", "110-03:", "TABLE") or "9,500" in w[4] or "22,000" in w[4] or "/" in w[4] and w[4][0].isdigit():
        print(f"  {w[4]!r:20s} x={w[0]:.1f} y={w[1]:.1f}")

print("\n=== ROLL AHEAD WINDOW (x 620-760, y 900-1175) ===")
sel = [w for w in words if 620 <= w[0] <= 760 and 900 <= w[1] <= 1175]
rows = defaultdict(list)
for w in sel:
    rows[round(w[1])].append(w)
for y in sorted(rows):
    cells = sorted(rows[y], key=lambda c: c[0])
    print(f"y={y:4d}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))

print("\n=== SIGN SIZE WINDOW (x 580-650, y 900-1100) ===")
sel = [w for w in words if 580 <= w[0] <= 650 and 900 <= w[1] <= 1100]
rows = defaultdict(list)
for w in sel:
    rows[round(w[1])].append(w)
for y in sorted(rows):
    cells = sorted(rows[y], key=lambda c: c[0])
    print(f"y={y:4d}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))

print("\n=== PV WINDOW (x 640-760, y 640-860) ===")
sel = [w for w in words if 640 <= w[0] <= 760 and 640 <= w[1] <= 860]
rows = defaultdict(list)
for w in sel:
    rows[round(w[1] / 3) * 3].append(w)
for y in sorted(rows):
    cells = sorted(rows[y], key=lambda c: c[0])
    print(f"y={y:4d}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))

# Also dump 112 page0 roll table same way
print("\n\n######## 112 page 0 roll/sizes ########")
doc2 = fitz.open(str(CAP / "619-112.pdf"))
for pi in range(2):
    pg = doc2[pi]
    words = list(pg.get_text("words"))
    print(f"\n--- page {pi} roll window ---")
    sel = [w for w in words if 620 <= w[0] <= 760 and 900 <= w[1] <= 1175]
    rows = defaultdict(list)
    for w in sel:
        rows[round(w[1])].append(w)
    for y in sorted(rows):
        cells = sorted(rows[y], key=lambda c: c[0])
        print(f"y={y:4d}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))
    print(f"--- page {pi} sign sizes ---")
    sel = [w for w in words if 560 <= w[0] <= 650 and 900 <= w[1] <= 1100]
    rows = defaultdict(list)
    for w in sel:
        rows[round(w[1])].append(w)
    for y in sorted(rows):
        cells = sorted(rows[y], key=lambda c: c[0])
        print(f"y={y:4d}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))
