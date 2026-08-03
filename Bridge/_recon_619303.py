"""Recon 619-303 vs Family 2 reference 619-302: title, tables, notes, signs."""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

import fitz

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "scripts"))
from pdf_table_extract import group_rows, row_text, words_in_window  # noqa: E402

PDF = pathlib.Path("Bridge/captures/619-303.pdf")
pg = fitz.open(PDF)[0]
W = pg.get_text("words")
print(f"page size {pg.rect}  words={len(W)}  drawings={len(pg.get_drawings())}")

# Title / furniture
print("\n=== TITLE BLOCK (y<40 or y>750) ===")
for w in sorted(W, key=lambda t: (t[1], t[0])):
    if w[1] < 35 or w[1] > 760:
        pass
# Join top band
top = [w for w in W if w[1] < 50]
print("TOP:", " ".join(w[4] for w in sorted(top, key=lambda t: (round(t[1]), t[0])))[:300])
bot = [w for w in W if w[1] > 760]
print("BOT:", " ".join(w[4] for w in sorted(bot, key=lambda t: (round(t[1]), t[0])))[:400])

print("\n=== TABLE titles ===")
for w in sorted(W, key=lambda t: (t[1], t[0])):
    if w[4].upper() == "TABLE" or (w[4].startswith("TABLE") and len(w[4]) > 5):
        # get neighbors on same y
        band = [x for x in W if abs(x[1] - w[1]) < 4 and x[0] >= w[0] - 5]
        print(f"  y={w[1]:6.1f} x={w[0]:6.1f}  {' '.join(x[4] for x in sorted(band, key=lambda t: t[0]))}")

print("\n=== NOTES header hits ===")
for w in W:
    if "NOTE" in w[4].upper() and w[0] > 400:
        print(f"  y={w[1]:6.1f} x={w[0]:6.1f} {w[4]}")

# Sign-ish diamond labels / MUTCD codes near plan
print("\n=== likely sign codes / legends (plan+notes, x<700) ===")
codes = []
for w in W:
    t = w[4]
    if any(t.startswith(p) for p in ("W20", "W4", "W04", "G20", "NYW", "R2", "NYR")):
        codes.append((w[1], w[0], t))
print(sorted(codes)[:40])

# Compare table id prefixes to 302
print("\n=== speed tokens in right half (tables) ===")
speeds = defaultdict(list)
for w in W:
    if w[0] > 700 and w[4].isdigit() and int(w[4]) in (25, 30, 35, 40, 45, 50, 55, 60, 65):
        speeds[w[4]].append(round(w[1], 1))
for s in sorted(speeds, key=int):
    print(f"  {s}: {sorted(set(speeds[s]))[:12]}")

# Load 302 corridor zone labels for comparison later
ref = json.loads(pathlib.Path("Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
print("\n=== 302 corridor zone labels (reference) ===")
for z in ref["corridor"]["zones"]:
    print(f"  {z['order']:2d} {z['id']:<22} {z.get('sheetLabel','')}")
print("signs:", [s["signCode"] for s in ref["signs"]["items"]])
print("notes count:", len(ref["notes"]["printed"]))
