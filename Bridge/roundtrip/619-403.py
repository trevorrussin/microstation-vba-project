"""Round-trip for 619-403.json.

Pages are rotation=270 (mediabox portrait, display landscape) — y/x windows
are fragile. Strategy: (1) tables that must match 402/303 are compared
JSON-to-JSON against the already-round-tripped siblings; (2) sheet-specific
tokens and Notes 1-8 are asserted present in the PDF text layer.

Run: python Bridge/roundtrip/619-403.py
"""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-403.json").read_text(encoding="utf-8"))
ref402 = json.loads((ROOT / "Data/sheet-specs/619-402.json").read_text(encoding="utf-8"))
ref303 = json.loads((ROOT / "Data/sheet-specs/619-303.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
assert doc.page_count >= 2
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
fails = []


def squash_glyphs(s: str) -> str:
    return squash(s).replace(">=", "W").replace("<=", "L")


# ---- identical-to-sibling table cells (already verified by extract script)
pairs = [
    ("403-01", spec["tables"]["403-01"]["rows"], ref402["tables"]["402-01"]["rows"]),
    ("403-02", spec["tables"]["403-02"]["rows"], ref402["tables"]["402-02"]["rows"]),
    ("403-03", spec["tables"]["403-03"]["rows"], ref402["tables"]["402-03"]["rows"]),
]
for label, a, b in pairs:
    if a != b:
        fails.append(f"{label} rows differ from sibling reference (expected identical)")
    else:
        print(f"{label}: identical to sibling ({len(a)} rows)")

# 65mph/12ft must appear as printed on this PDF
if "800/20/21" not in body:
    fails.append("PDF text missing 800/20/21 (65mph/12ft taper cell)")
else:
    print("800/20/21 present in PDF")

# two-lane markers
for tok in ("W20-5a", "2L"):
    if tok not in body:
        fails.append(f"PDF missing {tok}")
    else:
        print(f"{tok} present")

# intermediate markers
if "PVH" not in body and "PVH+TMIA" not in body:
    fails.append("PDF missing PVH codes")
if "20'" not in body and "EXCEED 20" not in squash(body).upper():
    # Note 4
    if "20" not in body:
        fails.append("PDF missing 20' spacing cue")

# advance URBAN A=100 present; FREEWAY inferred flagged
aw = spec["tables"]["403-06"]["rows"]
freeway = [r for r in aw if r["roadType"] == "FREEWAY"]
if not freeway:
    fails.append("403-06 missing FREEWAY row")
elif freeway[0].get("confidence") == "inferred":
    print("403-06 FREEWAY present (inferred C=2640 — knownAnomaly)")
if not any(r["roadType"] == "RURAL" and r["A"] == 500 for r in aw):
    fails.append("403-06 RURAL A!=500")

# sign sizes include W20-5a
codes = {r["signCode"] for r in spec["tables"]["403-05"]["rows"]}
if "W20-5a" not in codes:
    fails.append("403-05 missing W20-5a")
if "W20-5R" in codes or "W20-5" in codes:
    fails.append("403-05 should use W20-5a, not one-lane W20-5/W20-5R")
print("403-05 codes:", sorted(codes))

# notes 1-8 — rotation=270 interleaves columns; check distinctive phrases
phrases = [
    "INTERMEDIATE-TERM IS STATIONARY WORK",
    "RIGHT LANE CLOSURES ARE SYMMETRICAL TO LEFT",
    "W20-5a",
    "PROTECTIVE VEHICLE(S) SHALL MAINTAIN",
    "SHALL NOT EXCEED 20'",
    "NO WORK ACTIVITY, EQUIPMENT, OR STORAGE",
    "TRANSVERSELY A MINIMUM OF EVERY 800'",
    "NY9-11 SIGN IS RECOMMENDED",
    "REGULATORY SPEED LIMIT SIGN IS REQUIRED HALFWAY",
]
sb = squash_glyphs(body)
for p in phrases:
    if squash_glyphs(p) not in sb:
        fails.append(f"note phrase missing: {p}")
print("note phrases checked:", len(phrases))

# order table sanity
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"]
          for r in al["rows"]]
if labels.count("MERGING TAPER") != 2 or "2L" not in labels:
    fails.append(f"orderTable missing dual taper+2L: {labels}")

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
