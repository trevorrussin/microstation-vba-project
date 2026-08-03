"""Round-trip for 619-301.json (Family 3 shoulder-closure reference).

Page rotation=270 — prefer sibling/token checks + targeted windows.
"""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))
ref302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
fails = []

# Speeds
speeds = [r["speedMph"] for r in spec["tables"]["301-03"]["rows"]]
if speeds != [45, 50, 55, 65]:
    fails.append(f"301-03 speeds {speeds}")
else:
    print("301-03 speeds OK")

# Overlap with 302 shoulder/buffer
for row in spec["tables"]["301-03"]["rows"]:
    s = row["speedMph"]
    r302 = next(r for r in ref302["tables"]["302-02"]["rows"] if r["speedMph"] == s)
    if row["longitudinalBufferSpace"] != r302["longitudinalBufferSpace"]:
        fails.append(f"buffer mismatch at {s}")
    if row["shoulderTaper"] != r302["shoulderTaper"]:
        fails.append(f"shoulder taper mismatch at {s}")
print("301-03 vs 302-02 overlap OK" if not fails else "overlap diffs pending")

# Tokens
for tok in ("W21-5aR", "W21-5bR", "W20-1", "PVH", "1320'", "1500'", "1000'", "SHOULDER"):
    if tok not in body:
        fails.append(f"missing {tok}")
if "MERGING" in body and "MERGING TAPER" in body:
    fails.append("unexpected MERGING TAPER on shoulder sheet")
print("tokens checked")

# Roll ahead GVW bands
rad = " ".join(
    f"{r['min']['ft']}/{r['min']['skipLines']} {r['max']['ft']}/{r['max']['skipLines']}"
    for r in spec["tables"]["301-02"]["rows"]
)
for part in ("160/4", "200/5", "120/3", "160/4"):
    if part not in body and part not in rad:
        fails.append(f"roll-ahead ratio {part}")
# ratios should appear in PDF
for part in ("160/4", "120/3"):
    if part not in body:
        fails.append(f"PDF missing roll-ahead {part}")
print("roll-ahead ratios")

# No AW table role
if spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append("must not declare advanceWarningSpacing")

# Notes phrases
phrases = [
    "SHORT-TERM STATIONARY",
    "LEFT SHOULDER CLOSURES ARE SYMMETRICAL",
    "W21-5bL",
    "W21-5aL",
    "SHALL NOT EXCEED 40'",
    "W7-3a",
    "G20-1",
    "REGULATORY SPEED LIMIT SIGN IS REQUIRED",
]
sb = squash(body)
for p in phrases:
    if squash(p) not in sb:
        fails.append(f"note phrase missing: {p}")
print("notes phrases", len(phrases))

# Order table
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if "MERGING TAPER" in labels:
    fails.append("order has MERGING TAPER")
if labels[:3] != ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "SHOULDER TAPER"]:
    fails.append(f"unexpected upstream start: {labels[:3]}")
if "W21-5aR" not in labels or "W21-5bR" not in labels:
    fails.append(f"missing shoulder signs in order: {labels}")
print("order", labels)

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
