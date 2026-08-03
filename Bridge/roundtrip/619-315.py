"""Round-trip for 619-315.json (Family 3 ramp-approach short-term)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-315.json").read_text(encoding="utf-8"))
ref301 = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

if spec["sheet"].get("pageRotation") != 270:
    fails.append(f"expected pageRotation=270, got {spec['sheet'].get('pageRotation')}")

# Speeds match 301
speeds = [r["speedMph"] for r in spec["tables"]["315-03"]["rows"]]
if speeds != [45, 50, 55, 65]:
    fails.append(f"315-03 speeds {speeds}")

# Buffer + first three shoulder bands match 301-03
for row in spec["tables"]["315-03"]["rows"]:
    s = row["speedMph"]
    r301 = next(r for r in ref301["tables"]["301-03"]["rows"] if r["speedMph"] == s)
    if row["longitudinalBufferSpace"] != r301["longitudinalBufferSpace"]:
        fails.append(f"buffer mismatch at {s}")
    for b in ("<= 4 ft", "5 - 7 ft", ">= 8 ft"):
        if row["shoulderTaper"][b] != r301["shoulderTaper"][b]:
            fails.append(f"shoulder {b} mismatch at {s}")

# Ramp gap C = 2640 (not 301's 1320)
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
# Find W20-1 length
w20 = next(r for al in spec["orderTable"]["alignments"] for r in al["rows"] if r.get("signCode") == "W20-1")
# length comes from zone_length resolve — check corridor gapC
gap_c = next(z for z in spec["corridor"]["zones"] if z["id"] == "gapC")
src = gap_c["lengthSource"]
ft = src.get("fixedFt") if isinstance(src, dict) else None
if ft != 2640:
    fails.append(f"gapC should be 2640 fixedFt, got {src}")
if "2640" not in body:
    fails.append("PDF missing 2640 token")

for tok in ("RAMP APPROACH", "W21-5aR", "W21-5bR", "W3-7a", "LATERAL", "GROSS VEHICLE WEIGHT"):
    if squash(tok) not in sb and tok not in body:
        fails.append(f"missing {tok}")

if roles_aw := spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append(f"must not declare advanceWarningSpacing ({roles_aw})")

if "MERGING TAPER" in labels:
    fails.append("order has MERGING TAPER")
if labels[:3] != ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "SHOULDER TAPER"]:
    fails.append(f"unexpected upstream start: {labels[:3]}")

# Roll-ahead GVW like 301
if "315-02" != spec["tableRoles"].get("rollAheadDistance"):
    fails.append("rollAhead should be 315-02")
rad0 = spec["tables"]["315-02"]["rows"][0]
if "minGvwLbs" not in rad0 and "gvwBand" not in rad0:
    fails.append("315-02 must be GVW-keyed")

print("order", labels)
print("gapC", src)
print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
