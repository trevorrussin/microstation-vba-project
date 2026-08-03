"""Round-trip for 619-401.json (Family 3 intermediate shoulder)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-401.json").read_text(encoding="utf-8"))
ref301 = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

# Intermediate identity
for tok in ("INTERMEDIATE", "W20-5aR", "W21-5bR", "W21-5bU", "SHALL NOT EXCEED 20",
            "LEFT SHOULDER", "SYMMETRICAL", "1320"):
    if squash(tok) not in sb and tok not in body:
        fails.append(f"missing {tok}")
if "W21-5aR" in body:
    fails.append("401 uses W20-5aR not W21-5aR — PDF unexpectedly has W21-5aR")

# Buffer + shoulder overlap with 301
for row in spec["tables"]["401-03"]["rows"]:
    s = row["speedMph"]
    r301 = next(r for r in ref301["tables"]["301-03"]["rows"] if r["speedMph"] == s)
    if row["longitudinalBufferSpace"] != r301["longitudinalBufferSpace"]:
        fails.append(f"buffer mismatch at {s}")
    if row["shoulderTaper"] != r301["shoulderTaper"]:
        fails.append(f"shoulder mismatch at {s}")
    if "laneTaper" not in row:
        fails.append(f"401-03 missing laneTaper at {s}")

# Size table has W20-5aR
codes = {r["signCode"] for r in spec["tables"]["401-05"]["rows"]}
if "W20-5aR" not in codes or "W21-5bR" not in codes:
    fails.append(f"401-05 missing shoulder set: {codes}")

# Roll-ahead speed-keyed (not GVW)
rad0 = spec["tables"]["401-02"]["rows"][0]
if "speedBand" not in rad0 and "minMph" not in rad0 and "speedMph" not in str(rad0):
    # accept speedBand or similar
    if not any(k for k in rad0 if "speed" in k.lower() or "mph" in k.lower()):
        fails.append(f"401-02 expected speed-keyed roll-ahead, got keys {list(rad0)}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if "W20-5aR" not in labels or "W21-5bR" not in labels:
    fails.append(f"order missing intermediate signs: {labels}")
if "MERGING TAPER" in labels:
    fails.append("order has MERGING TAPER")
if labels[:3] != ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "SHOULDER TAPER"]:
    fails.append(f"unexpected upstream start: {labels[:3]}")

if spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append("must not declare advanceWarningSpacing")

print("order", labels)
print("codes", sorted(codes))
print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
