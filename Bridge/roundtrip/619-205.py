"""Round-trip for 619-205.json (Family 3 short-duration shoulder)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-205.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

# Short duration — no taper/buffer table
roles = spec["tableRoles"]
if roles.get("taperAndBuffer") or roles.get("advanceWarningSpacing"):
    fails.append("205 must not declare taperAndBuffer or advanceWarningSpacing")

for tok in ("W21-5", "W20-1", "SHORT DURATION", "CAUTION MODE", "ROLL AHEAD"):
    if squash(tok) not in sb and tok not in body:
        fails.append(f"missing {tok}")

# Sign sizes
codes = {r["signCode"] for r in spec["tables"]["205-03"]["rows"]}
if codes != {"W20-1", "W21-5", "WARNING FLAG"}:
    fails.append(f"205-03 unexpected codes: {codes}")

# Order: roll + two signs only (no taper/buffer)
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels != ["ROLL AHEAD DISTANCE", "W21-5", "W20-1"]:
    fails.append(f"unexpected order: {labels}")
if "SHOULDER TAPER" in labels or "BUFFER SPACE" in labels:
    fails.append("short-duration order must omit taper/buffer")

# No MERGING TAPER
if "MERGING TAPER" in body:
    fails.append("unexpected MERGING TAPER")

print("roles", {k: v for k, v in roles.items() if k != "note"})
print("order", labels)
print("codes", sorted(codes))
print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
