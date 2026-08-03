"""Batch round-trip for all Family 5 sheet specs."""
from __future__ import annotations

import json
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import squash  # noqa: E402

SPEC302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
SHEETS = [318, 316, 319, 113, 211, 416, 417, 418, 517, 518]


def body(n: int) -> str:
    doc = fitz.open(str(ROOT / f"Bridge/captures/619-{n}.pdf"))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def check_sheet(n: int) -> list[str]:
    fails = []
    spec = json.loads((ROOT / f"Data/sheet-specs/619-{n}.json").read_text(encoding="utf-8"))
    b = body(n)
    sb = squash(b)

    # Sign / size sync
    roles = spec["tableRoles"]
    size_id = roles.get("signSizes")
    if size_id:
        size_codes = {r["signCode"] for r in spec["tables"][size_id]["rows"]}
        sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
        if size_codes != sign_codes:
            fails.append(f"sign/size mismatch {size_codes ^ sign_codes}")

    # No AW role (Family 5 uses fixed gaps)
    if roles.get("advanceWarningSpacing"):
        fails.append("must not declare advanceWarningSpacing")

    # Taper identity vs 302 where applicable
    taper_id = roles.get("taperAndBuffer")
    if taper_id and n in (318, 319, 418, 518, 316):
        for row in spec["tables"][taper_id]["rows"]:
            r302 = next(
                (r for r in SPEC302["tables"]["302-02"]["rows"] if r["speedMph"] == row["speedMph"]),
                None,
            )
            if not r302:
                continue
            if "laneTaper" in row and row.get("laneTaper") == r302["laneTaper"]:
                pass
            elif "laneTaper" in row and n in (318, 319, 418, 518):
                if row["laneTaper"] != r302["laneTaper"]:
                    fails.append(f"{taper_id} laneTaper speed={row['speedMph']} != 302")
            if row["longitudinalBufferSpace"] != r302["longitudinalBufferSpace"]:
                fails.append(f"{taper_id} buffer speed={row['speedMph']} != 302")

    # Seven-band sheets: key tokens
    if n in (416, 417, 517):
        for tok in ("80/2/3", "160/4/5", "200/5/6", "280/7/8", "360/9"):
            if tok not in b:
                fails.append(f"missing seven-band token {tok}")

    # Plan gap tokens
    gaps_expected = {
        318: ("1000'", "1500'", "1320'"),
        319: ("1000'", "1500'", "1320'"),
        418: ("1000'", "1500'", "1320'"),
        417: ("1000'", "1500'", "1320'"),
        517: ("1000'", "1500'", "2640'"),
        518: ("1000'", "1500'", "2640'"),
        316: ("1000'", "1500'"),
        416: ("1000'", "500'"),
        113: ("1000'",),
        211: ("1000'",),
    }
    for tok in gaps_expected.get(n, ()):
        if tok not in b and tok.replace("'", "") not in b.replace("'", ""):
            fails.append(f"missing gap token {tok}")

    # Key signs on plan / tables
    sign_checks = {
        318: ("W20-1", "W20-5", "W4-2R", "G20-2", "NYW8-33", "MERGING"),
        319: ("W20-1", "W4-2R", "G20-2", "NYW8-33", "MERGING"),
        418: ("W20-1", "W4-2R", "G20-2", "MERGING"),
        417: ("W20-1", "W4-2R", "G20-2", "MERGING"),
        517: ("W20-1", "W4-2R", "G20-2", "MERGING"),
        518: ("W20-1", "W4-2R", "G20-2", "MERGING"),
        316: ("W20-1", "G20-2", "W21-5"),
        416: ("W20-1", "G20-2", "W21-5"),
        113: ("W21-5AL", "TMIA"),
        211: ("W21-5aL", "W20-1"),
    }
    for tok in sign_checks.get(n, ()):
        if tok not in b:
            fails.append(f"missing plan/table token {tok}")

    # Order table sanity
    labels = [r.get("label") or r.get("signCode")
              for al in spec["orderTable"]["alignments"] for r in al["rows"]]
    if n in (318, 319, 417, 418, 517, 518):
        if labels[0] != "ROLL AHEAD DISTANCE":
            fails.append(f"upstream start {labels[0]}")
        if "MERGING TAPER" not in labels and n not in (316, 416):
            fails.append("missing MERGING TAPER")
        if "G20-2" not in labels:
            fails.append("missing G20-2")
    if n in (316, 416):
        if "SHOULDER TAPER" not in labels:
            fails.append("partial-exit must have SHOULDER TAPER")
        if "MERGING TAPER" in labels:
            fails.append("partial-exit must not sequential MERGING TAPER")
    if n in (113, 211):
        if "ROLL AHEAD DISTANCE" not in labels:
            fails.append("mobile/short must have ROLL AHEAD")
        if "MERGING TAPER" in labels:
            fails.append("mobile must not have MERGING TAPER")

    # Note phrases
    if n == 318:
        for p in ("SHORT-TERM STATIONARY", "ENTRANCE RAMP", "REGULATORY SPEED LIMIT"):
            if squash(p) not in sb:
                fails.append(f"note phrase missing: {p}")

    return fails


def main() -> int:
    total = 0
    for n in SHEETS:
        fails = check_sheet(n)
        total += len(fails)
        status = "PASS" if not fails else f"FAIL ({len(fails)})"
        print(f"619-{n}: {status}")
        for f in fails:
            print(f"  {f}")
    print(f"TOTAL FAILURES: {total}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
