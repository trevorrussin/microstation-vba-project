"""Batch round-trip for Family 7 mobile sheets (110, 111, 112)."""
from __future__ import annotations

import json
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import squash  # noqa: E402

SHEETS = [110, 111, 112]


def body(n: int) -> str:
    doc = fitz.open(str(ROOT / f"Bridge/captures/619-{n}.pdf"))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def check(n: int) -> list[str]:
    fails: list[str] = []
    spec = json.loads((ROOT / f"Data/sheet-specs/619-{n}.json").read_text(encoding="utf-8"))
    b = body(n)
    sb = squash(b)
    roles = spec["tableRoles"]

    if roles.get("taperAndBuffer") or roles.get("advanceWarningSpacing"):
        fails.append("must not declare taperAndBuffer or advanceWarningSpacing")

    # Sign / size sync on primary signSizes role
    size_id = roles["signSizes"]
    size_codes = {r["signCode"] for r in spec["tables"][size_id]["rows"]}
    sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
    if size_codes != sign_codes:
        fails.append(f"sign/size mismatch {size_codes ^ sign_codes}")

    labels = [
        r.get("label") or r.get("signCode")
        for al in spec["orderTable"]["alignments"]
        for r in al["rows"]
    ]

    if n == 110:
        for tok in ("W8-23", "MOBILE", "PVH+TMIA", "200/5", "240/6", "160/4", "9,500", "22,000", "619-205"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if "MERGING TAPER" in b:
            fails.append("unexpected MERGING TAPER")
        ra = spec["tables"]["110-02"]["rows"]
        if ra[0]["min"]["ft"] != 200 or ra[0]["max"]["ft"] != 240:
            fails.append(f"110-02 light GVW unexpected {ra[0]}")
        if ra[1]["min"]["ft"] != 160 or ra[1]["max"]["ft"] != 200:
            fails.append(f"110-02 heavy GVW unexpected {ra[1]}")
        if labels != ["ROLL AHEAD DISTANCE", "W8-23"]:
            fails.append(f"unexpected order: {labels}")
        if roles["rollAheadDistance"] != "110-02":
            fails.append("rollAhead role not 110-02")

    elif n == 111:
        for tok in ("NYW8-33", "W4-2R", "W20-5R", "MOBILE", "200/5", "280/7", "160/4", "240/6", "619-206", "1500'"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if "MERGING TAPER" in b:
            fails.append("unexpected MERGING TAPER")
        # Sheet 1 must NOT be the only source — W20-5R is Sheet 2
        ra = spec["tables"]["111-05"]["rows"]
        if ra[0]["min"]["ft"] != 200 or ra[0]["max"]["ft"] != 280:
            fails.append(f"111-05 >=55 unexpected {ra[0]}")
        if labels != ["ROLL AHEAD DISTANCE", "W20-5R"]:
            fails.append(f"unexpected order: {labels}")
        # Both sheets' PV identical
        if spec["tables"]["111-01"]["rows"] != spec["tables"]["111-04"]["rows"]:
            fails.append("111-01 vs 111-04 PV mismatch")
        # Sheet 1 sizes omit W20-5R
        s1 = {r["signCode"] for r in spec["tables"]["111-03"]["rows"]}
        if "W20-5R" in s1:
            fails.append("111-03 should not include W20-5R")
        if "W20-5R" not in {r["signCode"] for r in spec["tables"]["111-06"]["rows"]}:
            fails.append("111-06 missing W20-5R")

    elif n == 112:
        for tok in ("NYW8-33", "W20-5AR", "W4-2R", "MOBILE", "200/5", "240/6", "160/4", "9,500", "22,000", "1500'"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if "MERGING TAPER" in b:
            fails.append("unexpected MERGING TAPER")
        ra = spec["tables"]["112-05"]["rows"]
        if ra[0]["minGvwLbs"] != 9500 or ra[0]["max"]["ft"] != 240:
            fails.append(f"112-05 light unexpected {ra[0]}")
        if labels != ["ROLL AHEAD DISTANCE", "W20-5AR"]:
            fails.append(f"unexpected order: {labels}")
        s1 = {r["signCode"] for r in spec["tables"]["112-03"]["rows"]}
        if "W4-2R" in s1:
            fails.append("112-03 should not include W4-2R")
        if "W4-2R" not in {r["signCode"] for r in spec["tables"]["112-06"]["rows"]}:
            fails.append("112-06 missing W4-2R")

    for p in ("15 MINUTE", "MOBILE WORK IS WORK"):
        if squash(p) not in sb:
            fails.append(f"note phrase missing: {p}")

    return fails


def main() -> int:
    total = 0
    for n in SHEETS:
        fails = check(n)
        print(f"619-{n}: ROUND-TRIP FAILURES: {len(fails)}")
        for f in fails:
            print(f"  {f}")
        total += len(fails)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
