"""Batch round-trip for Family 9 + Misc sheets."""
from __future__ import annotations

import json
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import squash  # noqa: E402

CORRIDOR = [21, 22, 23, 31, 32, 33, 60]
REF = [1, 4, 5, 6, 10, 12, 80]


def body_of(n: int) -> str:
    spec = json.loads((ROOT / f"Data/sheet-specs/619-{n:03d}.json").read_text(encoding="utf-8"))
    doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def check_corridor(n: int) -> list[str]:
    fails: list[str] = []
    spec = json.loads((ROOT / f"Data/sheet-specs/619-{n:03d}.json").read_text(encoding="utf-8"))
    b = body_of(n)
    sb = squash(b)
    roles = spec["tableRoles"]
    pref = f"{n:03d}"

    if roles.get("taperAndBuffer"):
        fails.append("must not declare taperAndBuffer")

    size_id = roles.get("signSizes")
    if size_id:
        size_codes = {r["signCode"] for r in spec["tables"][size_id]["rows"]}
        sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
        if size_codes != sign_codes:
            fails.append(f"sign/size mismatch {size_codes ^ sign_codes}")

    labels = [
        r.get("label") or r.get("signCode")
        for al in spec["orderTable"]["alignments"]
        for r in al["rows"]
    ]

    if n == 21:
        for tok in ("W21-8", "36x36", "48x48", "MOWING", "500"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if labels != ["W21-8"]:
            fails.append(f"unexpected order: {labels}")
        if roles.get("protectiveVehicle") or roles.get("rollAheadDistance"):
            fails.append("021 must not declare PV/roll roles")

    elif n in (22, 23):
        for tok in ("W21-8", "PVH+TMIA", "PVL+TMIA", "200/5", "160/4", "120/3", "MOWING"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if "MERGING TAPER" in b:
            fails.append("unexpected MERGING TAPER")
        ra = spec["tables"][f"{pref}-02"]["rows"]
        if ra[0]["lightGvw"]["ft"] != 200 or ra[0]["heavyGvw"]["ft"] != 160:
            fails.append(f"{pref}-02 45-55 unexpected {ra[0]}")
        if labels != ["ROLL AHEAD DISTANCE", "W21-8"]:
            fails.append(f"unexpected order: {labels}")

    elif n == 31:
        for tok in ("W20-1", "W8-23", "200/5", "280/7", "160/4", "240/6", "120/3", "P,", "TMIA"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        ra = spec["tables"]["031-02"]["rows"]
        if ra[0]["max"]["ft"] != 280 or ra[2]["min"]["ft"] != 120:
            fails.append(f"031-02 unexpected {ra}")
        if labels != ["ROLL AHEAD DISTANCE", "W20-1"]:
            fails.append(f"unexpected order: {labels}")

    elif n == 32:
        for tok in ("W23-1", "NYW8-32", "PVH+TMIA", "200/5", "160/4", "120/3", "HERBICIDE"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if labels != ["ROLL AHEAD DISTANCE", "W23-1"]:
            fails.append(f"unexpected order: {labels}")

    elif n == 33:
        for tok in ("W20-1", "W8-23", "PVH+TMIA", "200/5", "240/6", "160/4", "FREEWAY"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        ra = spec["tables"]["033-02"]["rows"]
        if ra[0]["lightGvw"]["ft"] != 240 or ra[1]["heavyGvw"]["ft"] != 160:
            fails.append(f"033-02 unexpected {ra}")
        if labels != ["ROLL AHEAD DISTANCE", "W20-1"]:
            fails.append(f"unexpected order: {labels}")

    elif n == 60:
        for tok in ("W23-1", "NYW8-30", "PVH+TMIA", "200/5", "160/4", "120/3", "PAVEMENT"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if labels != ["ROLL AHEAD DISTANCE", "W23-1"]:
            fails.append(f"unexpected order: {labels}")

    elif n == 80:
        for tok in ("W20-1", "G20-2", "G20-1", "100", "200", "350", "500", "1000"):
            if tok not in b and squash(tok) not in sb:
                fails.append(f"missing {tok}")
        if "ROLL AHEAD" in b and "ROLL AHEAD DISTANCE" in labels:
            fails.append("080 should not order roll-ahead")

    return fails


def check_ref(n: int) -> list[str]:
    fails: list[str] = []
    spec = json.loads((ROOT / f"Data/sheet-specs/619-{n:03d}.json").read_text(encoding="utf-8"))
    if spec["sheet"].get("kind") != "referenceLibrary" and n != 80:
        # 080 is plan-lite, handled in corridor
        fails.append("expected kind=referenceLibrary")
        return fails
    if n == 80:
        return check_corridor(80)

    b = body_of(n)
    sb = squash(b)
    phrases = {
        1: ["TEMPORARY", "BARRIER", "BAR"],
        4: ["WOODEN", "SIGN", "STAND", "2 X 4"],
        5: ["PTRS", "RUMBLE", "240", "W20-1"],
        6: ["PVMS", "RADAR", "SPEED", "160", "200", "240"],
        10: ["GENERAL NOTES", "MOBILE", "LONG-TERM", "SHORT-TERM"],
        12: ["SIGN TABLE", "COLOR CODE", "R2-1", "W20-1"],
    }
    for tok in phrases.get(n, []):
        if tok not in b and squash(tok) not in sb:
            fails.append(f"missing phrase {tok}")

    # Every tableRoles entry must exist with rows
    for role, tid in spec.get("tableRoles", {}).items():
        if role == "note":
            continue
        if tid not in spec["tables"] or not spec["tables"][tid].get("rows"):
            fails.append(f"role {role} -> {tid} missing/empty")

    return fails


def main() -> int:
    total = 0
    for n in CORRIDOR + [80]:
        fails = check_corridor(n)
        print(f"619-{n:03d}: ROUND-TRIP FAILURES: {len(fails)}")
        for f in fails:
            print(f"  {f}")
        total += len(fails)
    for n in REF:
        if n == 80:
            continue
        fails = check_ref(n)
        print(f"619-{n:03d}: ROUND-TRIP FAILURES: {len(fails)}")
        for f in fails:
            print(f"  {f}")
        total += len(fails)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
