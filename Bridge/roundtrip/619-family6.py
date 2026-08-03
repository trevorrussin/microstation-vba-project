"""Batch round-trip for Family 6 sheet specs against PDF text layers."""
from __future__ import annotations

import json
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import squash  # noqa: E402

SPEC311 = json.loads((ROOT / "Data/sheet-specs/619-311.json").read_text(encoding="utf-8"))
SHEETS = [307, 308, 309, 314, 321, 322, 323, 324, 407, 421, 422, 519, 524, 90, 91]


def pdf_path(n: int) -> pathlib.Path:
    return ROOT / (f"Bridge/captures/619-{n:03d}.pdf" if n < 100 else f"Bridge/captures/619-{n}.pdf")


def body(n: int) -> str:
    doc = fitz.open(str(pdf_path(n)))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def spec_path(n: int) -> pathlib.Path:
    return ROOT / (f"Data/sheet-specs/619-{n:03d}.json" if n < 100 else f"Data/sheet-specs/619-{n}.json")


def check_sheet(n: int) -> list[str]:
    fails: list[str] = []
    spec = json.loads(spec_path(n).read_text(encoding="utf-8"))
    b = body(n)
    sb = squash(b)
    roles = spec["tableRoles"]

    # Sign / size sync
    size_id = roles.get("signSizes")
    if size_id:
        size_codes = {r["signCode"] for r in spec["tables"][size_id]["rows"]}
        sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
        if size_codes != sign_codes:
            fails.append(f"sign/size mismatch {sorted(size_codes ^ sign_codes)}")

    # No merging taper on flagger / pedestrian / closure sheets
    labels = [r.get("label") or r.get("signCode")
              for al in spec["orderTable"]["alignments"] for r in al["rows"]]
    if n not in (324, 422) and ("MERGING TAPER" in labels or "LANE TAPER" in labels):
        fails.append(f"unexpected merging/lane taper in order: {labels}")

    # Buffer identity vs 311 where full 25-55 buffer-only
    buf_id = roles.get("taperAndBuffer")
    if buf_id and buf_id in spec["tables"] and n in (307, 308, 309, 314, 90, 91, 421, 323):
        for row in spec["tables"][buf_id]["rows"]:
            if "speedMph" not in row:
                continue
            r311 = next(
                (r for r in SPEC311["tables"]["311-02"]["rows"] if r["speedMph"] == row["speedMph"]),
                None,
            )
            if not r311:
                continue
            if row["longitudinalBufferSpace"] != r311["longitudinalBufferSpace"]:
                fails.append(f"{buf_id} buffer@{row['speedMph']} != 311")
            if "laneTaper" in row:
                fails.append(f"{buf_id} must not have laneTaper on flagger sheet")

    # AW identity vs 311
    aw_id = roles.get("advanceWarningSpacing")
    if aw_id and aw_id in spec["tables"] and n in (307, 308, 309, 323, 407, 421, 90, 91, 324, 422, 524):
        aw_rows = spec["tables"][aw_id]["rows"]
        ref = SPEC311["tables"]["311-03"]["rows"]
        if len(aw_rows) != len(ref):
            fails.append(f"{aw_id} row count {len(aw_rows)} != 311-03")
        else:
            for a, r in zip(aw_rows, ref):
                for k in ("A", "B", "C", "XX", "YY", "areaType"):
                    if a.get(k) != r.get(k):
                        fails.append(f"{aw_id} {k} mismatch vs 311-03")
                        break

    # Required PDF tokens by sheet family
    must = {
        307: ["FLAGGER", "W20-7", "W20-4", "W20-1", "W3-4", "G20-2", "BUFFER SPACE", "155/4", "495/13"],
        308: ["FLAGGER", "W20-7", "W20-4", "INTERSECTION", "155/4"],
        309: ["AFAD", "R10-6", "W20-7", "BUFFER"],
        314: ["FLAGGER", "W20-7", "500'", "425/11", "ROLL AHEAD"],
        321: ["SIDEWALK", "R9-9"],
        322: ["CROSSWALK", "R9-9", "CHANNELIZING"],
        323: ["FLAGGER", "W20-7", "INTERSECTION"],
        324: ["SHOULDER", "BUFFER", "W20-1"],
        407: ["FLAGGER", "360/9", "645/16", "NYR9-11"],
        421: ["FLAGGER", "NYR9-11", "155/4"],
        422: ["SHOULDER", "BUFFER", "NYR9-11"],
        519: ["SIDEWALK", "R9-9"],
        524: ["SIGNAL", "W3-3", "R10-6"],
        90: ["W20-7", "W3-4", "W20-1", "BUFFER", "155/4"],
        91: ["W20-7", "W3-4", "W20-1", "BUFFER", "155/4"],
    }
    for tok in must.get(n, []):
        if squash(tok) not in sb and tok not in b:
            fails.append(f"missing token {tok!r}")

    # 407 buffer speeds only 45-65
    if n == 407 and buf_id:
        speeds = [r["speedMph"] for r in spec["tables"][buf_id]["rows"]]
        if speeds != [45, 50, 55, 65]:
            fails.append(f"407 buffer speeds {speeds} != [45,50,55,65]")

    # Pedestrian sheets: no taperAndBuffer / roll roles
    if n in (321, 322, 519):
        if roles.get("taperAndBuffer") or roles.get("rollAheadDistance"):
            fails.append("pedestrian sheet must not declare taper/roll roles")
        if "BUFFER SPACE" in labels or "ROLL AHEAD DISTANCE" in labels:
            fails.append(f"pedestrian order must omit buffer/roll: {labels}")

    # Closure sheets: no roll
    if n in (90, 91):
        if roles.get("rollAheadDistance") or roles.get("protectiveVehicle"):
            fails.append("closure sheet must not declare PV/roll roles")
        if "ROLL AHEAD DISTANCE" in labels:
            fails.append("closure order must omit roll")

    # Page rotation recorded
    if spec["sheet"].get("pageRotation") != fitz.open(str(pdf_path(n)))[0].rotation:
        fails.append("pageRotation mismatch vs PDF")

    return fails


def main() -> int:
    total = 0
    for n in SHEETS:
        fails = check_sheet(n)
        name = f"619-{n:03d}" if n < 100 else f"619-{n}"
        print(f"{name}: ROUND-TRIP FAILURES: {len(fails)}")
        for f in fails:
            print(f"  {f}")
        total += len(fails)
    print(f"\nTOTAL FAILURES: {total}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
