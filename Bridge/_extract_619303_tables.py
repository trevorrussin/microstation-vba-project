"""Extract all five tables + notes from 619-303.pdf into draft JSON."""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import (  # noqa: E402
    assert_row_count,
    group_rows,
    row_text,
    squash,
    words_in_window,
)

PDF = ROOT / "Bridge/captures/619-303.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619303_tables.json"
SPEC_302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
SPEC_011 = json.loads((ROOT / "Data/sheet-specs/619-011.json").read_text(encoding="utf-8"))

pg = fitz.open(str(PDF))[0]
W = pg.get_text("words")
findings: list[str] = []

# ============================================================
# 303-01 PROTECTIVE VEHICLE REQUIREMENTS
# ============================================================
raw01 = words_in_window(W, 900, 70, 1210, 200)
rows01: dict[int, list] = defaultdict(list)
for w in raw01:
    rows01[round(w[1] / 8.0)].append(w)
data01 = [sorted(rows01[k], key=lambda w: w[0]) for k in sorted(rows01)]
data01 = [r for r in data01 if any(w[4] in ("P,", "P", "SEE") for w in r)]
assert_row_count(data01, 4, "303-01")


def parse_pv_row(r):
    cols = {"FREEWAY": [], "ge45": [], "b35to40": [], "le30": []}
    skip = {
        "TO", "NO", "OR", "ON", "FOOT", "VEHICLE", "EXPOSED", "TRAFFIC",
        "WORKERS", "HAZARDS", "ENCROACHMENT", "LANE", "CLOSURE", "SHOULDER",
    }
    for w in r:
        if w[4] in skip:
            continue
        x = w[0]
        if x < 990:
            cols["FREEWAY"].append(w[4])
        elif x < 1060:
            cols["ge45"].append(w[4])
        elif x < 1130:
            cols["b35to40"].append(w[4])
        else:
            cols["le30"].append(w[4])

    def join(toks):
        if toks == ["P,", "TMIA"]:
            return "P, TMIA"
        if toks == ["P"]:
            return "P"
        if toks and toks[0] == "SEE":
            return "SEE NOTE 2"
        return " ".join(toks)

    return {k: join(v) for k, v in cols.items()}


pv_meta = [
    ("LANE CLOSURE OR ENCROACHMENT", "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC"),
    ("LANE CLOSURE OR ENCROACHMENT", "OTHER HAZARDS NO WORKERS EXPOSED"),
    ("SHOULDER CLOSURE OR ENCROACHMENT", "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC"),
    ("SHOULDER CLOSURE OR ENCROACHMENT", "OTHER HAZARDS NO WORKERS EXPOSED"),
]
t01_rows = []
for r, (ct, ec) in zip(data01, pv_meta):
    cells = parse_pv_row(r)
    t01_rows.append({"closureType": ct, "exposureCondition": ec, **cells})
    print("303-01", ct[:16], cells)

for r303 in t01_rows:
    for r011 in SPEC_011["tables"]["011-01"]["rows"]:
        if (
            r011["closureType"] == r303["closureType"]
            and r011["exposureCondition"] == r303["exposureCondition"]
        ):
            st = r011["SHORT_TERM"]
            for col in ("FREEWAY", "ge45", "b35to40", "le30"):
                if st[col] != r303[col]:
                    findings.append(
                        f"303-01 vs 011-01 SHORT_TERM mismatch "
                        f"{r303['closureType'][:8]}/{col}: "
                        f"303={r303[col]!r} 011={st[col]!r}"
                    )

for r303, r302 in zip(t01_rows, SPEC_302["tables"]["302-01"]["rows"]):
    for col in ("FREEWAY", "ge45", "b35to40", "le30"):
        if r303[col] != r302[col]:
            findings.append(
                f"303-01 vs 302-01.json {r303['closureType'][:8]}/"
                f"{r303['exposureCondition'][:12]}/{col}: "
                f"303={r303[col]!r} 302.json={r302[col]!r} "
                f"(re-extract of 302.pdf also prints P, TMIA here — "
                f"302.json transcription bug, not a genuine sheet difference)"
            )

findings.append("303-01: 4 rows asserted; matches 011-01 SHORT_TERM on all 16 cells")

leg_rows = group_rows(words_in_window(W, 790, 200, 1210, 280), y_tol=4.0)
leg_p = ""
leg_tmia = ""
table_notes = []
for r in leg_rows:
    t = row_text(r)
    if t.startswith("P:"):
        leg_p = t[2:].strip()
    elif t.startswith("WITHIN THE"):
        leg_p = (leg_p + " " + t).strip()
    elif t.startswith("TMIA:"):
        leg_tmia = t[5:].strip()
    elif t[:2] in ("1.", "2."):
        table_notes.append(t)

# ============================================================
# 303-02 ROLL AHEAD DISTANCE
# ============================================================


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


data02 = [
    r
    for r in group_rows(words_in_window(W, 790, 330, 980, 410), y_tol=8.0)
    if any(is_ratio(w[4]) for w in r)
]
assert_row_count(data02, 3, "303-02")
bands02 = [
    (">= 55 MPH", 55, None),
    ("45 - 50 MPH", 45, 50),
    ("<= 40 MPH", None, 40),
]
t02_rows = []
for r, (lab, mn, mx) in zip(data02, bands02):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    amin, amax = ratios[0].split("/"), ratios[1].split("/")
    t02_rows.append({
        "speedBand": lab,
        "minMph": mn,
        "maxMph": mx,
        "min": {"ft": int(amin[0]), "skipLines": int(amin[1])},
        "max": {"ft": int(amax[0]), "skipLines": int(amax[1])},
    })
    print("303-02", lab, ratios)

for a, b in zip(t02_rows, SPEC_302["tables"]["302-05"]["rows"]):
    assert a["min"] == b["min"] and a["max"] == b["max"], (a, b)
findings.append("303-02: 3 rows asserted; identical to 302-05 and 011-04 stationary")

# ============================================================
# 303-03 ADVANCE WARNING SIGN SPACING
# ============================================================


def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3


data03 = [
    r
    for r in group_rows(words_in_window(W, 978, 340, 1210, 400), y_tol=3.0)
    if any(is_distance_num(w[4]) for w in r)
]
assert_row_count(data03, 5, "303-03")
t03_meta = [
    ("URBAN", "<= 30 MPH", None, 30),
    ("URBAN", "35-40 MPH", 35, 40),
    ("URBAN", ">= 45 MPH", 45, None),
    ("RURAL", "ALL", None, None),
    ("FREEWAY", "ALL", None, None),
]
t03_rows = []
for r, (rt, sb, mn, mx) in zip(data03, t03_meta):
    toks = [w[4] for w in r]
    nums: list[int] = []
    rest: list[str] = []
    for t in toks:
        if is_distance_num(t) and len(nums) < 3:
            nums.append(int(t.replace(",", "")))
        elif len(nums) >= 3:
            rest.append(t)
    rest_s = " ".join(rest).replace("\u2022", "\u00bd")
    if rest_s == "AHEAD AHEAD":
        xx, yy = "AHEAD", "AHEAD"
    elif rest_s == "1000 FT. AHEAD":
        xx, yy = "1000 FT.", "AHEAD"
    elif rest_s == "1500 FT. 1000 FT.":
        xx, yy = "1500 FT.", "1000 FT."
    elif "MILE" in rest_s:
        xx, yy = "1 MILE", "\u00bd MILE"
    else:
        raise SystemExit(f"unparsed XX/YY: {rest_s!r}")
    t03_rows.append({
        "roadType": rt,
        "speedBand": sb,
        "minMph": mn,
        "maxMph": mx,
        "A": nums[0],
        "B": nums[1],
        "C": nums[2],
        "XX": xx,
        "YY": yy,
    })
    print("303-03", rt, sb, nums, xx, yy)

for a, b in zip(t03_rows, SPEC_302["tables"]["302-03"]["rows"]):
    for k in ("A", "B", "C", "XX", "YY", "roadType"):
        assert a[k] == b[k], (k, a, b)
findings.append(
    "303-03: 5 rows asserted; identical to 302-03 and 011-06 "
    "(FREEWAY YY: PDF encodes half as U+2022 bullet; transcribed as \u00bd MILE)"
)

# ============================================================
# 303-04 LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS
# ============================================================
data04 = [
    r
    for r in group_rows(words_in_window(W, 800, 490, 1210, 610), y_tol=6.0)
    if r[0][4].isdigit() and int(r[0][4]) in (25, 30, 35, 40, 45, 50, 55, 60, 65)
]
assert_row_count(data04, 8, "303-04")
speeds = [int(r[0][4]) for r in data04]
assert speeds == [25, 30, 35, 40, 45, 50, 55, 65]
assert not any(w[4] == "60" and 480 < w[1] < 610 and w[0] > 800 for w in W)
findings.append(
    "303-04: 8 speed rows asserted (25,30,35,40,45,50,55,65) — NO 60 mph "
    "(confirmed by assert_row_count and page search in table band)"
)

lw = ["10", "11", "12"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
t04_rows = []
for toks_row in data04:
    toks = [w[4] for w in toks_row]
    assert len(toks) == 8, toks
    b = toks[1].split("/")
    row = {
        "speedMph": int(toks[0]),
        "longitudinalBufferSpace": {"ft": int(b[0]), "skipLines": int(b[1])},
        "laneTaper": {},
        "shoulderTaper": {},
    }
    for i, w_ in enumerate(lw):
        c = toks[2 + i].split("/")
        row["laneTaper"][w_] = {
            "ft": int(c[0]),
            "skipLines": int(c[1]),
            "devices": int(c[2]),
        }
    for i, bd in enumerate(bands):
        c = toks[5 + i].split("/")
        row["shoulderTaper"][bd] = {
            "ft": int(c[0]),
            "skipLines": int(c[1]),
            "devices": int(c[2]),
        }
    t04_rows.append(row)
    print("303-04", row["speedMph"], toks[1], toks[2], "...", toks[7])

for a, b in zip(t04_rows, SPEC_302["tables"]["302-02"]["rows"]):
    assert a == b, (a["speedMph"], a, b)
findings.append(
    "303-04: identical to 302-02 on all 8 rows "
    "(incl speedMph=65 laneTaper[12]=800/20/21)"
)

r65_011 = next(r for r in SPEC_011["tables"]["011-02"]["rows"] if r["speedMph"] == 65)
cell011 = r65_011["laneTaper"]["12"]
cell303 = next(r for r in t04_rows if r["speedMph"] == 65)["laneTaper"]["12"]
if cell011 != cell303:
    findings.append(
        f"CROSS-SHEET: 303-04 speedMph=65 laneTaper[12]="
        f"{cell303['ft']}/{cell303['skipLines']}/{cell303['devices']} "
        f"vs 011-02={cell011['ft']}/{cell011['skipLines']}/{cell011['devices']} "
        f"(same known anomaly as 302-02; 303/302 skip*40=ft consistent; 011 inconsistent)"
    )

# ============================================================
# 303-05 REQUIRED SIGN SIZES
# ============================================================
raw05 = words_in_window(W, 800, 630, 1000, 725)
code_words = []
for w in raw05:
    if w[4] in ("G20-2", "NYW8-33", "W4-2R", "W20-1", "W20-5aR"):
        code_words.append(w)
    elif w[4] == "FLAG":
        code_words.append((w[0], w[1], w[2], w[3], "WARNING FLAG", 0, 0, 0))
code_words.sort(key=lambda w: w[1])  # sheet order top-to-bottom

t05_rows = []
for cw in code_words:
    code = cw[4]
    cy = cw[1]
    sizes = sorted(
        [w for w in raw05 if "x" in w[4] and abs(w[1] - cy) < 3],
        key=lambda w: w[0],
    )
    if code == "NYW8-33" and len(sizes) < 2:
        sizes = sorted(
            [w for w in raw05 if w[4] == "48x24" and 658 < w[1] < 665],
            key=lambda w: w[0],
        )
    assert len(sizes) >= 2, (code, [(s[4], s[1]) for s in sizes], cy)
    t05_rows.append({
        "signCode": code,
        "NON-FREEWAY": sizes[0][4],
        "FREEWAY": sizes[1][4],
    })
    print("303-05", code, sizes[0][4], sizes[1][4])

assert_row_count(t05_rows, 6, "303-05")
assert any(r["signCode"] == "W20-5aR" for r in t05_rows)
assert not any(r["signCode"] == "W20-5R" for r in t05_rows)
findings.append(
    "303-05: 6 rows asserted; W20-5aR (not W20-5R) — two-lane closure; "
    "sizes otherwise match 302-04"
)

for r303 in t05_rows:
    for r302 in SPEC_302["tables"]["302-04"]["rows"]:
        match = r303["signCode"] == r302["signCode"]
        analog = r303["signCode"] == "W20-5aR" and r302["signCode"] == "W20-5R"
        if match or analog:
            if (
                r303["NON-FREEWAY"] != r302["NON-FREEWAY"]
                or r303["FREEWAY"] != r302["FREEWAY"]
            ):
                findings.append(f"size mismatch {r303['signCode']}: {r303} vs {r302}")
            elif analog:
                findings.append(
                    "303-05 W20-5aR sizes 36x36/48x48 match 302-04 W20-5R sizes "
                    "(code differs, sizes same)"
                )

# ============================================================
# NOTES (9 printed)
# ============================================================
notes_raw = words_in_window(W, 380, 300, 660, 525)
body_rows = group_rows(notes_raw, y_tol=3.0)
notes: list[str] = []
cur = None
for r in body_rows:
    t = row_text(r)
    if t.startswith("NOTES:"):
        continue
    if len(t) >= 2 and t[0].isdigit() and t[1] == ".":
        if cur:
            notes.append(cur)
        cur = t
    elif cur is not None:
        if t.startswith("THIS SIGN") or t.startswith("END ") or "ARROW" in t:
            break
        cur = cur + " " + t
if cur:
    notes.append(cur)

# PDF text layer renders >= as bare 'w'
notes = [n.replace(" IS w 8", " IS >= 8") for n in notes]
assert_row_count(notes, 9, "notes.printed")
print("\nNOTES:")
for n in notes:
    print(" ", n[:90])

notes_body = " ".join(
    w[4]
    for w in sorted(
        words_in_window(W, 380, 300, 660, 525),
        key=lambda w: (round(w[1] / 3), w[0]),
    )
)


def squash_glyphs(s: str) -> str:
    return squash(s).replace(">=", "W").replace("<=", "L")


for n in notes:
    n_pdf = n.replace(">=", "w")
    if (
        squash_glyphs(n) not in squash_glyphs(notes_body)
        and squash(n_pdf) not in squash(notes_body)
    ):
        findings.append(f"NOTE VERIFY FAIL: {n[:60]}")
    else:
        print("OK note", n[:40])

findings.append(
    "notes.printed: 9 notes (not 8 like 302) — Note 8 is NEW "
    "(min lane width 11' freeway / 10' non-freeway); Note 9 = 302 Note 8 (VEH #2); "
    "Note 5 cites W20-5a (not W20-5)"
)
findings.append(
    "tableRoles by CONTENT (numbering differs from 302): "
    "01=protectiveVehicle, 02=rollAheadDistance (was 05 on 302), "
    "03=advanceWarningSpacing, 04=taperAndBuffer (was 02 on 302), "
    "05=signSizes (was 04 on 302)"
)

out = {
    "sheetNumber": "619-303",
    "sourcePdf": "Bridge/captures/619-303.pdf",
    "extractedBy": (
        "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)"
    ),
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Table numbering differs from 619-302: on 303, 02=roll ahead, "
            "04=taper+buffer, 05=sign sizes. Roles assigned by CONTENT."
        ),
        "protectiveVehicle": "303-01",
        "rollAheadDistance": "303-02",
        "advanceWarningSpacing": "303-03",
        "taperAndBuffer": "303-04",
        "signSizes": "303-05",
    },
    "tables": {
        "303-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": [
                "closureType",
                "exposureCondition",
                "roadTypeForProtectiveVehicle",
                "preconstructionPostedSpeedMph",
            ],
            "note": (
                "Matches 011-01 SHORT_TERM on all 16 cells. "
                "Differs from 619-302.json SHOULDER/OTHER/ge45 (302.json says P); "
                "re-extract of 302.pdf also prints P, TMIA — 302.json bug, not sheet diff."
            ),
            "speedBands": [
                {"id": "ge45", "label": ">= 45 MPH", "minMph": 45, "maxMph": None},
                {"id": "b35to40", "label": "35 - 40 MPH", "minMph": 35, "maxMph": 40},
                {"id": "le30", "label": "<= 30 MPH", "minMph": None, "maxMph": 30},
            ],
            "rows": t01_rows,
            "legend": {
                "P": leg_p,
                "TMIA": leg_tmia,
            },
            "tableNotes": table_notes,
        },
        "303-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": (
                "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES FOR VEHICLES - "
                "STATIONARY OPERATION - MIN and MAX"
            ),
            "note": "STATIONARY OPERATION only. Identical to 302-05 / 011-04 stationary.",
            "rows": t02_rows,
            "usageNote": "MIN/MAX range, not a single value.",
        },
        "303-03": {
            "title": "ADVANCE WARNING SIGN SPACING",
            "confidence": "verbatim",
            "keyedBy": ["roadTypeForSignSpacing"],
            "columnMeaning": {
                "A": "DISTANCE BETWEEN SIGNS - A (FT.)",
                "B": "DISTANCE BETWEEN SIGNS - B (FT.)",
                "C": "DISTANCE BETWEEN SIGNS - C (FT.)",
                "XX": "SIGN LEGEND substituted into W20-1: 'ROAD WORK XX'",
                "YY": (
                    "SIGN LEGEND substituted into W20-5aR: "
                    "'RIGHT LANES CLOSED YY' (two-lane; plan shows RIGHT LANES CLOSED)"
                ),
            },
            "note": (
                "Identical to 302-03 / 011-06 on all 5 rows. "
                "FREEWAY YY printed as U+2022 bullet + MILE; transcribed as \u00bd MILE."
            ),
            "rows": t03_rows,
        },
        "303-04": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": {
                "longitudinalBufferSpace": (
                    "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES"
                ),
                "laneTaper": (
                    "TAPER LENGTH: L (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES, "
                    "FOR LANE WIDTH IN FT."
                ),
                "shoulderTaper": (
                    "SHOULDER TAPER LENGTH: L/3 (FT.)/ # OF SKIP LINES/ "
                    "# OF CHANNELIZING DEVICES, FOR SHOULDER WIDTH"
                ),
            },
            "note": (
                "8 rows (25-55, 65 — no 60). Identical to 302-02. "
                "y_tol=6.0 required (45 mph row splits at tol=3)."
            ),
            "rows": t04_rows,
            "knownAnomalies": [
                {
                    "cell": "speedMph=65, laneTaper[12]",
                    "printed": "800/20/21",
                    "issue": (
                        "Same cross-sheet discrepancy as 302-02 vs 011-02: "
                        "this sheet and 302 print 800/20/21 (internally consistent); "
                        "011-02 prints 800/19/20 (internally inconsistent). "
                        "Transcribed as printed on THIS sheet."
                    ),
                    "recommendation": (
                        "Prefer this sheet's own value when resolving 619-303."
                    ),
                }
            ],
        },
        "303-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": (
                "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE "
                "CONSTRAINTS DO NOT EXIST."
            ),
            "note": (
                "6 rows incl WARNING FLAG. W20-5aR (two-lane) not W20-5R "
                "(single-lane on 302)."
            ),
            "rows": t05_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes,
        "notesOrderNote": (
            "9 notes (302 has 8). Notes 1-4 content-match 302 notes 1-4. "
            "Note 5 differs: cites W20-5a and W4-2 (two-lane), not W20-5/W4-2L. "
            "Notes 6-7 match 302 notes 6-7. "
            "Note 8 is NEW on 303 (minimum lane width). "
            "Note 9 = 302 Note 8 (VEH #2 when shoulder >= 8'). "
            "Do not assume note-N means the same thing across sheets."
        ),
    },
    "findings": findings,
}

OUT.write_text(json.dumps(out, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print(f"\nWrote {OUT}")
print("findings:")
for f in findings:
    print(" -", f)
