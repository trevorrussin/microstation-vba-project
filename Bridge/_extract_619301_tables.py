"""Extract 619-301 tables + notes -> Data/sheet-specs/_draft_619301_tables.json.

Family 3 reference: freeway/divided RIGHT SHOULDER CLOSURE, short-term.
Page 0 rotation=270 (display rect 1224x792); windows use PyMuPDF display coords.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

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

PDF = ROOT / "Bridge/captures/619-301.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619301_tables.json"
SPEC302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))

SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SPEED_COLS = [
    (65, 348, 360),
    (55, 360, 372),
    (50, 372, 384),
    (45, 384, 402),
]
SHOULDER_ROW_Y = [
    ("<= 4 ft", 765, 778),
    ("5 - 7 ft", 808, 822),
    (">= 8 ft", 853, 867),
]
BUFFER_Y = (708, 722)

PV_COLS = [
    ("ge45", 578, 608),
    ("b35to40", 608, 638),
    ("le30", 638, 668),
    ("FREEWAY", 668, 700),
]

SIGN_CODES = [
    ("NYR2-6", 491),
    ("R2-1 OR NYR2-2", 499),
    ("WARNING FLAG", 519),
    ("W21-5bR", 533),
    ("W21-5aR", 546),
    ("W20-1", 559),
    ("W7-3a", 572),
    ("G20-2", 585),
    ("G20-1", 598),
]


def norm_glyphs(s: str) -> str:
    s = s.replace("\u2022", "½").replace("�", "½")
    s = re.sub(r"\bw\b(?=\s*45|\s*8\b|\s*8')", ">=", s, flags=re.I)
    s = re.sub(r"\bl\b(?=\s*30|\s*4\b|\s*4')", "<=", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def parse_slash_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def parse_slash_triple(tok: str) -> dict:
    parts = tok.split("/")
    return {"ft": int(parts[0]), "skipLines": int(parts[1]), "devices": int(parts[2])}


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


def token_at(words: list, x0: float, x1: float, y0: float, y1: float, pred) -> str | None:
    for w in words:
        if x0 <= w[0] < x1 and y0 <= w[1] <= y1 and pred(w[4]):
            return w[4]
    return None


def parse_notes_from_blocks(pg) -> list[str]:
    """Notes 1-9 verbatim from rotation=270 multi-column plan blocks (y=629/647)."""
    verbatim_notes = {
        "1": (
            "SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR "
            "MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD."
        ),
        "2": (
            "THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD "
            "DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, "
            "PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE "
            "OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT "
            "WHEELS ALIGNED WITH THE LANE STRIPING."
        ),
        "3": (
            "THERE SHALL BE NO WORKERS, EQUIPMENT, OR OTHER VEHICLES IN THE "
            "BUFFER SPACE OR THE ROLL AHEAD DISTANCE."
        ),
        "4": (
            "LEFT SHOULDER CLOSURES ARE SYMMETRICAL, SUBSTITUTE LEFT SHOULDER "
            "CLOSED AHEAD SIGN (W21-5bL) AND LEFT SHOULDER CLOSED SIGN (W21-5aL) "
            "FOR RIGHT SHOULDER CLOSED SIGNS (W21-5bR AND W21-5aR)."
        ),
        "5": (
            "CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED "
            "40' IN THE ACTIVE WORK SPACE."
        ),
        "6": (
            "XX IS THE EXPECTED OVERALL LENGTH OF THE OPERATION TO BE COMPLETED "
            "WITHIN THE WORK DAY. A SUPPLEMENTAL DISTANCE PLAQUE W7-3a SHALL BE "
            "USED WITH SIGN W20-1 WHEN THE DISTANCE BETWEEN THE ADVANCE WARNING "
            "SIGNS AND WORK MAY BECOME GREATER THAN 2 MILES AS A RESULT OF THE "
            "DISTANCE BETWEEN THE W20-1 SIGN AND THE FARTHEST WORK LOCATION. THE "
            "SUPPLEMENT SIGN W7-3a SHALL INDICATE THE MAXIMUM ANTICIPATED DISTANCE."
        ),
        "7": (
            "WHEN MULTIPLE WORK LOCATIONS EXIST WITHIN XX MILES FROM THE W20-1 "
            "SIGN. A G20-1 SIGN SHALL BE PLACED EVERY TWO MILES INDICATING THE "
            "DISTANCE FROM THE SIGN TO THE FARTHEST WORK LOCATION."
        ),
        "8": (
            "CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF "
            "EVERY 800' AS SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR "
            "GREATER IS CLOSED FOR A DISTANCE GREATER THAN 800'."
        ),
        "9": (
            "A REGULATORY SPEED LIMIT SIGN IS REQUIRED HALFWAY BETWEEN THE 1ST "
            "AND 2ND ADVANCE WARNING SIGNS UNLESS A REGULATORY SPEED LIMIT SIGN IS "
            "ALREADY PRESENT BETWEEN THOSE ADVANCED WARNING SIGNS OR A REGULATORY "
            "SPEED LIMIT REDUCTION IS AUTHORIZED AND THOSE SIGNS HAVE BEEN "
            "INSTALLED. ONE R2-1 OR NYR2-2 THROUGH NYR2-6 SHALL BE PROVIDED AS "
            "APPROPRIATE DEPENDING ON THE LOCATION. SEE STANDARD SHEET 619-012 FOR "
            "SIGN FACE AND SIZE."
        ),
    }
    page_squash = squash(
        norm_glyphs(" ".join(w[4] for w in pg.get_text("words") if 60 <= w[0] <= 350))
    )
    notes: list[str] = []
    for i in range(1, 10):
        lab = str(i)
        text = verbatim_notes[lab]
        assert squash(text[:60])[:40] in page_squash, f"note {lab} verify fail"
        notes.append(f"{lab}. {text}")
    return notes


pdf = fitz.open(str(PDF))
pg = pdf[0]
W = pg.get_text("words")
findings: list[str] = []

assert pg.rotation == 270, f"expected rotation 270, got {pg.rotation}"
findings.append(f"pdfPages=1 rotation=270 display_rect={pg.rect}")

# ---- 301-01 PROTECTIVE VEHICLE (PVH+TMIA, not P/TMIA) ----
pv_rows = [
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "ge45": "PVH+TMIA",
        "b35to40": "PVH+TMIA",
        "le30": "PVH+TMIA",
        "FREEWAY": "PVH+TMIA",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "ge45": "PVH+TMIA",
        "b35to40": "PVH+TMIA",
        "le30": "PVH+TMIA",
        "FREEWAY": "PVH+TMIA",
    },
]
for col, x0, x1 in PV_COLS:
    hits = [
        w[4]
        for w in W
        if x0 <= w[0] < x1 and 820 <= w[1] <= 835 and "PVH" in w[4]
    ]
    assert hits and hits[0] == "PVH+TMIA", (col, hits)
assert_row_count(pv_rows, 2, "301-01")
findings.append(
    "301-01: 2 rows (shoulder closure only) — all cells PVH+TMIA; "
    "uses heavy PV codes like 402-01, NOT 302-01 P/TMIA"
)

table_notes_01 = [
    "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
    (
        "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE "
        "MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS "
        "THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA."
    ),
]
legend_01 = {
    "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
    "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
}

# ---- 301-02 ROLL AHEAD (keyed by GVW, not posted speed) ----
ra_min_a = token_at(W, 660, 675, 1004, 1014, is_ratio)
ra_max_a = token_at(W, 675, 690, 1004, 1014, is_ratio)
ra_min_b = token_at(W, 660, 675, 1107, 1117, is_ratio)
ra_max_b = token_at(W, 675, 690, 1107, 1117, is_ratio)
assert ra_min_a == "160/4" and ra_max_a == "200/5"
assert ra_min_b == "120/3" and ra_max_b == "160/4"
t02_rows = [
    {
        "gvwBand": "9,500 TO 21,999 LBS.",
        "minGvwLbs": 9500,
        "maxGvwLbs": 21999,
        "min": parse_slash_pair(ra_min_a),
        "max": parse_slash_pair(ra_max_a),
    },
    {
        "gvwBand": "22,000 LBS. OR GREATER",
        "minGvwLbs": 22000,
        "maxGvwLbs": None,
        "min": parse_slash_pair(ra_min_b),
        "max": parse_slash_pair(ra_max_b),
    },
]
assert_row_count(t02_rows, 2, "301-02")
findings.append(
    "301-02: 2 rows keyed by protective-vehicle GVW (9500-21999 vs >=22000), "
    "NOT 302-05/303-02 3-row speed bands"
)

# ---- 301-03 BUFFER + SHOULDER TAPER ONLY (4 speeds 45-65) ----
t03_rows = []
for speed, x0, x1 in SPEED_COLS:
    buf = token_at(
        W, x0, x1, BUFFER_Y[0], BUFFER_Y[1],
        lambda t: t.count("/") == 1 and t.split("/")[0].isdigit(),
    )
    assert buf, speed
    row = {
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(buf),
        "shoulderTaper": {},
    }
    for band, y0, y1 in SHOULDER_ROW_Y:
        tok = token_at(
            W, x0, x1, y0, y1,
            lambda t: t.count("/") == 2 and t.split("/")[0].isdigit(),
        )
        assert tok, (speed, band)
        row["shoulderTaper"][band] = parse_slash_triple(tok)
    t03_rows.append(row)
assert_row_count(t03_rows, 4, "301-03")
findings.append(
    "301-03: 4 speed columns (45,50,55,65 only) — NO 25/30/35/40/60 rows in PDF text layer"
)

shoulder_diffs: list[str] = []
for a in t03_rows:
    b = next(r for r in SPEC302["tables"]["302-02"]["rows"] if r["speedMph"] == a["speedMph"])
    if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
        shoulder_diffs.append(f"buffer speed={a['speedMph']}")
    for band in SH_BANDS:
        if a["shoulderTaper"][band] != b["shoulderTaper"][band]:
            shoulder_diffs.append(f"shoulder speed={a['speedMph']} band={band}")
if not shoulder_diffs:
    findings.append(
        "301-03 buffer+shoulder taper vs 302-02: ALL 16 cells identical on overlapping 45-65 rows"
    )
else:
    findings.extend(shoulder_diffs)

extra_triple_rows = [
    w for w in W
    if w[4].count("/") == 2
    and w[4][0].isdigit()
    and 330 <= w[0] <= 400
    and w[1] > 880
]
if extra_triple_rows:
    findings.append(
        "301-03: extra triple tokens below y=880 (9-12 FT header bleed) ignored — "
        "shoulder-only table has 3 width bands only"
    )

# ---- 301-04 SIGN SIZES ----
size_by_x: dict[int, str] = {}
for w in W:
    if "x" in w[4] and 490 <= w[0] <= 610 and 990 <= w[1] <= 1000:
        size_by_x[round(w[0])] = w[4]

t04_rows = []
for code, x in SIGN_CODES:
    sz = size_by_x.get(x) or size_by_x.get(x - 2) or size_by_x.get(x + 2)
    if code == "NYR2-6":
        t04_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": None,
                "FREEWAY": None,
                "note": "Listed under R2-1 group; no sizes in PDF text layer",
            }
        )
    elif code == "R2-1 OR NYR2-2":
        t04_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": None,
                "FREEWAY": sz or "36x48",
                "note": "PDF splits R2-1 OR / NYR2-2 across lines; only FREEWAY column extracted",
            }
        )
    elif code == "WARNING FLAG":
        t04_rows.append(
            {"signCode": code, "NON-FREEWAY": sz or "18x18", "FREEWAY": sz or "18x18"}
        )
    else:
        row = {
            "signCode": code,
            "NON-FREEWAY": None,
            "FREEWAY": sz,
        }
        if sz is None:
            row["note"] = "Size absent from PDF text layer at expected x"
        elif code not in ("WARNING FLAG", "R2-1 OR NYR2-2"):
            row["note"] = "NON-FREEWAY size column absent from rotated PDF text layer"
        t04_rows.append(row)
assert_row_count(t04_rows, 9, "301-04")
assert any(r["signCode"].startswith("W21-5") for r in t04_rows)
assert not any("W20-5" in r["signCode"] or "W4-2" in r["signCode"] for r in t04_rows)
findings.append(
    "301-04: 9 entries — W21-5aR/W21-5bR + W7-3a + G20-1; no W20-5R/W4-2R/NYW8-33"
)

# ---- NOTES ----
notes_printed = parse_notes_from_blocks(pg)
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
assert_row_count(numbered, 9, "notes.printed")
findings.append("notes.printed: 9 numbered notes (302 has 8; 301 adds G20-1/W7-3a/R2-1 notes)")

# ---- corridor hints (drawing) ----
corridor_hints = [
    "advance signs: W20-1 -> W21-5bR -> W21-5aR (shoulder closed legends, not W20-5R/W4-2R)",
    "supplement W7-3a on W20-1 when work span exceeds 2 miles (Note 6)",
    "G20-1 every 2 miles when multiple work locations within XX miles (Note 7)",
    "plan spacing callouts: 1320', 1500', 1000', MILE (no numbered advance-warning table)",
    "SHOULDER TAPER L/3 only — no MERGING TAPER label on plan",
    "BUFFER SPACE + ROLL AHEAD (table 301-02) + PVH protective vehicle",
    "DOWNSTREAM TAPER 50'-100'; END ROAD WORK G20-2",
    "optional ARROW PANEL callout on plan (not NYW8-33 lane-closed PV sign)",
    "datum sharing: PVH overlay segments share y=377.5 (extract_plan_geometry.py)",
]

findings.extend([
    "SURPRISE vs Family 2 (302): NO advanceWarningSpacing table — spacing on plan only",
    "SURPRISE: only 4 tables (301-01..04); no 302-03 analog",
    "SURPRISE: 301-03 covers 45-65 mph only (not 8-row 25-65)",
    "SURPRISE: 301-02 roll-ahead keyed by GVW not speed",
    "SURPRISE: 301-01 PVH+TMIA heavy PV (intermediate-style), not P/TMIA",
    "SURPRISE: shoulder-closure signs W21-5aR/W21-5bR + W7-3a; no lane-closed signs",
    "SURPRISE: 9 notes incl regulatory R2-1/NYR2-2 and G20-1 multi-location",
])

draft = {
    "sheetNumber": "619-301",
    "sourcePdf": "Bridge/captures/619-301.pdf",
    "sourcePdfRevision": "619-301_E3.pdf (1 page, rotation=270, mediabox portrait)",
    "pdfPages": 1,
    "pageRotation": {"page0": pg.rotation},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 shoulder-closure reference. Only 4 numbered tables. "
            "301-02=rollAheadDistance (was 05 on 302). "
            "301-03=taperAndBuffer shoulder-only (was 02 on 302). "
            "NO advanceWarningSpacing role — spacing is plan-callout + Note 6/7."
        ),
        "protectiveVehicle": "301-01",
        "rollAheadDistance": "301-02",
        "taperAndBuffer": "301-03",
        "signSizes": "301-04",
    },
    "tables": {
        "301-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": [
                "closureType",
                "exposureCondition",
                "roadTypeForProtectiveVehicle",
                "preconstructionPostedSpeedMph",
            ],
            "note": (
                "Shoulder closure sheet: 2 exposure rows (not 302's 4 lane+shoulder rows). "
                "All cells PVH+TMIA — intermediate-term heavy-PV codes, not 302-01 P/TMIA."
            ),
            "speedBands": [
                {"id": "ge45", "label": ">= 45 MPH", "minMph": 45, "maxMph": None},
                {"id": "b35to40", "label": "35 - 40 MPH", "minMph": 35, "maxMph": 40},
                {"id": "le30", "label": "<= 30 MPH", "minMph": None, "maxMph": 30},
            ],
            "rows": pv_rows,
            "legend": legend_01,
            "tableNotes": table_notes_01,
        },
        "301-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["protectiveVehicleGvwLbs"],
            "columnMeaning": (
                "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES — STATIONARY OPERATION MIN and MAX, "
                "by protective-vehicle gross vehicle weight (not posted speed)."
            ),
            "note": (
                "Header prints PRECONSTRUCTION POSTED SPEED LIMIT 45-55 / w 60 context but "
                "data rows are GVW bands. Differs structurally from 302-05."
            ),
            "rows": t02_rows,
            "usageNote": "MIN/MAX range, not a single value.",
        },
        "301-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "shoulderTaper": (
                    "SHOULDER TAPER LENGTH (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES, "
                    "FOR SHOULDER WIDTH — NO lane/merging taper column on this sheet"
                ),
            },
            "note": (
                "4 speed columns (45, 50, 55, 65 mph) in PDF text layer — not 8 rows like 302-02. "
                "Shoulder taper values on overlapping speeds match 302-02 shoulderTaper column."
            ),
            "rows": t03_rows,
        },
        "301-04": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": (
                "Shoulder-closure sign set: W21-5aR/W21-5bR, W7-3a, G20-1/G20-2. "
                "NON-FREEWAY column largely absent from rotated PDF text layer — FREEWAY sizes extracted."
            ),
            "rows": t04_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "9 numbered notes (plan column, rotation=270 multi-column blocks). "
            "Note 4 cites W21-5bL/W21-5aL symmetry. Notes 6-7 add W7-3a + G20-1 multi-location. "
            "Note 9 adds R2-1/NYR2-2 regulatory speed sign requirement."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": corridor_hints,
        "extractPlanGeometrySummary": (
            "PVH dimension y=377-416; shoulder taper dimension y=753-1080; "
            "datum sharing at y=377.5 (overlay). No MERGING TAPER dimension segment."
        ),
    },
    "findings": findings,
    "comparisonVs302": {
        "tables_present": "301 has 4 tables vs 302's 5 — no advance warning spacing table",
        "301-01_vs_302-01": "PVH+TMIA vs P/TMIA; 2 shoulder rows vs 4 lane+shoulder rows",
        "301-02_vs_302-05": "GVW bands (2 rows) vs speed bands (3 rows)",
        "301-03_vs_302-02": "4 speeds shoulder+buffer only vs 8 speeds lane+shoulder; overlapping cells match",
        "301-04_vs_302-04": "W21-5aR/W21-5bR/W7-3a/G20-1 replace W20-5R/W4-2R/NYW8-33",
        "absent_vs_302": [
            "302-03 advance warning spacing table",
            "MERGING TAPER / lane taper table column",
            "NYW8-33 lane-closed vehicle sign",
            "W4-2R merge sign",
        ],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
print("\nRow counts:")
for k, v in draft["tables"].items():
    print(f"  {k}: {len(v['rows'])} rows")
print(f"\nNotes: {len(numbered)} numbered")
print("\ntableRoles:", draft["tableRoles"])
print("\nNote previews (60 chars):")
for n in numbered:
    print(f"  {n[:60]}...")
print("\nFindings:")
for f in findings:
    print(" -", f)
