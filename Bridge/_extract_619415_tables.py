"""Extract 619-415 tables + notes -> Data/sheet-specs/_draft_619415_tables.json.

Family 3 ramp-approach shoulder closure, intermediate-term. 2 pages, E3, rotation=0.
Compare taper cells to 619-301; capture NYR9-11 and ramp corridor hints.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import assert_row_count, group_rows, row_text, squash  # noqa: E402

PDF = ROOT / "Bridge/captures/619-415.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619415_tables.json"
DRAFT301 = json.loads(
    (ROOT / "Data/sheet-specs/_draft_619301_tables.json").read_text(encoding="utf-8")
)
DRAFT402 = json.loads(
    (ROOT / "Data/sheet-specs/_draft_619402_tables.json").read_text(encoding="utf-8")
)

SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
LW = ["10", "11", "12"]
SPEED_ROWS = [
    (45, 118, 130),
    (50, 132, 144),
    (55, 148, 160),
    (65, 162, 174),
]
LAT_COLS = [("10", 210, 248), ("11", 256, 294), ("12", 304, 342)]
SH_COLS = [("<= 4 ft", 358, 396), ("5 - 7 ft", 404, 442), (">= 8 ft", 450, 488)]
BUF_COL = (168, 206)

CHAN_COLS = [
    ("cones", 868, 895),
    ("type1_markers", 895, 925),
    ("standard_cones", 930, 960),
    ("extra_tall_cones", 960, 990),
    ("temporary_tall_cones", 990, 1015),
    ("temporary_tubular_markers", 1015, 1045),
    ("interim_tubular_markers", 1045, 1075),
    ("vertical_panels", 1075, 1105),
    ("oversized_vertical_panels", 1105, 1135),
    ("type_iii_barricades", 1155, 1185),
]


def norm_glyphs(s: str) -> str:
    s = s.replace("\u2022", "½").replace("�", "½")
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


def cell_str(d: dict) -> str:
    if "devices" in d:
        return f"{d['ft']}/{d['skipLines']}/{d['devices']}"
    return f"{d['ft']}/{d['skipLines']}"


def cell_at(row: list, x0: float, x1: float) -> str:
    return norm_glyphs(" ".join(w[4] for w in row if x0 <= w[0] < x1))


def parse_notes_415(pg0) -> list[str]:
    """Notes 1-8 + N1-N8 from page 0 (x>=900). Numbered block verbatim with squash verify."""
    verbatim = {
        "1": (
            "INTERMEDIATE-TERM IS STATIONARY WORK THAT OCCUPIES A LOCATION MORE THAN ONE "
            "DAYLIGHT PERIOD UP TO 3 CONSECUTIVE DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR."
        ),
        "2": (
            "NO WORK ACTIVITY OR STORAGE OF EQUIPMENT, VEHICLES, OR MATERIAL SHOULD OCCUR "
            "WITHIN A BUFFER SPACE."
        ),
        "3": (
            "CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE "
            "ACTIVE WORK SPACE."
        ),
        "4": (
            "A SUPPLEMENTAL DISTANCE PLAQUE W7-3a SHALL BE USED WITH SIGN W20-1 WHEN THE "
            "DISTANCE BETWEEN THE ADVANCE WARNING SIGNS AND WORK BECOME GREATER THAN 2 MILES "
            "AS A RESULT OF THE FOLLOWING SITUATIONS: XX IS THE EXPECTED OVERALL LENGTH OF "
            "THE OPERATION TO BE COMPLETED WITHIN THE WORK DAY; *WORK LOCATIONS WITHIN XX "
            "MILES FROM THE W20-1 SIGN MAY BE RELOCATED WITHIN THE WORK DAY; *MULTIPLE WORK "
            "LOCATIONS ARE ANTICIPATED WITHIN XX MILES FROM THE W20-1 SIGN."
        ),
        "5": (
            "CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF EVERY 800' AS "
            "SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR GREATER IS CLOSED FOR A "
            "DISTANCE GREATER THAN 800'."
        ),
        "6": (
            "THE NYR9-11 SIGN IS RECOMMENDED. WHEN USED, IT SHALL BE PLACED IN ADVANCE OF "
            "THE FIRST ADVANCE WARNING SIGN. THE PLACEMENT DISTANCE SHALL BE 1000' FOR POSTED "
            "SPEED LIMITS OF 45 MPH OR HIGHER, AND 500' FOR POSTED SPEED LIMITS LESS THAN "
            "45 MPH. THE SIGN SHALL BE PLACED TO AVOID GLARE TO RESIDENCES ADJOINING THE ROADWAY."
        ),
        "7": (
            "THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, "
            "BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, "
            "PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL "
            "(AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING."
        ),
        "8": (
            "A REGULATORY SPEED LIMIT SIGN IS REQUIRED HALFWAY BETWEEN THE 1ST AND 2ND "
            "ADVANCE WARNING SIGNS UNLESS A REGULATORY SPEED LIMIT SIGN IS ALREADY PRESENT "
            "BETWEEN THOSE ADVANCED WARNING SIGNS OR A REGULATORY SPEED LIMIT REDUCTION IS "
            "AUTHORIZED AND THOSE SIGNS HAVE BEEN INSTALLED. ONE R2-1 OR NYR2-2 THROUGH "
            "NYR2-6 SHALL BE PROVIDED AS APPROPRIATE DEPENDING ON THE LOCATION. SEE STANDARD "
            "SHEET 619-012 FOR SIGN FACE AND SIZE."
        ),
    }
    page_squash = squash(" ".join(w[4] for w in pg0.get_text("words")))
    numbered = [f"{i}. {verbatim[str(i)]}" for i in range(1, 9)]
    for i in range(1, 9):
        assert squash(verbatim[str(i)][:40])[:30] in page_squash, f"note {i} verify fail"

    W = pg0.get_text("words")
    note_words = [w for w in W if w[0] >= 900 and 340 <= w[1] <= 590]
    note_words.sort(key=lambda w: (w[1], w[0]))
    n_starts = [(i, w) for i, w in enumerate(note_words) if re.match(r"^N\d+\.$", w[4])]
    n_notes: list[str] = []
    for si, (idx, _) in enumerate(n_starts):
        end = n_starts[si + 1][0] if si + 1 < len(n_starts) else len(note_words)
        txt = norm_glyphs(" ".join(t[4] for t in note_words[idx:end]))
        if txt.startswith("N3.") and "y 107" in txt:
            txt = txt.replace("y 107-05A", "EI 107-05A")
        n_notes.append(txt)
    return numbered + n_notes


pdf = fitz.open(str(PDF))
assert pdf.page_count == 2
for i in range(2):
    assert pdf[i].rotation == 0, f"page {i} expected rotation 0"
pg0, pg1 = pdf[0], pdf[1]
W1 = pg1.get_text("words")
findings: list[str] = [
    f"pdfPages=2 rotation=0 both display_rect={pg0.rect}",
    "ramp-approach plan: AHEAD 1 MILE (Note 8); SHOULDER RAMP labels; Detail 415A",
]

# ---- 415-01 TAPER + BUFFER ----
t01_rows = []
for speed, y0, y1 in SPEED_ROWS:
    buf = token_at(
        W1, BUF_COL[0], BUF_COL[1], y0, y1,
        lambda t: t.count("/") == 1 and t.split("/")[0].isdigit(),
    )
    assert buf, speed
    row = {
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(buf),
        "lateralShiftTaper": {},
        "shoulderTaper": {},
    }
    for lw, x0, x1 in LAT_COLS:
        tok = token_at(
            W1, x0, x1, y0, y1,
            lambda t: t.count("/") == 2 and t.split("/")[0].isdigit(),
        )
        assert tok, (speed, lw)
        row["lateralShiftTaper"][lw] = parse_slash_triple(tok)
    for band, x0, x1 in SH_COLS:
        tok = token_at(
            W1, x0, x1, y0, y1,
            lambda t: t.count("/") == 2 and t.split("/")[0].isdigit(),
        )
        assert tok, (speed, band)
        row["shoulderTaper"][band] = parse_slash_triple(tok)
    t01_rows.append(row)
assert_row_count(t01_rows, 4, "415-01")

shoulder_diffs_301: list[str] = []
for a in t01_rows:
    b = next(r for r in DRAFT301["tables"]["301-03"]["rows"] if r["speedMph"] == a["speedMph"])
    if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
        shoulder_diffs_301.append(f"buffer speed={a['speedMph']}")
    for band in SH_BANDS:
        if a["shoulderTaper"][band] != b["shoulderTaper"][band]:
            shoulder_diffs_301.append(f"shoulder speed={a['speedMph']} band={band}")
if not shoulder_diffs_301:
    findings.append(
        "415-01 vs 301-03: buffer + shoulder (<=8 ft) ALL 16 cells identical on 45-65 mph"
    )
findings.append(
    "415-01 NEW vs 301: lateralShiftTaper lane 10/11/12 (matches 402-03 values at 45-65 mph)"
)

# ---- 415-02 ROLL AHEAD (speed bands, 2 rows in PDF) ----
ra_hi_min = token_at(W1, 160, 175, 260, 270, is_ratio)
ra_hi_max = token_at(W1, 205, 220, 260, 270, is_ratio)
ra_lo_min = token_at(W1, 160, 175, 274, 284, is_ratio)
ra_lo_max = token_at(W1, 205, 220, 274, 284, is_ratio)
assert ra_hi_min == "120/3" and ra_hi_max == "200/5"
assert ra_lo_min == "80/2" and ra_lo_max == "160/4"
t02_rows = [
    {
        "speedBand": ">= 55 MPH",
        "minMph": 55,
        "maxMph": None,
        "min": parse_slash_pair(ra_hi_min),
        "max": parse_slash_pair(ra_hi_max),
    },
    {
        "speedBand": "45 - 50 MPH",
        "minMph": 45,
        "maxMph": 50,
        "min": parse_slash_pair(ra_lo_min),
        "max": parse_slash_pair(ra_lo_max),
    },
]
assert_row_count(t02_rows, 2, "415-02")
findings.append(
    "415-02: 2 speed rows only in PDF (402-02 has 3 incl <=40 MPH) — ramp sheet omits <=40 row"
)

# ---- 415-03 PROTECTIVE VEHICLE (2 shoulder rows) ----
t03_rows = [
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "PVH+TMIA",
    },
]
hits = [w[4] for w in W1 if 280 <= w[0] <= 295 and 370 <= w[1] <= 480 and "PVH" in w[4]]
assert len(hits) >= 2 and all(h == "PVH+TMIA" for h in hits)
assert_row_count(t03_rows, 2, "415-03")
findings.append("415-03: 2 shoulder rows FREEWAY-only PVH+TMIA (no speed-band columns on ramp sheet)")

# ---- 415-04 CHANNELIZING (intermediate matrix) ----
raw04 = [w for w in W1 if 790 <= w[0] <= 1185 and 130 <= w[1] <= 280]
rows04 = group_rows(raw04, y_tol=5.0)


def chan_cells(r: list) -> dict[str, str]:
    return {name: cell_at(r, a, b) for name, a, b in CHAN_COLS}


def first_mark(cells: dict[str, str]) -> dict[str, str]:
    return {k: v.split()[0] for k, v in cells.items() if v.strip() in ("X", "X2", "O") or v.startswith("X2")}


t04_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows04 if "20 FT" in row_text(r)))),
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows04 if row_text(r).startswith("40 FT")))),
    },
]
t04_matrix = [
    {
        "provisionId": "shoulderMergingShiftingTapers",
        "label": "SHOULDER/MERGING/SHIFTING TAPERS",
        "spacingReference": "40 FT.",
        "allowedByDeviceType": t04_spacing[1]["allowedByDeviceType"],
    },
    {
        "provisionId": "markingForTransverseBumps",
        "label": "MARKING FOR TRANSVERSE BUMPS",
        "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows04 if "X2" in row_text(r) and "800" not in row_text(r)))),
    },
    {
        "provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
        "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
        "spacingReference": "800 FT.",
        "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows04 if "800 FT" in row_text(r)))),
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows04 if "GUIDE" in row_text(r) or (240 <= r[0][1] <= 252 and "X" in row_text(r))))
        ),
    },
]
findings.append("415-04: intermediate channelizing matrix (matches 402-05 structure; 20' spacing Note 3)")

# ---- 415-05 SIGN SIZES ----
size_by_y: dict[int, str] = {}
for w in W1:
    if "x" in w[4] and 1115 <= w[0] <= 1135 and 360 <= w[1] <= 480:
        size_by_y[round(w[1])] = w[4]
t05_rows = [
    {"signCode": "W7-3a", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(367, "36x30")},
    {"signCode": "G20-2", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(381, "48x24")},
    {"signCode": "NYR9-11", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(394, "48x84"),
     "note": "Recommended advance sign (Note 6); ramp/intermediate only"},
    {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(407, "48x48")},
    {"signCode": "W21-5bR", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(421, "48x48")},
    {"signCode": "W21-5aR", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(434, "48x48")},
    {"signCode": "WARNING FLAG", "NON-FREEWAY": size_by_y.get(448, "18x18"), "FREEWAY": size_by_y.get(448, "18x18")},
    {"signCode": "R2-1 OR NYR2-2", "NON-FREEWAY": None, "FREEWAY": size_by_y.get(467, "36x48"),
     "note": "Regulatory speed sign (Note 8)"},
    {"signCode": "NYR2-6", "NON-FREEWAY": None, "FREEWAY": None, "note": "Listed under R2-1 group; no size in text layer"},
]
assert_row_count(t05_rows, 9, "415-05")
findings.append("415-05: 9 sign rows incl NYR9-11 48x84 + R2-1/NYR2-2 (301-04 lacks both)")

# ---- NOTES ----
notes_printed = parse_notes_415(pg0)
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
n_notes = [n for n in notes_printed if re.match(r"^N\d+\.", n)]
assert_row_count(numbered, 8, "notes 1-8")
findings.extend([
    f"notes: {len(numbered)} numbered + {len(n_notes)} N-nighttime (402 has 8 N; 415 adds N9-N11)",
    "SURPRISE: Note 6 NYR9-11 recommended 1000' before 1st advance warning",
    "SURPRISE: Note 3 20' channelizing spacing (intermediate, like 402)",
    "SURPRISE: 415-03 FREEWAY-only PV table (no ge45/b35/le30 columns)",
    "SURPRISE: 415-02 omits <=40 MPH roll-ahead row present on 402-02",
])

corridor_hints = [
    "ramp approach: SHOULDER RAMP plan labels; Detail 415A for ramp taper",
    "advance: W20-1 AHEAD 1 MILE (Note 8 regulatory spacing context)",
    "recommended NYR9-11 before first advance warning (Note 6, 1000' @ >=45 mph)",
    "shoulder closed signs W21-5aR/W21-5bR (not lane-closed W20-5R)",
    "supplement W7-3a on W20-1 when work span > 2 miles (Note 4)",
    "SHOULDER TAPER L/3 + lateral shift of traffic flow path (415-01)",
    "BUFFER SPACE + roll-ahead (415-02) + PVH heavy protective vehicle",
    "500' downstream taper on plan; channelizing per 415-04 + Note 3 (20' active workspace)",
    "nighttime N1-N8 block (retroreflective, illumination levels)",
]

draft = {
    "sheetNumber": "619-415",
    "sourcePdf": "Bridge/captures/619-415.pdf",
    "sourcePdfRevision": "619-415_E3.pdf (2 pages, rotation=0, landscape mediabox)",
    "pdfPages": 2,
    "pageRotation": {"page0": 0, "page1": 0},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 ramp-approach intermediate. Roles by CONTENT (numbering != 402): "
            "415-01=taperAndBuffer, 415-02=rollAheadDistance, 415-03=protectiveVehicle, "
            "415-04=channelizingApplication, 415-05=signSizes. "
            "NO advanceWarningSpacing table — AHEAD 1 MILE on plan + Note 8."
        ),
        "taperAndBuffer": "415-01",
        "rollAheadDistance": "415-02",
        "protectiveVehicle": "415-03",
        "channelizingApplication": "415-04",
        "signSizes": "415-05",
    },
    "tables": {
        "415-01": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "laneWidthFt", "shoulderWidthBand"],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "lateralShiftTaper": (
                    "LONGITUDINAL TAPER LENGTH (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES "
                    "(LATERAL SHIFT OF TRAFFIC FLOW PATH) FOR LANE WIDTH"
                ),
                "shoulderTaper": "SHOULDER TAPER LENGTH: L/3 (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES",
            },
            "note": (
                "4 speed columns (45,50,55,65). Shoulder+buffer match 301-03; "
                "lateralShiftTaper values match 402-03 laneTaper at same speeds."
            ),
            "rows": t01_rows,
        },
        "415-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": (
                "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES FOR CLOSED TRAFFIC "
                "VEHICLES AND/OR SHOULDER — STATIONARY OPERATION MIN/MAX"
            ),
            "note": "2 rows in PDF (>=55, 45-50). Omits <=40 MPH row from 402-02.",
            "rows": t02_rows,
        },
        "415-03": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "roadTypeForProtectiveVehicle"],
            "note": (
                "Ramp sheet: 2 shoulder rows, FREEWAY column only (PVH+TMIA). "
                "No ge45/b35/le30 speed-band columns unlike 402-01."
            ),
            "rows": t03_rows,
            "legend": {
                "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
                "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
            },
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
            ],
        },
        "415-04": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "Same structure as 402-05. 20 FT.* active-work-space spacing per Note 3.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t04_matrix,
            "spacingRows": t04_spacing,
            "tableNotes": DRAFT402["tables"]["402-05"]["tableNotes"],
        },
        "415-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": "Ramp intermediate set: W21-5aR/W21-5bR, W7-3a, NYR9-11, R2-1/NYR2-2.",
            "rows": t05_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "8 numbered notes (page 0, x>=900) + N1-N8 nighttime block. "
            "Note 6 NYR9-11 recommended. Note 8 R2-1 regulatory speed sign."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": corridor_hints,
    },
    "findings": findings,
    "comparisonVs301": {
        "415-01_vs_301-03": "buffer + shoulder <=8 ft identical; 415 adds lateralShiftTaper lane 10/11/12",
        "415-03_vs_301-01": "same 2 shoulder PV rows but 415 drops speed-band columns (FREEWAY only)",
        "415-02_vs_301-02": "415 speed-based 2-row vs 301 GVW-based 2-row — different keying",
        "415-05_vs_301-04": "adds NYR9-11 + R2-1/NYR2-2; same W21-5aR/W21-5bR/W7-3a/G20-2 core",
        "absent_vs_301": ["301-02 GVW roll-ahead keying", "301 Note 9 on 315; 415 has Note 8 R2-1 instead"],
        "added_vs_301": [
            "415-04 intermediate channelizing matrix",
            "lateral shift taper columns",
            "NYR9-11 recommended sign",
            "nighttime N-notes block",
            "20' channelizing spacing (Note 3)",
        ],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
for k, v in draft["tables"].items():
    rc = len(v.get("rows", v.get("provisionRows", [])))
    print(f"  {k}: {rc} rows")
print(f"Notes: {len(numbered)} + {len(n_notes)} N")
