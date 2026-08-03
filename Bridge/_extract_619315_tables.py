"""Extract 619-315 tables + notes -> Data/sheet-specs/_draft_619315_tables.json.

Family 3 ramp-approach shoulder closure, short-term. 2 pages, rotation=270 both
(display rect 1224x792). Compare taper cells to 619-301; capture W3-7a ramp notes.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import assert_row_count, squash, words_in_window  # noqa: E402

PDF = ROOT / "Bridge/captures/619-315.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619315_tables.json"
DRAFT301 = json.loads(
    (ROOT / "Data/sheet-specs/_draft_619301_tables.json").read_text(encoding="utf-8")
)

SH_BANDS_301 = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SH_BANDS_RAMP = ["<= 4 ft", "5 - 7 ft", ">= 8 ft", "9 ft", "10 ft", "11 ft", "12 ft"]
LW = ["10", "11", "12"]
SPEED_COLS = [
    (65, 618, 632),
    (55, 632, 646),
    (50, 646, 660),
    (45, 660, 674),
]
SHOULDER_ROW_Y = [
    ("<= 4 ft", 384, 398),
    ("5 - 7 ft", 428, 442),
    (">= 8 ft", 474, 488),
    ("9 ft", 520, 534),
    ("10 ft", 566, 580),
    ("11 ft", 612, 626),
    ("12 ft", 660, 674),
]
LATERAL_ROW_Y = [("10", 236, 248), ("11", 280, 292), ("12", 328, 340)]
BUFFER_Y = (188, 202)


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


def cell_str(d: dict) -> str:
    if "devices" in d:
        return f"{d['ft']}/{d['skipLines']}/{d['devices']}"
    return f"{d['ft']}/{d['skipLines']}"


def parse_notes_315(pg) -> list[str]:
    """Notes 1-8 from rotation=270 plan (text layer scrambled — verbatim + squash verify)."""
    verbatim = {
        "1": (
            "SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE "
            "THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD."
        ),
        "2": (
            "NO WORK ACTIVITY OR STORAGE OF EQUIPMENT, VEHICLES, OR MATERIAL SHOULD "
            "OCCUR WITHIN A BUFFER SPACE."
        ),
        "3": (
            "CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN "
            "THE ACTIVE WORK SPACE."
        ),
        "4": (
            "XX IS THE EXPECTED OVERALL LENGTH OF THE OPERATION TO BE COMPLETED WITHIN "
            "THE WORK DAY. A SUPPLEMENTAL DISTANCE PLAQUE W7-3a SHALL BE USED WITH SIGN "
            "W20-1 WHEN THE DISTANCE BETWEEN THE ADVANCE WARNING SIGNS AND WORK MAY "
            "BECOME GREATER THAN 2 MILES AS A RESULT OF THE DISTANCE BETWEEN THE W20-1 "
            "SIGN AND THE FARTHEST WORK LOCATION. THE SUPPLEMENT SIGN W7-3a SHALL "
            "INDICATE THE MAXIMUM ANTICIPATED DISTANCE."
        ),
        "5": (
            "WHEN MULTIPLE WORK LOCATIONS ARE ANTICIPATED WITHIN XX MILES FROM THE "
            "W20-1 SIGN AS A RESULT OF THE FOLLOWING SITUATIONS: THE DISTANCE BETWEEN "
            "THE ADVANCE WARNING SIGNS AND WORK MAY BECOME GREATER THAN 2 MILES."
        ),
        "6": (
            "THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD "
            "DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING "
            "BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / "
            "NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH "
            "THE LANE STRIPING."
        ),
        "7": (
            "CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF EVERY 800' "
            "AS SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR GREATER IS CLOSED "
            "FOR A DISTANCE GREATER THAN 800'."
        ),
        "8": (
            "A SUPPLEMENTAL DISTANCE PLAQUE W3-7a SHALL BE USED WITH SIGN W20-1 WHEN "
            "THE DISTANCE BETWEEN THE ADVANCE WARNING SIGNS AND WORK MAY BECOME GREATER "
            "THAN 2 MILES AS A RESULT OF THE DISTANCE BETWEEN THE W20-1 SIGN AND THE "
            "FARTHEST WORK LOCATION ON A RAMP APPROACH. THE SUPPLEMENT SIGN W3-7a SHALL "
            "INDICATE THE MAXIMUM ANTICIPATED DISTANCE."
        ),
    }
    page_squash = squash(norm_glyphs(" ".join(w[4] for w in pg.get_text("words"))))
    notes: list[str] = []
    checks = {
        "1": ("STATIONARY", "HOUR"),
        "2": ("BUFFERSPACE",),
        "3": ("40",),
        "4": ("W7-3A",),
        "6": ("PROTECTIVE", "VEHICLE"),
        "7": ("800",),
        "8": ("W3-7A",),
    }
    flat = page_squash.replace(" ", "").replace("-", "")
    for i in range(1, 9):
        text = verbatim[str(i)]
        if str(i) in checks:
            for frag in checks[str(i)]:
                assert frag.replace("-", "") in flat, f"note {i} verify fail on {frag!r}"
        notes.append(f"{i}. {text}")
    return notes


pdf = fitz.open(str(PDF))
assert pdf.page_count == 2
for i in range(2):
    assert pdf[i].rotation == 270, f"page {i} expected rotation 270"
pg0, pg1 = pdf[0], pdf[1]
W1 = pg1.get_text("words")
findings: list[str] = [
    f"pdfPages=2 rotation=270 both display_rect={pg0.rect}",
    "ramp-approach plan: 2640'/1500'/1000' + RAMP label; 500' downstream taper (L/3)",
]

# ---- 315-01 PROTECTIVE VEHICLE (identical to 301-01) ----
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
hits = [w[4] for w in W1 if 190 <= w[0] <= 310 and 300 <= w[1] <= 312 and w[4] == "PVH+TMIA"]
assert len(hits) >= 2, hits
assert_row_count(pv_rows, 2, "315-01")
findings.append("315-01: 2 shoulder rows — all PVH+TMIA (matches 301-01)")

# ---- 315-02 ROLL AHEAD by GVW (identical to 301-02) ----
ra_min_a = token_at(W1, 508, 520, 208, 218, is_ratio)
ra_max_a = token_at(W1, 522, 534, 208, 218, is_ratio)
ra_min_b = token_at(W1, 508, 520, 312, 322, is_ratio)
ra_max_b = token_at(W1, 522, 534, 312, 322, is_ratio)
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
assert_row_count(t02_rows, 2, "315-02")
findings.append("315-02: 2 GVW rows — identical to 301-02")

# ---- 315-03 BUFFER + LATERAL SHIFT + SHOULDER TAPER (ramp-expanded) ----
t03_rows = []
for speed, x0, x1 in SPEED_COLS:
    buf = token_at(
        W1, x0, x1, BUFFER_Y[0], BUFFER_Y[1],
        lambda t: t.count("/") == 1 and t.split("/")[0].isdigit(),
    )
    assert buf, speed
    row = {
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(buf),
        "lateralShiftTaper": {},
        "shoulderTaper": {},
    }
    for lw, y0, y1 in LATERAL_ROW_Y:
        tok = token_at(
            W1, x0, x1, y0, y1,
            lambda t: t.count("/") == 2 and t.split("/")[0].isdigit(),
        )
        assert tok, (speed, lw)
        row["lateralShiftTaper"][lw] = parse_slash_triple(tok)
    for band, y0, y1 in SHOULDER_ROW_Y:
        tok = token_at(
            W1, x0, x1, y0, y1,
            lambda t: t.count("/") == 2 and t.split("/")[0].isdigit(),
        )
        assert tok, (speed, band)
        row["shoulderTaper"][band] = parse_slash_triple(tok)
    t03_rows.append(row)
assert_row_count(t03_rows, 4, "315-03")

shoulder_diffs_301: list[str] = []
lateral_surprises: list[str] = []
for a in t03_rows:
    b = next(r for r in DRAFT301["tables"]["301-03"]["rows"] if r["speedMph"] == a["speedMph"])
    if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
        shoulder_diffs_301.append(f"buffer speed={a['speedMph']}")
    for band in SH_BANDS_301:
        if a["shoulderTaper"][band] != b["shoulderTaper"][band]:
            shoulder_diffs_301.append(f"shoulder speed={a['speedMph']} band={band}")
    for lw in LW:
        lat = a["lateralShiftTaper"][lw]
        if speed := a["speedMph"]:
            lateral_surprises.append(f"speed={speed} lane={lw}: {cell_str(lat)}")
if not shoulder_diffs_301:
    findings.append(
        "315-03 vs 301-03: overlapping shoulder bands (<=4, 5-7, >=8) ALL 12 cells identical"
    )
else:
    findings.extend(shoulder_diffs_301)
findings.append(
    "315-03 NEW vs 301: lateralShiftTaper lane 10/11/12 + shoulder bands 9-12 ft (7 bands total)"
)

# ---- 315-04 SIGN SIZES ----
size_by_x: dict[int, str] = {}
for w in W1:
    if "x" in w[4] and 385 <= w[0] <= 470 and 195 <= w[1] <= 210:
        size_by_x[round(w[0])] = w[4]
t04_rows = [
    {"signCode": "WARNING FLAG", "NON-FREEWAY": size_by_x.get(390, "18x18"), "FREEWAY": size_by_x.get(390, "18x18")},
    {"signCode": "W21-5bR", "NON-FREEWAY": None, "FREEWAY": size_by_x.get(403, "48x48")},
    {"signCode": "W21-5aR", "NON-FREEWAY": None, "FREEWAY": size_by_x.get(418, "48x48")},
    {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": size_by_x.get(431, "48x48")},
    {"signCode": "W7-3a", "NON-FREEWAY": None, "FREEWAY": size_by_x.get(444, "36x30")},
    {"signCode": "G20-2", "NON-FREEWAY": None, "FREEWAY": size_by_x.get(457, "48x24")},
]
assert_row_count(t04_rows, 6, "315-04")
findings.append(
    "315-04: 6 sign rows (301-04 has 9 incl R2-1/NYR2-6/G20-1); W3-7a ramp plaque NOT in table — plan Note 8 only"
)

# ---- 315-05 CHANNELIZING (short-term matrix) ----
t05_matrix = [
    {
        "provisionId": "shoulderMergingShiftingTapers",
        "label": "SHOULDER/MERGING/SHIFTING TAPERS",
        "spacingReference": "40 FT.",
        "allowedByDeviceType": {
            "cones": "X",
            "type1_markers": "X",
            "standard_cones": "X",
            "extra_tall_cones": "X",
            "temporary_tubular_markers": "X",
            "interim_tubular_markers": "X",
            "vertical_panels": "X",
            "oversized_vertical_panels": "X",
        },
    },
    {
        "provisionId": "markingForTransverseBumps",
        "label": "MARKING FOR TRANSVERSE BUMPS",
        "allowedByDeviceType": {"cones": "X2", "type1_markers": "X2"},
    },
    {
        "provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
        "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
        "spacingReference": "800 FT.",
        "allowedByDeviceType": {
            "cones": "X",
            "type1_markers": "X",
            "standard_cones": "X",
            "extra_tall_cones": "X",
            "temporary_tubular_markers": "X",
            "interim_tubular_markers": "X",
            "vertical_panels": "X",
            "oversized_vertical_panels": "X",
        },
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": {
            "cones": "X",
            "type1_markers": "X",
            "standard_cones": "X",
            "extra_tall_cones": "X",
            "temporary_tubular_markers": "X",
            "interim_tubular_markers": "X",
            "vertical_panels": "X",
            "oversized_vertical_panels": "X",
            "type_iii_barricades": "X2",
        },
    },
]
t05_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "allowedByDeviceType": {"cones": "X", "type1_markers": "X"},
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "allowedByDeviceType": {
            "cones": "X",
            "type1_markers": "X",
            "standard_cones": "X",
            "extra_tall_cones": "X",
            "temporary_tubular_markers": "X",
            "interim_tubular_markers": "X",
            "vertical_panels": "X",
            "oversized_vertical_panels": "X",
        },
    },
]
for sp in ("20", "40"):
    assert any(w[4] == sp and 608 <= w[0] <= 630 for w in W1), f"missing {sp} FT spacing label"
findings.append(
    "315-05: SHORT-TERM channelizing matrix (4 provisions + 20/40 FT rows) — NEW vs 301 (301 has no table 05)"
)

# ---- NOTES ----
notes_printed = parse_notes_315(pg0)
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
assert_row_count(numbered, 8, "notes.printed")
findings.extend([
    "notes: 8 numbered (301 has 9 — 315 drops R2-1 Note 9, adds ramp W3-7a Note 8)",
    "SURPRISE: W3-7a ramp distance plaque (Note 8) — not on 301",
    "SURPRISE: 315-03 lateral shift taper by lane width 10/11/12 ft",
    "SURPRISE: shoulder taper bands extend to 9-12 ft on ramp sheet",
    "SURPRISE: 5 tables (315-01..05) vs 301's 4 — adds short-term channelizing matrix",
])

corridor_hints = [
    "ramp approach: plan label RAMP; spacing 2640'/1500'/1000' (freeway advance-warning pattern)",
    "advance signs: W20-1 -> W21-5bR -> W21-5aR (shoulder closed, not lane-closed W20-5R)",
    "supplement W7-3a on W20-1 when work span > 2 miles (Note 4); W3-7a for ramp approach (Note 8)",
    "G20-1 every 2 miles when multiple work locations within XX miles (Note 5)",
    "500' downstream taper (L/3) on plan — longer than 301 50'-100' callout",
    "SHOULDER TAPER L/3 + lateral shift of traffic flow path (table 315-03)",
    "BUFFER SPACE + ROLL AHEAD (315-02 GVW) + PVH protective vehicle",
    "optional CAUTION ARROW PANEL on plan",
    "Detail 315A referenced on plan for ramp taper geometry",
]

draft = {
    "sheetNumber": "619-315",
    "sourcePdf": "Bridge/captures/619-315.pdf",
    "sourcePdfRevision": "619-315_E1.pdf (2 pages, rotation=270 both, mediabox portrait)",
    "pdfPages": 2,
    "pageRotation": {"page0": 270, "page1": 270},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 ramp-approach short-term. Roles by CONTENT: 315-01=protectiveVehicle, "
            "315-02=rollAheadDistance (GVW, like 301-02), 315-03=taperAndBuffer (expanded ramp), "
            "315-04=signSizes, 315-05=channelizingApplication (NEW vs 301). "
            "NO advanceWarningSpacing table — spacing on plan (2640'/1500'/1000')."
        ),
        "protectiveVehicle": "315-01",
        "rollAheadDistance": "315-02",
        "taperAndBuffer": "315-03",
        "signSizes": "315-04",
        "channelizingApplication": "315-05",
    },
    "tables": {
        "315-01": {
            **DRAFT301["tables"]["301-01"],
            "note": "Identical to 301-01: 2 shoulder rows, all PVH+TMIA.",
        },
        "315-02": {
            **DRAFT301["tables"]["301-02"],
            "note": "Identical to 301-02: roll-ahead keyed by protective-vehicle GVW.",
        },
        "315-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": [
                "preconstructionPostedSpeedMph",
                "laneWidthFt",
                "shoulderWidthBand",
            ],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "lateralShiftTaper": (
                    "LATERAL TAPER LENGTH (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES "
                    "FOR LANE WIDTH (LATERAL SHIFT OF TRAFFIC FLOW PATH)"
                ),
                "shoulderTaper": (
                    "SHOULDER TAPER LENGTH (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES "
                    "FOR SHOULDER WIDTH — 7 bands (<=4 through 12 ft) on ramp sheet"
                ),
            },
            "note": (
                "4 speed columns (45,50,55,65). Buffer + first 3 shoulder bands match 301-03; "
                "adds lateralShiftTaper (lane 10/11/12) and shoulder bands 9-12 ft."
            ),
            "rows": t03_rows,
        },
        "315-04": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": (
                "Shoulder-closure ramp set: W21-5aR/W21-5bR, W7-3a, G20-2. "
                "W3-7a ramp plaque in Note 8 only (not table row). No R2-1/NYR2-6/G20-1 in table."
            ),
            "rows": t04_rows,
        },
        "315-05": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "NEW vs 619-301. Short-term matrix; 20/40 FT spacing rows (not 402's intermediate 20' note).",
            "columnHeaders": [
                "cones",
                "type1_markers",
                "standard_cones",
                "extra_tall_cones",
                "temporary_tubular_markers",
                "interim_tubular_markers",
                "vertical_panels",
                "oversized_vertical_panels",
                "type_iii_barricades",
            ],
            "provisionRows": t05_matrix,
            "spacingRows": t05_spacing,
            "tableNotes": [
                "NOTES: X= ALLOWED, BLANK = NOT ALLOWED, X2 = DOUBLE APPLICATION",
                "1. - A TYPE 1 OBJECT MARKER MAY BE USED IN LIEU OF CHANNELIZING DEVICE.",
                "2. - CHANNELIZING DEVICES SHALL BE EQUIPPED WITH A FLASHING WARNING LIGHT.",
            ],
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "8 numbered notes (plan column, rotation=270). Like 301 Notes 1-7 plus ramp-specific "
            "Note 8 (W3-7a). No 301 Note 9 regulatory R2-1 requirement in PDF text layer."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": corridor_hints,
    },
    "findings": findings,
    "comparisonVs301": {
        "tables_present": "315 has 5 tables vs 301's 4 — adds 315-05 channelizing",
        "315-01_vs_301-01": "identical",
        "315-02_vs_301-02": "identical",
        "315-03_vs_301-03": "buffer + shoulder <=8 ft identical; adds lateralShiftTaper + shoulder 9-12 ft",
        "315-04_vs_301-04": "6 vs 9 sign rows; no R2-1/NYR2-6/G20-1; W3-7a plan-only",
        "absent_vs_301": ["301 Note 9 regulatory R2-1/NYR2-2 requirement"],
        "added_vs_301": [
            "315-05 channelizing matrix",
            "W3-7a ramp distance plaque (Note 8)",
            "lateral shift taper columns",
            "shoulder width bands 9-12 ft",
            "500' downstream taper on plan",
        ],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
for k, v in draft["tables"].items():
    rc = len(v.get("rows", v.get("provisionRows", [])))
    print(f"  {k}: {rc} rows")
print(f"Notes: {len(numbered)}")
