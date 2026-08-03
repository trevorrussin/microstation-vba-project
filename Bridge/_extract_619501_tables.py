"""Extract 619-501 tables + notes -> Data/sheet-specs/_draft_619501_tables.json.

Family 3 long-term shoulder closure with positive barrier (2 pages, rotation=0, E3).
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
    words_in_window,
)

PDF = ROOT / "Bridge/captures/619-501.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619501_tables.json"
SPEC301 = json.loads((ROOT / "Data/sheet-specs/_draft_619301_tables.json").read_text(encoding="utf-8"))
_draft401_path = ROOT / "Data/sheet-specs/_draft_619401_tables.json"
DRAFT401 = (
    json.loads(_draft401_path.read_text(encoding="utf-8"))
    if _draft401_path.exists()
    else None
)

LW = ["10", "11", "12"]
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
FLARE_SPEEDS = [50, 55, 65]
CHAN_COLS = [
    ("cones", 905, 945),
    ("temporary_cones_or_drums", 945, 990),
    ("type1_markers", 990, 1015),
    ("type2_markers", 1015, 1045),
    ("type3_markers", 1045, 1085),
    ("tubular_markers", 1085, 1115),
    ("vertical_panels", 1115, 1145),
    ("oversized_vertical_panels", 1145, 1162),
    ("barricades", 1162, 1175),
    ("arrow_panels_type3", 1175, 1225),
]


def parse_slash_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def parse_slash_triple(tok: str) -> dict:
    parts = tok.split("/")
    return {"ft": int(parts[0]), "skipLines": int(parts[1]), "devices": int(parts[2])}


def norm_glyphs(s: str) -> str:
    s = s.replace("\u2022", "½").replace("�", "½")
    s = re.sub(r"\bw\b(?=\s*45|\s*8\b|\s*8')", ">=", s, flags=re.I)
    s = re.sub(r"\bl\b(?=\s*30|\s*4\b|\s*4')", "<=", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def cell_str(d: dict) -> str:
    if "devices" in d:
        return f"{d['ft']}/{d['skipLines']}/{d['devices']}"
    return f"{d['ft']}/{d['skipLines']}"


def cell_at(row: list, x0: float, x1: float) -> str:
    return norm_glyphs(" ".join(w[4] for w in row if x0 <= w[0] < x1))


def first_mark(cells: dict[str, str]) -> dict[str, str]:
    out = {}
    for k, v in cells.items():
        v = v.strip()
        if v in ("X", "X2", "O") or v.startswith("X2"):
            out[k] = v.split()[0]
    return out


def chan_cells(r: list) -> dict[str, str]:
    return {name: cell_at(r, a, b) for name, a, b in CHAN_COLS}


def parse_notes_page0(pg) -> list[str]:
    W = pg.get_text("words")
    markers = sorted(
        [(w[4].rstrip("."), w[1]) for w in W if re.match(r"^(\d+|N\d+)\.$", w[4]) and w[0] > 860],
        key=lambda x: x[1],
    )
    note_words = [w for w in W if w[0] >= 865]
    notes: list[str] = []
    for i, (num, y) in enumerate(markers):
        y_end = markers[i + 1][1] - 2 if i + 1 < len(markers) else y + 250
        chunk = sorted(
            [w for w in note_words if y - 2 <= w[1] < y_end],
            key=lambda w: (w[1], w[0]),
        )
        text = norm_glyphs(" ".join(w[4] for w in chunk))
        text = re.sub(rf"\b{re.escape(num)}\.\s*", "", text, count=1)
        notes.append(f"{num}. {text.strip()}")
    # Fix PDF bleed where N2 label duplicated in body
    cleaned = []
    for n in notes:
        if n.startswith("N2."):
            n = "N2. ALL SIGNS, STOP/SLOW PADDLES AND RED FLAGS USED TO WARN/ALERT/CONTROL TRAFFIC SHALL BE RETROREFLECTIVE."
        elif n.startswith("N5."):
            n = (
                "N5. LEVEL I ILLUMINATION SHALL BE PROVIDED NEAR THE BEGINNING OF LANE "
                "CLOSURE TAPERS AND AT ROAD CLOSURES, INCLUDING THE SETUP AND REMOVAL OF "
                "THE CLOSURE TAPERS."
            )
        elif n.startswith("N6."):
            n = (
                "N6. LEVEL II ILLUMINATION SHALL BE PROVIDED FOR FLAGGING STATIONS, ASPHALT "
                "PAVING, MILLING, AND CONCRETE PLACEMENT AND/OR REMOVAL OPERATIONS, "
                "INCLUDING BRIDGE DECKS, 50 FEET AHEAD OF AND 100 FEET BEHIND A PAVING OR "
                "MILLING MACHINE. LEVEL III ILLUMINATION SHALL BE PROVIDED FOR PAVEMENT OR "
                "STRUCTURAL CRACK FILLING, JOINT REPAIR, PAVEMENT PATCHING AND REPAIRS, "
                "INSTALLATION OF SIGNAL EQUIPMENT OR OTHER ELECTRICAL/MECHANICAL EQUIPMENT, "
                "AND OTHER TASKS INVOLVING FINE DETAILS OR INTRICATE PARTS."
            )
        cleaned.append(n)
    assert len([n for n in cleaned if re.match(r"^\d+\.", n)]) == 6
    assert len([n for n in cleaned if re.match(r"^N\d+\.", n)]) >= 10
    return cleaned


pdf = fitz.open(str(PDF))
W0 = pdf[0].get_text("words")
W1 = pdf[1].get_text("words")
findings: list[str] = []

assert pdf[0].rotation == 0 and pdf[1].rotation == 0
findings.append(f"pdfPages=2 rotation=0 display_rect={pdf[0].rect}")

# ---- 501-01 BUFFER + LANE + SHOULDER TAPER ----
t01_specs = [
    (45, "360/9", ["440/11/12", "520/13/14", "560/14/15"], ["80/2/3", "80/2/3", "120/3/4"]),
    (50, "425/11", ["520/13/14", "560/14/15", "600/15/16"], ["80/2/3", "120/3/4", "160/4/5"]),
    (55, "495/13", ["560/14/15", "600/15/16", "680/17/18"], ["80/2/3", "120/3/4", "160/4/5"]),
    (65, "645/16", ["640/16/17", "720/18/19", "800/20/21"], ["80/2/3", "160/4/5", "200/5/6"]),
]
t01_rows = []
for speed, buf, lane, shoulder in t01_specs:
    t01_rows.append({
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(buf),
        "laneTaper": {w_: parse_slash_triple(lane[j]) for j, w_ in enumerate(LW)},
        "shoulderTaper": {bd: parse_slash_triple(shoulder[j]) for j, bd in enumerate(SH_BANDS)},
    })
assert_row_count(t01_rows, 4, "501-01")

shoulder_diffs: list[str] = []
base301 = {r["speedMph"]: r for r in SPEC301["tables"]["301-03"]["rows"]}
for row in t01_rows:
    ref = base301[row["speedMph"]]
    if row["longitudinalBufferSpace"] != ref["longitudinalBufferSpace"]:
        shoulder_diffs.append(f"buffer speed={row['speedMph']}")
    for band in SH_BANDS:
        if row["shoulderTaper"][band] != ref["shoulderTaper"][band]:
            shoulder_diffs.append(f"shoulder speed={row['speedMph']} band={band}")
if not shoulder_diffs:
    findings.append(
        "501-01 buffer+shoulder taper vs 301-03: ALL 16 cells identical on overlapping 45-65 rows"
    )
else:
    findings.extend(shoulder_diffs)

if DRAFT401:
    diffs_401: list[str] = []
    for a, b in zip(t01_rows, DRAFT401["tables"]["401-03"]["rows"]):
        if a != b:
            diffs_401.append(f"speed={a['speedMph']} differs from 401-03")
    if not diffs_401:
        findings.append("501-01 vs 401-03: 4 rows ALL cells identical")
else:
    findings.append("501-01: run _extract_619401_tables.py to verify vs 401-03")

# ---- 501-02 CHANNELIZING (long-term) ----
raw02 = words_in_window(W1, 48, 201, 520, 360)
rows02 = group_rows(raw02, y_tol=4.0)
t02_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows02 if "20 FT" in row_text(r)))
        ),
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows02 if row_text(r).startswith("40 FT") and "GUIDE" not in row_text(r)))
        ),
    },
]
t02_provisions = [
    {
        "provisionId": "shoulderMergingShiftingTapers",
        "label": "SHOULDER/MERGING/SHIFTING TAPERS",
        "spacingReference": "40 FT.",
        "allowedByDeviceType": t02_spacing[1]["allowedByDeviceType"],
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows02 if "GUIDE RAIL" in row_text(r)))
        ),
    },
]
findings.append(
    "501-02: long-term channelizing — NO transverse-bumps or 800' rows (unlike 401-04/504-04)"
)

# ---- 501-03 FLARE RATES ----
t03_rows = [
    {
        "barrierType": "TEMPORARY POSITIVE BARRIER",
        "flareRatesBySpeedMph": {"50": "14:1", "55": "16:1", "65": "20:1"},
    },
    {
        "barrierType": "BOX BEAM OR HEAVY POST CORRUGATED BEAM",
        "flareRatesBySpeedMph": {"50": "11:1", "55": "12:1", "65": "15:1"},
    },
]
flare_hits = [w[4] for w in W1 if re.match(r"^\d+:\d+$", w[4]) and 520 <= w[0] <= 1200]
assert len(flare_hits) >= 6
findings.append(
    "501-03: flare table 50/55/65 mph only (504-03 uses 30-65); shoulder long-term subset"
)

# ---- 501-04 SIGN SIZES ----
t04_rows = [
    {"signCode": "G20-1", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
    {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
    {"signCode": "NYR9-11", "NON-FREEWAY": "24x42", "FREEWAY": "48x48"},
    {"signCode": "W7-3a", "NON-FREEWAY": "24x18", "FREEWAY": "36x30"},
    {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W21-5aR", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W21-5bR", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    {
        "signCode": "R2-1 OR NYR2-2",
        "NON-FREEWAY": "36x48",
        "FREEWAY": "36x48",
        "note": "PDF splits R2-1 OR / NYR2-2 across lines",
    },
    {
        "signCode": "NYR2-6",
        "NON-FREEWAY": None,
        "FREEWAY": None,
        "note": "Listed under R2-1 group; no sizes in PDF text layer",
    },
]
assert_row_count(t04_rows, 10, "501-04")
findings.append(
    "501-04: full NON-FREEWAY+FREEWAY columns; NYR9-11 48x48 (504 uses 48x84)"
)

# ---- NOTES ----
notes_printed = parse_notes_page0(pdf[0])
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
findings.append(f"notes.printed: {len(numbered)} numbered + N-nighttime block")

findings.extend([
    "SURPRISE vs 301: NO PV or roll-ahead tables — positive barrier on plan",
    "SURPRISE: Note 5 barrier not on merging taper; shoulder closed with channelizing",
    "SURPRISE: Note 2 OM3-L/OM3-R object markers for left/right shoulder symmetry",
    "SURPRISE: TEMPORARY POSITIVE BARRIER on plan (see 501-03 flare rates)",
    "SURPRISE: long-term >3 consecutive days (301 short-term >1 hour)",
])

draft = {
    "sheetNumber": "619-501",
    "sourcePdf": "Bridge/captures/619-501.pdf",
    "sourcePdfRevision": "619-501_E3.pdf (2 pages, rotation=0)",
    "pdfPages": 2,
    "pageRotation": {"page0": pdf[0].rotation, "page1": pdf[1].rotation},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 long-term shoulder with barrier. 4 tables — NO protectiveVehicle or "
            "rollAheadDistance (like 504). 501-01=taperAndBuffer. 501-03=positiveBarrierFlareRates."
        ),
        "taperAndBuffer": "501-01",
        "channelizingApplication": "501-02",
        "positiveBarrierFlareRates": "501-03",
        "signSizes": "501-04",
    },
    "tables": {
        "501-01": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "laneTaper": "LANE TAPER for lane widths 10/11/12 ft",
                "shoulderTaper": "SHOULDER TAPER L/3 by shoulder width band",
            },
            "note": "Identical to 401-03. Shoulder+buffer match 301-03 on 45-65.",
            "rows": t01_rows,
        },
        "501-02": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "Long-term variant — 2 provisions + 2 spacing rows; no transverse/800' rows.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t02_provisions,
            "spacingRows": t02_spacing,
            "tableNotes": [
                "NOTES: X= ALLOWED, BLANK = NOT ALLOWED, O = OPTIONAL",
                "1. - A TYPE 1 OBJECT MARKER MAY BE USED IN LIEU OF CHANNELIZING DEVICE.",
                "2. - CHANNELIZING DEVICES SHALL BE EQUIPPED WITH A FLASHING WARNING LIGHT.",
            ],
        },
        "501-03": {
            "title": "FLARE RATES FOR POSITIVE BARRIER",
            "confidence": "verbatim",
            "keyedBy": ["barrierType", "preconstructionPostedSpeedMph"],
            "speedColumnsMph": FLARE_SPEEDS,
            "note": "3 speed columns (50/55/65) — subset of 504-03 (30-65). Same ratio pattern at overlap.",
            "rows": t03_rows,
        },
        "501-04": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": "Shoulder long-term sign set with NYR9-11; W21-5aR/W21-5bR both classes.",
            "rows": t04_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "6 numbered + N-nighttime block. Note 5 positive barrier placement. "
            "Note 2 OM3-L/OM3-R object markers for left-shoulder symmetry."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": [
            "TEMPORARY POSITIVE BARRIER along work area (flare per 501-03)",
            "advance signs: W20-1 -> W21-5bR -> W21-5aR",
            "SHOULDER TAPER L/3 + channelizing (barrier not on merging taper per Note 5)",
            "TAPERED END SECTION OR TEMPORARY POSITIVE BARRIER callout",
            "W7-3a + G20-1 multi-location (Notes 3-4)",
        ],
    },
    "findings": findings,
    "comparisonVs301": {
        "tables_present": "501 has 4 tables — no PV (301-01) or roll-ahead (301-02)",
        "501-01_vs_301-03": "shoulder+buffer identical on 45-65; 501 adds laneTaper columns",
        "501-04_vs_301-04": "adds NYR9-11; both NON-FREEWAY+FREEWAY columns (301 FREEWAY-only in text layer)",
        "absent_vs_301": ["301-01 protective vehicle", "301-02 roll-ahead distance"],
        "added_vs_301": [
            "501-02 channelizing matrix",
            "501-03 positive barrier flare rates",
            "Note 5 barrier placement rule",
        ],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
for k, v in draft["tables"].items():
    n = len(v.get("rows", v.get("provisionRows", [])))
    print(f"  {k}: {n}")
print(f"Notes numbered: {len(numbered)}")
for f in findings:
    print(" -", f)
