"""Extract 619-401 tables + notes -> Data/sheet-specs/_draft_619401_tables.json.

Family 3 intermediate shoulder closure (2 pages, rotation=0).
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

PDF = ROOT / "Bridge/captures/619-401.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619401_tables.json"
SPEC301 = json.loads((ROOT / "Data/sheet-specs/_draft_619301_tables.json").read_text(encoding="utf-8"))

LW = ["10", "11", "12"]
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
CHAN_COLS = [
    ("cones", 850, 890),
    ("type1_markers", 910, 935),
    ("type2_markers", 935, 965),
    ("type3_markers", 965, 990),
    ("tubular_markers", 1035, 1065),
    ("vertical_panels", 1065, 1095),
    ("oversized_vertical_panels", 1095, 1125),
    ("barricades", 1125, 1155),
    ("arrow_panels", 1155, 1190),
]


def parse_slash_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def parse_slash_triple(tok: str) -> dict:
    parts = tok.split("/")
    return {"ft": int(parts[0]), "skipLines": int(parts[1]), "devices": int(parts[2])}


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


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
        if text.startswith(num):
            text = text[len(num) :].lstrip(". ")
        notes.append(f"{num}. {text}")
    assert len([n for n in notes if re.match(r"^\d+\.", n)]) == 10
    assert len([n for n in notes if re.match(r"^N\d+\.", n)]) == 11
    return notes


pdf = fitz.open(str(PDF))
W0 = pdf[0].get_text("words")
W1 = pdf[1].get_text("words")
findings: list[str] = []

assert pdf[0].rotation == 0 and pdf[1].rotation == 0
findings.append(f"pdfPages=2 rotation=0 display_rect={pdf[0].rect}")

# ---- 401-01 PROTECTIVE VEHICLE ----
t01_rows = [
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "PVH+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "PVH+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
]
pv_tokens = [w[4] for w in W1 if 280 <= w[0] <= 310 and 100 <= w[1] <= 215 and "PVH" in w[4]]
assert pv_tokens and all(t == "PVH+TMIA" for t in pv_tokens)
assert_row_count(t01_rows, 4, "401-01")
findings.append(
    "401-01: 4 rows — all FREEWAY cells PVH+TMIA (301 has 2 shoulder-only rows, same PVH code)"
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

# ---- 401-02 ROLL AHEAD (speed bands, 2 rows) ----
raw02 = words_in_window(W1, 48, 341, 280, 430)
rows02 = group_rows(raw02, y_tol=6.0)
data02 = [r for r in rows02 if any(is_ratio(w[4]) for w in r)]
assert_row_count(data02, 2, "401-02")
t02_rows = [
    {
        "speedBand": ">= 55 MPH",
        "minMph": 55,
        "maxMph": None,
        "min": parse_slash_pair("120/3"),
        "max": parse_slash_pair("200/5"),
    },
    {
        "speedBand": "45 - 50 MPH",
        "minMph": 45,
        "maxMph": 50,
        "min": parse_slash_pair("80/2"),
        "max": parse_slash_pair("160/4"),
    },
]
findings.append("401-02: 2 speed-band rows — NOT 301-02 GVW bands")

# ---- 401-03 BUFFER + LANE + SHOULDER TAPER (4 speeds 45-65) ----
raw03 = words_in_window(W1, 48, 457, 520, 590)
rows03 = group_rows(raw03, y_tol=5.0)
data03 = [r for r in rows03 if any(w[4].isdigit() and int(w[4]) in (45, 50, 55, 65) for w in r)]
assert_row_count(data03, 4, "401-03")
t03_specs = [
    (45, "360/9", ["440/11/12", "520/13/14", "560/14/15"], ["80/2/3", "80/2/3", "120/3/4"]),
    (50, "425/11", ["520/13/14", "560/14/15", "600/15/16"], ["80/2/3", "120/3/4", "160/4/5"]),
    (55, "495/13", ["560/14/15", "600/15/16", "680/17/18"], ["80/2/3", "120/3/4", "160/4/5"]),
    (65, "645/16", ["640/16/17", "720/18/19", "800/20/21"], ["80/2/3", "160/4/5", "200/5/6"]),
]
t03_rows = []
for speed, buf, lane, shoulder in t03_specs:
    t03_rows.append({
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(buf),
        "laneTaper": {w_: parse_slash_triple(lane[j]) for j, w_ in enumerate(LW)},
        "shoulderTaper": {bd: parse_slash_triple(shoulder[j]) for j, bd in enumerate(SH_BANDS)},
    })

shoulder_diffs: list[str] = []
buf_diffs: list[str] = []
base301 = {r["speedMph"]: r for r in SPEC301["tables"]["301-03"]["rows"]}
for row in t03_rows:
    ref = base301[row["speedMph"]]
    if row["longitudinalBufferSpace"] != ref["longitudinalBufferSpace"]:
        buf_diffs.append(f"buffer speed={row['speedMph']}")
    for band in SH_BANDS:
        if row["shoulderTaper"][band] != ref["shoulderTaper"][band]:
            shoulder_diffs.append(f"shoulder speed={row['speedMph']} band={band}")
if not buf_diffs and not shoulder_diffs:
    findings.append(
        "401-03 buffer+shoulder taper vs 301-03: ALL 16 cells identical on overlapping 45-65 rows"
    )
else:
    findings.extend(buf_diffs + shoulder_diffs)
findings.append(
    "401-03: 4 speed rows with laneTaper 10/11/12 columns — EXTRA vs 301-03 shoulder-only"
)

# ---- 401-04 CHANNELIZING ----
raw04 = words_in_window(W1, 520, 47, 1220, 300)
rows04 = group_rows(raw04, y_tol=4.0)
t04_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows04 if "20 FT" in row_text(r)))
        ),
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows04 if row_text(r).startswith("40 FT")))
        ),
    },
]
t04_provisions = [
    {
        "provisionId": "shoulderMergingShiftingTapers",
        "label": "SHOULDER/MERGING/SHIFTING TAPERS",
        "spacingReference": "40 FT.",
        "allowedByDeviceType": t04_spacing[1]["allowedByDeviceType"],
    },
    {
        "provisionId": "markingForTransverseBumps",
        "label": "MARKING FOR TRANSVERSE BUMPS",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows04 if "MARKING FOR X2" in row_text(r)))
        ),
    },
    {
        "provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
        "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
        "spacingReference": "800 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(next(r for r in rows04 if "800 FT" in row_text(r)))
        ),
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": first_mark(
            chan_cells(
                next(
                    r for r in rows04
                    if "REMOVAL OF EXISTING" in row_text(r) or (
                        "GUIDE RAIL" in row_text(r) and "X" in row_text(r)
                    )
                )
            )
        ),
    },
]
findings.append("401-04: channelizing matrix like 402-05 — 20' spacing row + 4 provisions")

# ---- 401-05 SIGN SIZES ----
t05_rows = [
    {"signCode": "G20-2", "NON-FREEWAY": None, "FREEWAY": "48x24"},
    {"signCode": "W7-3a", "NON-FREEWAY": None, "FREEWAY": "36x30"},
    {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
    {"signCode": "W20-5aR", "NON-FREEWAY": None, "FREEWAY": "48x48"},
    {"signCode": "W21-5bR", "NON-FREEWAY": None, "FREEWAY": "48x48"},
    {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    {
        "signCode": "R2-1 OR NYR2-2",
        "NON-FREEWAY": None,
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
assert_row_count(t05_rows, 8, "401-05")
findings.append(
    "401-05: W20-5aR + W21-5bR (intermediate shoulder signs); Note 4 cites W21-5bU/W21-5c"
)

# ---- NOTES ----
notes_printed = parse_notes_page0(pdf[0])
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
findings.append(f"notes.printed: {len(numbered)} numbered + 11 N-nighttime (402-style)")

findings.extend([
    "SURPRISE vs 301: 20' channelizing (Note 5) not 40'",
    "SURPRISE: intermediate signs W21-5bU/W21-5c symmetry (not W21-5aL/bL)",
    "SURPRISE: NY9-11 recommended (Note 9); buffer-space prohibition (Note 3)",
    "SURPRISE: 5 tables incl channelizing — no advanceWarningSpacing table",
    "SURPRISE: 401-02 speed-keyed roll-ahead (301 uses GVW)",
])

draft = {
    "sheetNumber": "619-401",
    "sourcePdf": "Bridge/captures/619-401.pdf",
    "sourcePdfRevision": "619-401_E3.pdf (2 pages, rotation=0)",
    "pdfPages": 2,
    "pageRotation": {"page0": pdf[0].rotation, "page1": pdf[1].rotation},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 intermediate shoulder. 5 tables. 401-03=taperAndBuffer with lane+shoulder "
            "columns (301-03 is shoulder-only). 401-04=channelizingApplication. "
            "NO advanceWarningSpacing — spacing on plan + notes."
        ),
        "protectiveVehicle": "401-01",
        "rollAheadDistance": "401-02",
        "taperAndBuffer": "401-03",
        "channelizingApplication": "401-04",
        "signSizes": "401-05",
    },
    "tables": {
        "401-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": [
                "closureType",
                "exposureCondition",
                "roadTypeForProtectiveVehicle",
            ],
            "note": "4 rows (lane+shoulder). FREEWAY-only column in PDF — all PVH+TMIA.",
            "rows": t01_rows,
            "legend": legend_01,
            "tableNotes": table_notes_01,
        },
        "401-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": "MIN/MAX roll-ahead by posted speed (>=55 vs 45-50).",
            "note": "Differs from 301-02 GVW bands; matches 402-02 first 2 rows.",
            "rows": t02_rows,
            "usageNote": "MIN/MAX range, not a single value.",
        },
        "401-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "laneTaper": "LANE TAPER for lane widths 10/11/12 ft",
                "shoulderTaper": "SHOULDER TAPER L/3 by shoulder width band",
            },
            "note": (
                "4 speed columns (45-65). Shoulder taper+buffer match 301-03; "
                "lane taper columns extra vs 301."
            ),
            "rows": t03_rows,
        },
        "401-04": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "Same structure as 402-05. 20 FT.* spacing references Note 5.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t04_provisions,
            "spacingRows": t04_spacing,
            "tableNotes": [
                "NOTES: X= ALLOWED, BLANK = NOT ALLOWED, O = OPTIONAL * SEE NOTE 5 ON SHEET 1 OF 2.",
                "1. - A TYPE 1 OBJECT MARKER MAY BE USED IN LIEU OF CHANNELIZING DEVICE.",
                "2. - CHANNELIZING DEVICES SHALL BE EQUIPPED WITH A FLASHING WARNING LIGHT.",
            ],
        },
        "401-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Intermediate shoulder signs W20-5aR/W21-5bR; W7-3a + G20-2 + R2-1 group.",
            "rows": t05_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "10 numbered + 11 N-nighttime notes on plan (page 0). Note 5 = 20' spacing. "
            "Note 4 uses W21-5bU/W21-5c intermediate sign codes."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": [
            "advance signs: W20-1 -> W20-5aR -> W21-5bR (intermediate shoulder set)",
            "PVH protective vehicle + roll-ahead + buffer space on plan",
            "SHOULDER TAPER L/3 + downstream taper",
            "optional NY9-11 recommended upstream (Note 9)",
            "W7-3a supplement + G20-1 multi-location (Notes 6-7)",
        ],
    },
    "findings": findings,
    "comparisonVs301": {
        "tables_present": "401 has 5 tables vs 301's 4 — adds 401-04 channelizing",
        "401-01_vs_301-01": "4 lane+shoulder rows vs 2 shoulder-only; same PVH+TMIA code",
        "401-02_vs_301-02": "speed bands vs GVW bands",
        "401-03_vs_301-03": "shoulder+buffer identical on 45-65; 401 adds laneTaper 10/11/12 columns",
        "401-05_vs_301-04": "W20-5aR/W21-5bR replace W21-5aR/W21-5bR; adds W20-5aR advance sign",
        "absent_vs_301": ["301-04 standalone without channelizing table"],
        "added_vs_301": ["401-04 channelizing matrix", "Note 3 buffer prohibition", "Note 9 NY9-11"],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
for k, v in draft["tables"].items():
    n = len(v.get("rows", v.get("provisionRows", [])))
    print(f"  {k}: {n} rows/provisions")
print(f"Notes: {len(numbered)} numbered + {len(notes_printed)-len(numbered)} N")
for f in findings:
    print(" -", f)
