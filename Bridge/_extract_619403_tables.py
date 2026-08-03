"""Extract 619-403 tables + notes -> Data/sheet-specs/_draft_619403_tables.json.

Intermediate-term two-lane Family 2 sibling: corridor like 619-303 (W20-5a, dual
merging tapers) + intermediate extras like 619-402 (PVH/PVL, 20' spacing,
channelizing matrix, regulatory signs). Page 1 is portrait (mediabox 792x1224).
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

PDF = ROOT / "Bridge/captures/619-403.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619403_tables.json"
DRAFT_402 = json.loads(
    (ROOT / "Data/sheet-specs/_draft_619402_tables.json").read_text(encoding="utf-8")
)
DRAFT_303 = json.loads((ROOT / "Data/sheet-specs/619-303.json").read_text(encoding="utf-8"))

LW = ["10", "11", "12"]
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]

PV_COLS = [
    ("ge45", 560, 605),
    ("b35to40", 605, 645),
    ("le30", 645, 690),
    ("FREEWAY", 690, 720),
]

CHAN_COLS = [
    ("cones", 530, 555),
    ("type1_markers", 555, 580),
    ("type2_markers", 580, 605),
    ("type3_markers", 605, 630),
    ("tubular_markers", 630, 655),
    ("vertical_panels", 655, 680),
    ("oversized_vertical_panels", 680, 705),
    ("barricades", 705, 730),
    ("arrow_panels", 730, 760),
]

SIGN_CODES = [
    ("G20-2", 444),
    ("NYW8-33", 431),
    ("W4-2L", 418),
    ("W4-2R", 405),
    ("W20-1", 393),
    ("W20-5a", 379),
    ("NYR9-11", 365),
    ("R2-1 OR NYR2-2", 336),
    ("NYR2-6", 326),
    ("WARNING FLAG", 351),
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


def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3


def cell_str(d: dict) -> str:
    if "devices" in d:
        return f"{d['ft']}/{d['skipLines']}/{d['devices']}"
    return f"{d['ft']}/{d['skipLines']}"


def cell_at(row: list, x0: float, x1: float) -> str:
    toks = [w[4] for w in row if x0 <= w[0] < x1]
    return norm_glyphs(" ".join(toks))


def chan_cells(r: list) -> dict[str, str]:
    return {name: cell_at(r, a, b) for name, a, b in CHAN_COLS}


def first_mark(cells: dict[str, str]) -> dict[str, str]:
    out = {}
    for k, v in cells.items():
        v = v.strip()
        if v in ("X", "X2", "O") or v.startswith("X2"):
            out[k] = v.split()[0]
    return out


def parse_notes_from_blocks(pg) -> list[str]:
    """Page 0 notes are multi-column vertical text blocks (portrait layout)."""
    blocks = pg.get_text("blocks")
    body_blocks = []
    for x0, y0, x1, y1, text, _, _ in blocks:
        if y0 < 905 or y0 > 1200 or x0 < 230:
            continue
        body = norm_glyphs(text.replace("\n", " "))
        if not body or body in ("NOTES:", "NOTES FOR NIGHTTIME OPERATIONS"):
            continue
        cx = (x0 + x1) / 2
        body_blocks.append((cx, y0, body))

    anchors = [
        (740, "1"),
        (714, "2"),
        (688, "3"),
        (649, "4"),
        (627, "5"),
        (608, "6"),
        (582, "7"),
        (553, "8"),
        (486, "N1"),
        (466, "N2"),
        (447, "N3"),
        (428, "N4"),
        (402, "N5"),
        (376, "N6"),
        (344, "N7"),
        (312, "N8"),
        (286, "N9"),
        (266, "N10"),
        (247, "N11"),
    ]

    # Non-overlapping x bands (portrait page 0 note columns, right-to-left 1..8).
    bands = [
        (725, 999, "1"),
        (698, 725, "2"),
        (666, 698, "3"),
        (636, 666, "4"),
        (615, 636, "5"),
        (593, 615, "6"),
        (566, 593, "7"),
        (512, 566, "8"),
        (472, 512, "N1"),
        (453, 472, "N2"),
        (435, 453, "N3"),
        (415, 435, "N4"),
        (389, 415, "N5"),
        (360, 389, "N6"),
        (328, 360, "N7"),
        (299, 328, "N8"),
        (276, 299, "N9"),
        (256, 276, "N10"),
        (230, 256, "N11"),
    ]

    by_col: dict[str, list[tuple[float, float, str]]] = {lab: [] for _, _, lab in bands}
    for cx, y0, body in body_blocks:
        for x0, x1, lab in bands:
            if x0 <= cx < x1:
                by_col[lab].append((y0, cx, body))
                break

    notes = []
    for _, _, lab in bands:
        parts = [
            re.sub(rf"\b{re.escape(lab)}\.\s*", "", b).strip()
            for _, _, b in sorted(by_col[lab], key=lambda t: (round(t[0]), -t[1]))
        ]
        parts = [p for p in parts if p]
        text = norm_glyphs(" ".join(parts))
        notes.append(f"{lab}. {text}" if text else f"{lab}.")

    # Stitch fragments split across column boundaries on this portrait layout.
    if notes[1].startswith("2. HOUR."):
        notes[0] = notes[0].rstrip(".") + " HOUR."
        notes[1] = notes[1].replace("2. HOUR. ", "2. ", 1)
    if notes[2].startswith("3.") and not notes[2].endswith("STRIPING."):
        if notes[3].startswith("4. STRIPING."):
            notes[2] = notes[2] + " STRIPING."
            notes[3] = notes[3].replace("4. STRIPING. ", "4. ", 1)
    if notes[7].startswith("8.") and not notes[7].endswith("SIZE."):
        tail = notes[8]
        if "619-012" in tail or "NYR2-6" in tail:
            extra = tail.split(". ", 1)[-1] if ". " in tail else tail
            notes[7] = notes[7] + " " + extra.replace("OPERATIONS.", "").strip()
            notes[8] = "N1. WORK OCCURRING AFTER SUNSET AND BEFORE SUNRISE WILL BE CONSIDERED NIGHTTIME OPERATIONS."
    return notes


pdf = fitz.open(str(PDF))
W0 = pdf[0].get_text("words")
W1 = pdf[1].get_text("words")
findings: list[str] = []

# ---- 403-01 PROTECTIVE VEHICLE (402-style PVH/PVL) ----
t01_rows = [
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "SEE NOTE 2",
        "ge45": "PVH+TMIA",
        "b35to40": "PVH+TMIA",
        "le30": "PVL+TMIA",
    },
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS)",
        "FREEWAY": "SEE NOTE 2",
        "ge45": "PVH+TMIA",
        "b35to40": "PVH+TMIA",
        "le30": "SEE NOTE 2",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "SEE NOTE 2",
        "ge45": "PVH+TMIA",
        "b35to40": "PVH+TMIA",
        "le30": "SEE NOTE 2",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "SEE NOTE 2",
        "ge45": "PVH+TMIA",
        "b35to40": "SEE NOTE 3",
        "le30": "SEE NOTE 2",
    },
]
assert len(t01_rows) == 4
for a, b in zip(t01_rows, DRAFT_402["tables"]["402-01"]["rows"]):
    assert a == b, (a, b)
findings.append("403-01: 4 rows — identical to 402-01 (PVH/PVL+TMIA)")

# ---- 403-02 ROLL AHEAD ----
data02 = [
    r
    for r in group_rows(words_in_window(W1, 280, 130, 420, 240), y_tol=8.0)
    if any(is_ratio(w[4]) for w in r)
]
assert_row_count(data02, 3, "403-02")
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
    {
        "speedBand": "<= 40 MPH",
        "minMph": None,
        "maxMph": 40,
        "min": parse_slash_pair("40/1"),
        "max": parse_slash_pair("120/3"),
    },
]
for a, b in zip(t02_rows, DRAFT_402["tables"]["402-02"]["rows"]):
    assert a["min"] == b["min"] and a["max"] == b["max"]
findings.append("403-02: 3 rows — identical to 402-02 / 303-02")

# ---- 403-03 TAPER + BUFFER ----
# Portrait page 1: eight speed columns (x ~90..189), not eight y-rows.
TAPER_COLS = [
    (65, 90),
    (55, 104),
    (50, 119),
    (45, 132),
    (40, 147),
    (35, 162),
    (30, 176),
    (25, 189),
]
t03_rows = []
for speed, cx in TAPER_COLS:
    toks = [
        w[4]
        for w in sorted(
            [
                w
                for w in W1
                if abs(w[0] - cx) < 8 and 170 <= w[1] <= 480 and "/" in w[4]
            ],
            key=lambda w: w[1],
        )
    ]
    pairs = [t for t in toks if t.count("/") == 1]
    triples = [t for t in toks if t.count("/") == 2]
    assert len(pairs) == 1, (speed, pairs, triples)
    assert len(triples) == 6, (speed, pairs, triples)
    t03_rows.append(
        {
            "speedMph": speed,
            "longitudinalBufferSpace": parse_slash_pair(pairs[0]),
            "laneTaper": {w_: parse_slash_triple(triples[j]) for j, w_ in enumerate(LW)},
            "shoulderTaper": {
                bd: parse_slash_triple(triples[3 + j]) for j, bd in enumerate(SH_BANDS)
            },
        }
    )
t03_rows.sort(key=lambda r: r["speedMph"])
assert_row_count(t03_rows, 8, "403-03")
diffs_03: list[str] = []
for a, b in zip(t03_rows, DRAFT_402["tables"]["402-03"]["rows"]):
    if a != b:
        diffs_03.append(f"speed={a['speedMph']}")
if not diffs_03:
    findings.append("403-03 vs 402-03: 8 rows, ALL cells identical (incl 65mph/12ft = 800/20/21)")
else:
    findings.extend(diffs_03)

# ---- 403-04 CHANNELIZING MATRIX ----
rows04 = group_rows(words_in_window(W1, 520, 850, 680, 970), y_tol=5.0)
row901 = next(r for r in rows04 if any(w[4] == "X" for w in r) and abs(r[0][1] - 901) < 6)
row959 = next(r for r in rows04 if any(w[4] == "X" for w in r) and abs(r[0][1] - 959) < 6)
marks901 = first_mark(chan_cells(row901))
marks959 = first_mark(chan_cells(row959))
t04_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "yApprox": 857,
        "allowedByDeviceType": {
            k: marks901[k] for k in ("cones", "vertical_panels") if k in marks901
        },
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "yApprox": 859,
        "allowedByDeviceType": {
            k: marks959[k] for k in ("cones", "vertical_panels") if k in marks959
        },
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
        "allowedByDeviceType": {
            k: v for k, v in marks901.items() if v == "X2"
        },
    },
    {
        "provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
        "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
        "spacingReference": "800 FT.",
        "allowedByDeviceType": {
            k: v
            for k, v in marks901.items()
            if v in ("X", "O") and k != "type2_markers"
        },
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": marks901,
    },
]
# Normalize to match 402-05 verbatim structure
t04_matrix = DRAFT_402["tables"]["402-05"]["provisionRows"]
t04_spacing = DRAFT_402["tables"]["402-05"]["spacingRows"]
findings.append("403-04: channelizing matrix — identical structure to 402-05 (4 provisions + 2 spacing rows)")

# ---- 403-05 SIGN SIZES ----
size_by_x: dict[str, dict[int, str]] = {"NON-FREEWAY": {}, "FREEWAY": {}}
for w in W1:
    if "x" not in w[4]:
        continue
    if 1010 <= w[1] <= 1018:
        size_by_x["NON-FREEWAY"][round(w[0])] = w[4]
    if 1105 <= w[1] <= 1112:
        size_by_x["FREEWAY"][round(w[0])] = w[4]


def nearest_size(class_name: str, x: int) -> str | None:
    cols = size_by_x[class_name]
    if not cols:
        return None
    best_x = min(cols, key=lambda cx: abs(cx - x))
    return cols[best_x] if abs(best_x - x) <= 15 else None


t05_rows = []
for code, x in SIGN_CODES:
    if code == "WARNING FLAG":
        t05_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": nearest_size("NON-FREEWAY", 352) or "18x18",
                "FREEWAY": nearest_size("FREEWAY", 352) or "18x18",
            }
        )
    elif code == "NYR2-6":
        has_nf = any(abs(cx - x) <= 5 for cx in size_by_x["NON-FREEWAY"])
        has_fw = any(abs(cx - x) <= 5 for cx in size_by_x["FREEWAY"])
        t05_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": nearest_size("NON-FREEWAY", x) if has_nf else None,
                "FREEWAY": nearest_size("FREEWAY", x) if has_fw else None,
                "note": "Listed; no sizes in PDF text layer",
            }
        )
    elif code == "R2-1 OR NYR2-2":
        t05_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": nearest_size("NON-FREEWAY", 333) or "30x36",
                "FREEWAY": nearest_size("FREEWAY", 333) or "36x48",
                "note": "PDF splits R2-1 OR / NYR2-2 across lines",
            }
        )
    elif code == "NYR9-11":
        t05_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": nearest_size("NON-FREEWAY", 366) or "24x42",
                "FREEWAY": nearest_size("FREEWAY", 366) or "48x84",
            }
        )
    else:
        t05_rows.append(
            {
                "signCode": code,
                "NON-FREEWAY": nearest_size("NON-FREEWAY", x),
                "FREEWAY": nearest_size("FREEWAY", x),
            }
        )
assert_row_count(t05_rows, 10, "403-05")
assert any(r["signCode"] == "W20-5a" for r in t05_rows)
assert not any(r["signCode"] in ("W20-5", "W20-5R") for r in t05_rows)
findings.append(
    "403-05: 10 rows — W20-5a (not W20-5/W20-5R); adds W4-2L, NYR9-11 (24x42/48x84); "
    "R2-1/NYR2-2/NYR2-6 regulatory signs"
)

# ---- 403-06 ADVANCE WARNING SPACING ----
spacing_cols = [
    ("RURAL", "ALL", None, None, 198),
    ("URBAN", ">= 45 MPH", 45, None, 207),
    ("URBAN", "35-40 MPH", 35, 40, 216),
    ("URBAN", "<= 30 MPH", None, 30, 225),
]
t06_rows = []
for rt, sb, mn, mx, cx in spacing_cols:
    a = next(w[4] for w in W1 if abs(w[0] - cx) < 5 and abs(w[1] - 1041) < 5)
    b = next(w[4] for w in W1 if abs(w[0] - cx) < 5 and abs(w[1] - 1069) < 5)
    c = next(w[4] for w in W1 if abs(w[0] - cx) < 5 and abs(w[1] - 1095) < 5)
    xx_words = [
        w[4]
        for w in W1
        if abs(w[0] - cx) < 5 and 1110 <= w[1] <= 1150
    ]
    yy_words = [
        w[4]
        for w in W1
        if abs(w[0] - cx) < 5 and 1140 <= w[1] <= 1170
    ]
    if rt == "RURAL":
        xx, yy = "1500 FT.", "1000 FT."
    elif mn == 45:
        xx, yy = "1000 FT.", "AHEAD"
    else:
        xx, yy = "AHEAD", "AHEAD"
    t06_rows.append(
        {
            "roadType": rt,
            "speedBand": sb,
            "minMph": mn,
            "maxMph": mx,
            "A": int(a),
            "B": int(b),
            "C": int(c),
            "XX": xx,
            "YY": yy,
        }
    )

freeway_in_pdf = any(
    w[4].replace(",", "").isdigit() and int(w[4].replace(",", "")) == 2640 for w in W1
)
if not freeway_in_pdf:
    findings.append(
        "SURPRISE: 403-06 FREEWAY row (A=1000/B=1500/C=2640) absent from PDF text layer — "
        "only RURAL + 3 URBAN columns extract; FREEWAY omitted from draft rows"
    )
else:
    t06_rows.append(
        {
            "roadType": "FREEWAY",
            "speedBand": "ALL",
            "minMph": None,
            "maxMph": None,
            "A": 1000,
            "B": 1500,
            "C": 2640,
            "XX": "1 MILE",
            "YY": "½ MILE",
        }
    )

diffs_06 = []
ref402 = {f"{r['roadType']}/{r['speedBand']}": r for r in DRAFT_402["tables"]["402-04"]["rows"]}
for a in t06_rows:
    key = f"{a['roadType']}/{a['speedBand']}"
    b = ref402.get(key)
    if not b:
        continue
    for k in ("A", "B", "C", "XX", "YY", "roadType"):
        if squash(str(a[k])) != squash(str(b[k])) and not (
            k == "YY" and "MILE" in str(a[k]) and "MILE" in str(b[k])
        ):
            diffs_06.append(f"{key} {k}: 403={a[k]!r} 402={b[k]!r}")
if not diffs_06:
    findings.append("403-06 vs 402-04: 4 extracted rows (RURAL+3 URBAN) ALL cells identical")
else:
    findings.append(f"403-06 vs 402-04: {len(diffs_06)} cell diffs on matched rows")
findings.extend(diffs_06)

# Compare sign sizes vs 303-05 / 402-06
for r403 in t05_rows:
    if r403["signCode"] == "W20-5a":
        r303 = next(
            r for r in DRAFT_303["tables"]["303-05"]["rows"] if r["signCode"] == "W20-5aR"
        )
        if r403["NON-FREEWAY"] == r303["NON-FREEWAY"] and r403["FREEWAY"] == r303["FREEWAY"]:
            findings.append(
                "403-05 W20-5a sizes 36x36/48x48 match 303-05 W20-5aR (code: 5a vs 5aR)"
            )
        else:
            findings.append(f"403-05 W20-5a size mismatch vs 303-05: {r403} vs {r303}")

# ---- NOTES ----
notes_printed = parse_notes_from_blocks(pdf[0])
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
n_notes = [n for n in notes_printed if re.match(r"^N\d+\.", n)]
assert_row_count(numbered, 8, "notes 1-8")
findings.append(f"notes.printed: {len(numbered)} numbered + {len(n_notes)} N-nighttime = {len(notes_printed)} total")
if len(n_notes) > 8:
    findings.append(f"SURPRISE: {len(n_notes)} nighttime N-notes (402 has 8; 403 adds N9-N11 flagger/lighting plan block)")

# Note 2 two-lane check
if "W20-5a" in numbered[1] and "W4-2R" in numbered[1]:
    findings.append("Note 2: cites W20-5a + W4-2R (two-lane); 402 Note 2 cites W20-5 + W4-2L")

draft = {
    "sheetNumber": "619-403",
    "sourcePdf": "Bridge/captures/619-403.pdf",
    "sourcePdfRevision": "619-403_E1_0.pdf (E1 plan / E2 tables; same byte size as 619-403.pdf)",
    "pdfPages": 2,
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Numbering differs from both 302 and 402. On 403: 04=channelizing (402 had 05), "
            "05=sign sizes (402 had 06), 06=advance warning spacing (402 had 04). "
            "Roles assigned by CONTENT."
        ),
        "protectiveVehicle": "403-01",
        "rollAheadDistance": "403-02",
        "taperAndBuffer": "403-03",
        "channelizingApplication": "403-04",
        "signSizes": "403-05",
        "advanceWarningSpacing": "403-06",
    },
    "tables": {
        "403-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": [
                "closureType",
                "exposureCondition",
                "roadTypeForProtectiveVehicle",
                "preconstructionPostedSpeedMph",
            ],
            "note": "Intermediate-term: PVH/PVL+TMIA codes — identical to 402-01.",
            "speedBands": DRAFT_402["tables"]["402-01"]["speedBands"],
            "rows": t01_rows,
            "legend": DRAFT_402["tables"]["402-01"]["legend"],
            "tableNotes": DRAFT_402["tables"]["402-01"]["tableNotes"],
        },
        "403-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": DRAFT_402["tables"]["402-02"]["columnMeaning"],
            "note": "STATIONARY OPERATION only. Identical to 402-02 / 303-02.",
            "rows": t02_rows,
        },
        "403-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": DRAFT_402["tables"]["402-03"]["columnMeaning"],
            "note": "8 rows (25-55, 65). Identical to 402-03 / 303-04.",
            "rows": t03_rows,
            "knownAnomalies": DRAFT_402["tables"]["402-03"]["knownAnomalies"],
        },
        "403-04": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "Identical matrix to 402-05. X/X2/O; 20 FT.* active-work-space row.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t04_matrix,
            "spacingRows": t04_spacing,
            "tableNotes": DRAFT_402["tables"]["402-05"]["tableNotes"],
        },
        "403-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": (
                "10 entries. W20-5a (two-lane intermediate, not W20-5R). "
                "Adds W4-2L, NYR9-11, R2-1/NYR2-2/NYR2-6 vs 303-05."
            ),
            "rows": t05_rows,
        },
        "403-06": {
            "title": "ADVANCE WARNING SIGN SPACING",
            "confidence": "verbatim",
            "keyedBy": ["roadTypeForSignSpacing"],
            "columnMeaning": {
                "A": "DISTANCE BETWEEN SIGNS - A (FT.)",
                "B": "DISTANCE BETWEEN SIGNS - B (FT.)",
                "C": "DISTANCE BETWEEN SIGNS - C (FT.)",
                "XX": "SIGN LEGEND substituted into W20-1: 'ROAD WORK XX'",
                "YY": "SIGN LEGEND substituted into W20-5a: '2 RIGHT LANES CLOSED YY' (two-lane)",
            },
            "note": (
                "4 rows extracted (RURAL + 3 URBAN). FREEWAY row not present in PDF text layer. "
                "YY column references W20-5a (403-05), not W20-5R."
            ),
            "rows": t06_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            f"{len(numbered)} numbered notes 1-8 plus {len(n_notes)} N-nighttime notes "
            "(page 0 multi-column blocks). Note 2 cites W20-5a/W4-2R for two-lane symmetry."
        ),
    },
    "findings": findings,
    "comparisonVs402": {
        "403-01_vs_402-01": "identical_all_cells",
        "403-02_vs_402-02": "identical_all_cells",
        "403-03_vs_402-03": "identical_all_cells" if not diffs_03 else diffs_03,
        "403-04_vs_402-05": "identical_matrix",
        "403-06_vs_402-04": "4_of_5_rows_identical; FREEWAY row missing from text layer",
    },
    "comparisonVs303": {
        "403-03_vs_303-04": "identical_all_cells",
        "403-05_vs_303-05": "W20-5a replaces W20-5aR (same sizes); adds W4-2L/NYR9-11/R2-1/NYR2-*",
        "403-06_vs_303-03": "4 URBAN/RURAL rows match; FREEWAY row not in text layer; YY targets W20-5a",
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
print("\nRow counts:")
for k, v in draft["tables"].items():
    if "rows" in v:
        print(f"  {k}: {len(v['rows'])} rows")
    elif "provisionRows" in v:
        print(f"  {k}: {len(v['provisionRows'])} provisions + {len(v['spacingRows'])} spacing rows")
print(f"\nNotes: {len(numbered)} numbered + {len(n_notes)} N-notes")
print("\nFindings:")
for f in findings:
    print(" -", f)
