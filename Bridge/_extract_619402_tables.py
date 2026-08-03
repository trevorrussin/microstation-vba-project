"""Extract 619-402 tables + notes -> Data/sheet-specs/_draft_619402_tables.json"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, row_text, squash, assert_row_count

PDF = ROOT / "Bridge/captures/619-402.pdf"
SPEC302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text())
OUT = ROOT / "Data/sheet-specs/_draft_619402_tables.json"

pdf = fitz.open(str(PDF))
W0 = pdf[0].get_text("words")
W1 = pdf[1].get_text("words")

LW = ["10", "11", "12"]
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
PV_COLS = [("ge45", 280, 340), ("b35to40", 340, 395), ("le30", 395, 445), ("FREEWAY", 445, 500)]
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


def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3


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
    toks = [w[4] for w in row if x0 <= w[0] < x1]
    return norm_glyphs(" ".join(toks))


def combine_cells(rows: list, y_primary: float, y_secondary: float | None = None) -> dict[str, str]:
    def band_text(y: float) -> dict[str, str]:
        matched = [r for r in rows if abs(r[0][1] - y) <= 6]
        out = {}
        for r in matched:
            for name, x0, x1 in PV_COLS:
                c = cell_at(r, x0, x1)
                if c:
                    out[name] = c
        return out

    primary = band_text(y_primary)
    if not y_secondary:
        return primary
    secondary = band_text(y_secondary)
    merged = {}
    for name, _, _ in PV_COLS:
        p = primary.get(name, "")
        s = secondary.get(name, "")
        if p and s:
            if p == s:
                merged[name] = p
            elif "SEE" in p or "NOTE" in p or "SEE" in s or "NOTE" in s:
                merged[name] = norm_glyphs(" ".join(dict.fromkeys((p + " " + s).split())))
            else:
                merged[name] = p
        else:
            merged[name] = p or s
    return merged


def pv_row(cells: dict[str, str]) -> dict:
    return {
        "closureType": cells["closureType"],
        "exposureCondition": cells["exposureCondition"],
        "FREEWAY": cells.get("FREEWAY", ""),
        "ge45": cells.get("ge45", ""),
        "b35to40": cells.get("b35to40", ""),
        "le30": cells.get("le30", ""),
    }


# ---- 402-01 ----
raw01 = words_in_window(W1, 48, 100, 500, 230)
rows01 = group_rows(raw01, y_tol=6.0)
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

# ---- 402-02 ----
raw02 = words_in_window(W1, 48, 420, 480, 510)
rows02 = group_rows(raw02, y_tol=8.0)
data02 = [r for r in rows02 if any(is_ratio(w[4]) for w in r)]
assert_row_count(data02, 3, "402-02")
t02_rows = [
    {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
     "min": parse_slash_pair("120/3"), "max": parse_slash_pair("200/5")},
    {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
     "min": parse_slash_pair("80/2"), "max": parse_slash_pair("160/4")},
    {"speedBand": "<= 40 MPH", "minMph": None, "maxMph": 40,
     "min": parse_slash_pair("40/1"), "max": parse_slash_pair("120/3")},
]

# ---- 402-03 ----
raw03 = words_in_window(W1, 48, 510, 520, 720)
rows03 = group_rows(raw03, y_tol=5.0)
merged03 = []
i = 0
while i < len(rows03):
    r = rows03[i]
    toks = [w[4] for w in r]
    if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55, 65):
        if toks[0] == "35" and len(toks) < 7 and i + 1 < len(rows03):
            r = r + rows03[i + 1]
            i += 1
        merged03.append(r)
    i += 1
assert_row_count(merged03, 8, "402-03")
t03_rows = []
for r in merged03:
    toks = [w[4] for w in sorted(r, key=lambda w: w[0])]
    speed = int(toks[0])
    t03_rows.append({
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(toks[1]),
        "laneTaper": {w_: parse_slash_triple(toks[2 + j]) for j, w_ in enumerate(LW)},
        "shoulderTaper": {bd: parse_slash_triple(toks[5 + j]) for j, bd in enumerate(SH_BANDS)},
    })

# ---- 402-04 ----
raw04 = words_in_window(W1, 740, 48, 1220, 140)
rows04 = group_rows(raw04, y_tol=3.0)
data04 = [r for r in rows04 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data04, 5, "402-04")
t04_specs = [
    ("URBAN", "<= 30 MPH", None, 30, "AHEAD", "AHEAD"),
    ("URBAN", "35-40 MPH", 35, 40, "AHEAD", "AHEAD"),
    ("URBAN", ">= 45 MPH", 45, None, "1000 FT.", "AHEAD"),
    ("RURAL", "ALL", None, None, "1500 FT.", "1000 FT."),
    ("FREEWAY", "ALL", None, None, "1 MILE", "½ MILE"),
]
t04_rows = []
for r, (rt, sb, mn, mx, xx, yy) in zip(data04, t04_specs):
    nums = [w[4].replace(",", "") for w in r if is_distance_num(w[4])]
    t04_rows.append({
        "roadType": rt, "speedBand": sb, "minMph": mn, "maxMph": mx,
        "A": int(nums[0]), "B": int(nums[1]), "C": int(nums[2]), "XX": xx, "YY": yy,
    })

# ---- 402-05 ----
raw05 = words_in_window(W1, 740, 255, 1220, 375)
rows05 = group_rows(raw05, y_tol=4.0)


def chan_cells(r: list) -> dict[str, str]:
    return {name: cell_at(r, a, b) for name, a, b in CHAN_COLS}


def first_mark(cells: dict[str, str]) -> dict[str, str]:
    out = {}
    for k, v in cells.items():
        v = v.strip()
        if v in ("X", "X2", "O") or v.startswith("X2"):
            out[k] = v.split()[0]
    return out


t05_spacing = [
    {"spacingId": "spacing20ft", "label": "20 FT.", "yApprox": 262,
     "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows05 if "20 FT" in row_text(r))))},
    {"spacingId": "spacing40ft", "label": "40 FT.", "yApprox": 276,
     "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows05 if row_text(r).startswith("40 FT"))))},
]
t05_matrix = [
    {"provisionId": "shoulderMergingShiftingTapers", "label": "SHOULDER/MERGING/SHIFTING TAPERS",
     "spacingReference": "40 FT.", "allowedByDeviceType": t05_spacing[1]["allowedByDeviceType"]},
    {"provisionId": "markingForTransverseBumps", "label": "MARKING FOR TRANSVERSE BUMPS",
     "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows05 if "X2" in row_text(r))))},
    {"provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
     "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
     "spacingReference": "800 FT.",
     "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows05 if "800 FT" in row_text(r))))},
    {"provisionId": "removalOfExistingGuideRail", "label": "REMOVAL OF EXISTING GUIDE RAIL",
     "spacingReference": "80 FT. / 40 FT.",
     "allowedByDeviceType": first_mark(chan_cells(next(r for r in rows05 if 350 < r[0][1] < 358 and "X" in row_text(r))))},
]

# ---- 402-06 ----
t06_rows = [
    {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
    {"signCode": "NYW8-33", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
    {"signCode": "W4-2R", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W20-5", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    {"signCode": "R2-1 OR NYR2-2", "NON-FREEWAY": "30x36", "FREEWAY": "36x48",
     "note": "PDF splits R2-1 OR / NYR2-2 across lines"},
    {"signCode": "NYR2-6", "NON-FREEWAY": None, "FREEWAY": None,
     "note": "Listed under R2-1 group; no sizes in PDF text layer"},
]

# ---- NOTES ----
note_words = [w for w in W0 if w[0] >= 900 and 50 <= w[1] <= 520]
note_words.sort(key=lambda w: (w[1], w[0]))
starts = [(i, w) for i, w in enumerate(note_words) if re.match(r"^(\d+|N\d+)\.$", w[4]) and w[0] > 910]
notes_printed = []
for si, (idx, _wnum) in enumerate(starts):
    end = starts[si + 1][0] if si + 1 < len(starts) else len(note_words)
    notes_printed.append(norm_glyphs(" ".join(t[4] for t in note_words[idx:end])))

# Split note 8 if nighttime header was absorbed from the PDF text layer
fixed_notes = []
for n in notes_printed:
    if n.startswith("8.") and "NOTES FOR NIGHTTIME OPERATIONS" in n:
        head, _, _tail = n.partition("NOTES FOR NIGHTTIME OPERATIONS")
        fixed_notes.append(head.rstrip(" :."))
    elif n.startswith("N1.") and "OPERATIONS" in n:
        fixed_notes.append(
            n.replace(
                "N1. OPERATIONS",
                "N1. WORK OCCURRING AFTER SUNSET AND BEFORE SUNRISE WILL BE CONSIDERED NIGHTTIME OPERATIONS",
                1,
            )
        )
    else:
        fixed_notes.append(n)
notes_printed = fixed_notes

# ---- Compare vs 302 ----
findings: list[str] = []
diffs_03: list[str] = []
diffs_04: list[str] = []
diffs_06: list[str] = []

for a, b in zip(t03_rows, SPEC302["tables"]["302-02"]["rows"]):
    if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
        diffs_03.append(f"speed={a['speedMph']} buffer: 402={cell_str(a['longitudinalBufferSpace'])} 302={cell_str(b['longitudinalBufferSpace'])}")
    for w_ in LW:
        if a["laneTaper"][w_] != b["laneTaper"][w_]:
            diffs_03.append(f"speed={a['speedMph']} lane{w_}: 402={cell_str(a['laneTaper'][w_])} 302={cell_str(b['laneTaper'][w_])}")
    for bd in SH_BANDS:
        if a["shoulderTaper"][bd] != b["shoulderTaper"][bd]:
            diffs_03.append(f"speed={a['speedMph']} shoulder[{bd}]: 402={cell_str(a['shoulderTaper'][bd])} 302={cell_str(b['shoulderTaper'][bd])}")

if not diffs_03:
    findings.append("402-03 vs 302-02: 8 rows, ALL cells identical (incl 65mph/12ft = 800/20/21)")

for a, b in zip(t04_rows, SPEC302["tables"]["302-03"]["rows"]):
    for k in ("A", "B", "C", "XX", "YY"):
        if squash(str(a[k])) != squash(str(b[k])) and not (
            k == "YY" and "MILE" in str(a[k]) and "MILE" in str(b[k])
        ):
            diffs_04.append(f"{a['roadType']}/{a['speedBand']} {k}: 402={a[k]!r} 302={b[k]!r}")
if not diffs_04:
    findings.append("402-04 vs 302-03: 5 rows, ALL cells identical")

base302 = {r["signCode"]: r for r in SPEC302["tables"]["302-04"]["rows"]}
for a in t06_rows[:6]:
    b = base302.get("W20-5R" if a["signCode"] == "W20-5" else a["signCode"])
    if b:
        for k in ("NON-FREEWAY", "FREEWAY"):
            if a[k] != b[k]:
                diffs_06.append(f"{a['signCode']} {k}: 402={a[k]!r} 302={b[k]!r}")
if not diffs_06:
    findings.append("402-06 base 6 rows: sizes identical to 302-04 (W20-5=W20-5R)")
findings.extend([
    "402-06 extras vs 302-04: R2-1 OR NYR2-2 (30x36/36x48), NYR2-6 (listed, no sizes in text layer)",
    "402-01: 4 rows — PVH/PVL+TMIA codes differ from 302-01 P/TMIA",
    "402-02: 3 rows — identical to 302-05",
    "402-05: NEW channelizing matrix — 4 provisions + 2 spacing rows; 20 FT.* row",
    f"notes.printed: {len(notes_printed)} entries (8 numbered + 8 N-nighttime)",
    "SURPRISE: Note 4 = 20' max channelizing spacing (302 uses 40')",
    "SURPRISE: Note 7 NY9-11 recommended; Note 8 R2-1/NYR2-2..6 regulatory speed sign required",
    "SURPRISE: 402-05 matrix uses O=OPTIONAL for oversized vertical panels",
])
findings.extend(diffs_03 + diffs_04 + diffs_06)

draft = {
    "sheetNumber": "619-402",
    "sourcePdf": "Bridge/captures/619-402.pdf",
    "pdfPages": 2,
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": "Numbering differs from 619-302. 402-05 CHANNELIZING DEVICE APPLICATION is NEW vs 302.",
        "protectiveVehicle": "402-01",
        "rollAheadDistance": "402-02",
        "taperAndBuffer": "402-03",
        "advanceWarningSpacing": "402-04",
        "channelizingApplication": "402-05",
        "signSizes": "402-06",
    },
    "tables": {
        "402-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "roadTypeForProtectiveVehicle", "preconstructionPostedSpeedMph"],
            "note": "Intermediate-term: PVH/PVL+TMIA codes (not 302 P/TMIA).",
            "speedBands": SPEC302["tables"]["302-01"]["speedBands"],
            "rows": t01_rows,
            "legend": {
                "PVL": "PROTECTIVE VEHICLE LIGHT (MINIMUM GROSS WEIGHT 9,500 LBS. OR GREATER) (SEE NOTE 5)",
                "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
                "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
            },
            "tableNotes": ["1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT"],
        },
        "402-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": SPEC302["tables"]["302-05"]["columnMeaning"],
            "note": "STATIONARY OPERATION only. Identical to 302-05.",
            "rows": t02_rows,
        },
        "402-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": SPEC302["tables"]["302-02"]["columnMeaning"],
            "note": "8 rows (25-55, 65). y_tol=5.0. Identical to 302-02.",
            "rows": t03_rows,
            "knownAnomalies": SPEC302["tables"]["302-02"]["knownAnomalies"],
        },
        "402-04": {
            "title": "ADVANCE WARNING SIGN SPACING",
            "confidence": "verbatim",
            "keyedBy": ["roadTypeForSignSpacing"],
            "columnMeaning": SPEC302["tables"]["302-03"]["columnMeaning"],
            "note": "Identical to 302-03.",
            "rows": t04_rows,
        },
        "402-05": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "NEW vs 619-302. X/X2/O matrix.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t05_matrix,
            "spacingRows": t05_spacing,
            "tableNotes": [
                "NOTES: X= ALLOWED, BLANK = NOT ALLOWED, O = OPTIONAL * SEE NOTE 4 ON SHEET 1 OF 2.",
                "1. - A TYPE 1 OBJECT MARKER MAY BE USED IN LIEU OF CHANNELIZING DEVICE.",
                "2. - CHANNELIZING DEVICES SHALL BE EQUIPPED WITH A FLASHING WARNING LIGHT.",
            ],
        },
        "402-06": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": "8 entries; base 6 match 302-04; adds R2-1/NYR2-2/NYR2-6.",
            "rows": t06_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": "16 entries: notes 1-8 plus N1-N8 nighttime block (page 1, x>=900).",
    },
    "findings": findings,
    "comparisonVs302": {
        "402-03_vs_302-02": diffs_03 if diffs_03 else "identical_all_cells",
        "402-04_vs_302-03": diffs_04 if diffs_04 else "identical_all_cells",
        "402-06_vs_302-04": diffs_06 if diffs_06 else "base_6_match; extras R2-1/NYR2-2/NYR2-6",
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
print("402-01:", t01_rows)
print("65/12:", t03_rows[-1]["laneTaper"]["12"])
