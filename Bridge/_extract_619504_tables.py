"""Extract 619-504 tables + notes -> Data/sheet-specs/_draft_619504_tables.json"""
from __future__ import annotations

import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, row_text, squash, assert_row_count

PDF = ROOT / "Bridge/captures/619-504.pdf"
DRAFT402 = json.loads((ROOT / "Data/sheet-specs/_draft_619402_tables.json").read_text())
OUT = ROOT / "Data/sheet-specs/_draft_619504_tables.json"

pdf = fitz.open(str(PDF))
W0 = pdf[0].get_text("words")
W1 = pdf[1].get_text("words")

LW = ["10", "11", "12"]
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
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
FLARE_SPEEDS = [30, 40, 50, 55, 65]


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


def first_mark(cells: dict[str, str]) -> dict[str, str]:
    out = {}
    for k, v in cells.items():
        v = v.strip()
        if v in ("X", "X2", "O") or v.startswith("X2"):
            out[k] = v.split()[0]
    return out


def chan_cells(r: list) -> dict[str, str]:
    return {name: cell_at(r, a, b) for name, a, b in CHAN_COLS}


# ---- 504-01 ADVANCE WARNING SIGN SPACING ----
raw01 = words_in_window(W1, 48, 48, 500, 150)
rows01 = group_rows(raw01, y_tol=3.0)
data01 = [r for r in rows01 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data01, 5, "504-01")
t01_specs = [
    ("URBAN", "<= 30 MPH", None, 30, "AHEAD", "AHEAD"),
    ("URBAN", "35-40 MPH", 35, 40, "AHEAD", "AHEAD"),
    ("URBAN", ">= 45 MPH", 45, None, "1000 FT.", "AHEAD"),
    ("RURAL", "ALL", None, None, "1500 FT.", "1000 FT."),
    ("FREEWAY", "ALL", None, None, "1 MILE", "½ MILE"),
]
t01_rows = []
for r, (rt, sb, mn, mx, xx, yy) in zip(data01, t01_specs):
    nums = [w[4].replace(",", "") for w in r if is_distance_num(w[4])]
    t01_rows.append({
        "roadType": rt, "speedBand": sb, "minMph": mn, "maxMph": mx,
        "A": int(nums[0]), "B": int(nums[1]), "C": int(nums[2]), "XX": xx, "YY": yy,
    })

# ---- 504-02 LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS ----
raw02 = words_in_window(W1, 48, 155, 520, 365)
rows02 = group_rows(raw02, y_tol=5.0)
merged02 = []
i = 0
while i < len(rows02):
    r = rows02[i]
    toks = [w[4] for w in r]
    if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55, 65):
        if toks[0] == "35" and len(toks) < 7 and i + 1 < len(rows02):
            r = r + rows02[i + 1]
            i += 1
        merged02.append(r)
    i += 1
assert_row_count(merged02, 8, "504-02")
t02_rows = []
for r in merged02:
    toks = [w[4] for w in sorted(r, key=lambda w: w[0])]
    speed = int(toks[0])
    t02_rows.append({
        "speedMph": speed,
        "longitudinalBufferSpace": parse_slash_pair(toks[1]),
        "laneTaper": {w_: parse_slash_triple(toks[2 + j]) for j, w_ in enumerate(LW)},
        "shoulderTaper": {bd: parse_slash_triple(toks[5 + j]) for j, bd in enumerate(SH_BANDS)},
    })

# ---- 504-03 FLARE RATES FOR POSITIVE BARRIER ----
raw03 = words_in_window(W1, 48, 410, 520, 490)
rows03 = group_rows(raw03, y_tol=4.0)
flare_data = [r for r in rows03 if any(re.match(r"^\d+:\d+$", w[4]) for w in r)]
assert_row_count(flare_data, 2, "504-03")
t03_rows = [
    {
        "barrierType": "TEMPORARY POSITIVE BARRIER",
        "flareRatesBySpeedMph": {"30": "8:1", "40": "11:1", "50": "14:1", "55": "16:1", "65": "20:1"},
    },
    {
        "barrierType": "BOX BEAM OR HEAVY POST CORRUGATED BEAM",
        "flareRatesBySpeedMph": {"30": "7:1", "40": "9:1", "50": "11:1", "55": "12:1", "65": "15:1"},
    },
]
for r, row in zip(flare_data, t03_rows):
    toks = [w[4] for w in sorted(r, key=lambda w: w[0])]
    ratios = [t for t in toks if re.match(r"^\d+:\d+$", t)]
    assert len(ratios) == 5, f"504-03 {row['barrierType']}: expected 5 ratios, got {ratios}"
    row["flareRatesBySpeedMph"] = dict(zip(FLARE_SPEEDS, ratios))

# ---- 504-04 CHANNELIZING DEVICE APPLICATION ----
raw04 = words_in_window(W1, 740, 130, 1220, 380)
rows04 = group_rows(raw04, y_tol=4.0)

spacing20_row = next(r for r in rows04 if row_text(r).startswith("20 FT"))
spacing40_row = next(r for r in rows04 if row_text(r).startswith("40 FT") and "REMOVAL" not in row_text(r))
transverse_row = next(r for r in rows04 if "800 FT" in row_text(r))
removal_mark_row = next(r for r in rows04 if row_text(r).startswith("X") and len(first_mark(chan_cells(r))) >= 5)

t04_spacing = [
    {
        "spacingId": "spacing20ft",
        "label": "20 FT.",
        "allowedByDeviceType": first_mark(chan_cells(spacing20_row)),
    },
    {
        "spacingId": "spacing40ft",
        "label": "40 FT.",
        "allowedByDeviceType": first_mark(chan_cells(spacing40_row)),
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
        "provisionId": "transverseDeviceWithinClosedLaneOrShoulder",
        "label": "TRANSVERSE DEVICE WITHIN CLOSED TRAFFIC LANE AND/OR SHOULDER",
        "spacingReference": "800 FT.",
        "allowedByDeviceType": first_mark(chan_cells(transverse_row)),
    },
    {
        "provisionId": "removalOfExistingGuideRail",
        "label": "REMOVAL OF EXISTING GUIDE RAIL",
        "spacingReference": "80 FT. / 40 FT.",
        "allowedByDeviceType": first_mark(chan_cells(removal_mark_row)),
    },
]

# ---- 504-05 REQUIRED SIGN SIZES ----
t05_rows = [
    {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
    {"signCode": "NYR9-11", "NON-FREEWAY": "24x42", "FREEWAY": "48x84"},
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
note_words = [w for w in W0 if w[0] >= 900 and 50 <= w[1] <= 560]
note_words.sort(key=lambda w: (w[1], w[0]))
starts = [(i, w) for i, w in enumerate(note_words) if re.match(r"^(\d+|N\d+)\.$", w[4]) and w[0] > 910]
notes_printed = []
for si, (idx, _wnum) in enumerate(starts):
    end = starts[si + 1][0] if si + 1 < len(starts) else len(note_words)
    notes_printed.append(norm_glyphs(" ".join(t[4] for t in note_words[idx:end])))

fixed_notes = []
for n in notes_printed:
    if n.startswith("7.") and "NOTES FOR NIGHTTIME OPERATIONS" in n:
        head, _, _tail = n.partition("NOTES FOR NIGHTTIME OPERATIONS")
        fixed_notes.append(head.rstrip(" :."))
    else:
        fixed_notes.append(n)
notes_printed = fixed_notes

# ---- Compare vs 402 ----
findings: list[str] = []
diffs_02: list[str] = []
diffs_01: list[str] = []
diffs_05: list[str] = []

for a, b in zip(t02_rows, DRAFT402["tables"]["402-03"]["rows"]):
    if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
        diffs_02.append(f"speed={a['speedMph']} buffer: 504={cell_str(a['longitudinalBufferSpace'])} 402={cell_str(b['longitudinalBufferSpace'])}")
    for w_ in LW:
        if a["laneTaper"][w_] != b["laneTaper"][w_]:
            diffs_02.append(f"speed={a['speedMph']} lane{w_}: 504={cell_str(a['laneTaper'][w_])} 402={cell_str(b['laneTaper'][w_])}")
    for bd in SH_BANDS:
        if a["shoulderTaper"][bd] != b["shoulderTaper"][bd]:
            diffs_02.append(f"speed={a['speedMph']} shoulder[{bd}]: 504={cell_str(a['shoulderTaper'][bd])} 402={cell_str(b['shoulderTaper'][bd])}")

if not diffs_02:
    findings.append("504-02 vs 402-03: 8 rows, ALL cells identical")

for a, b in zip(t01_rows, DRAFT402["tables"]["402-04"]["rows"]):
    for k in ("A", "B", "C", "XX", "YY"):
        if squash(str(a[k])) != squash(str(b[k])) and not (
            k == "YY" and "MILE" in str(a[k]) and "MILE" in str(b[k])
        ):
            diffs_01.append(f"{a['roadType']}/{a['speedBand']} {k}: 504={a[k]!r} 402={b[k]!r}")
if not diffs_01:
    findings.append("504-01 vs 402-04: 5 rows, ALL cells identical")

base402 = {r["signCode"]: r for r in DRAFT402["tables"]["402-06"]["rows"]}
for a in t05_rows:
    lookup = "W20-5" if a["signCode"] == "W20-5" else a["signCode"]
    b = base402.get(lookup)
    if b and a["signCode"] not in ("NYR9-11",):
        for k in ("NON-FREEWAY", "FREEWAY"):
            if a.get(k) != b.get(k):
                diffs_05.append(f"{a['signCode']} {k}: 504={a.get(k)!r} 402={b.get(k)!r}")

findings.extend([
    "504-03: NEW — FLARE RATES FOR POSITIVE BARRIER (2 barrier types × 5 speed columns)",
    "NO 504 PV table (402-01) or ROLL AHEAD table (402-02) — long-term uses positive barrier, not PV",
    "504-04 vs 402-05: LONG-TERM title; no 'marking for transverse bumps' row; O on barricades not oversized panels for transverse/removal",
    "504-04 columns include TEMPORARY CONES/DRUMS (402 used type1/2/3 marker columns)",
    "504-05: NYR9-11 replaces NYW8-33 from 402-06; W20-5 (not W20-5R) same sizes",
    f"notes.printed: {len(notes_printed)} entries (7 numbered + 11 N-nighttime)",
    "SURPRISE: Note 1 = >3 consecutive days (402 intermediate = up to 3 days)",
    "SURPRISE: Note 4-6 positive/movable barrier + min lane widths 11'/10'",
    "SURPRISE: Note 2 uses W20-5L + OM3-L (402 uses W20-5 without L suffix)",
    "SURPRISE: No 20' channelizing note (402 Note 4); no R2-1 regulatory speed note (402 Note 8)",
    "SURPRISE: N9 written nighttime plan required; N11 flagger flashlight",
    f"pdfPages rotation: page0={pdf[0].rotation} page1={pdf[1].rotation} (no rotation)",
])
findings.extend(diffs_02 + diffs_01 + diffs_05)

draft = {
    "sheetNumber": "619-504",
    "sourcePdf": "Bridge/captures/619-504.pdf",
    "pdfPages": 2,
    "pageRotation": {"page0": pdf[0].rotation, "page1": pdf[1].rotation},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": "5 tables (504-01..504-05). No PV or roll-ahead tables — long-term positive-barrier sheet.",
        "advanceWarningSpacing": "504-01",
        "taperAndBuffer": "504-02",
        "positiveBarrierFlareRates": "504-03",
        "channelizingApplication": "504-04",
        "signSizes": "504-05",
    },
    "tables": {
        "504-01": {
            "title": "ADVANCE WARNING SIGN SPACING",
            "confidence": "verbatim",
            "keyedBy": ["roadTypeForSignSpacing"],
            "columnMeaning": DRAFT402["tables"]["402-04"]["columnMeaning"],
            "note": "Identical to 402-04 / 302-03.",
            "rows": t01_rows,
        },
        "504-02": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": DRAFT402["tables"]["402-03"]["columnMeaning"],
            "note": "8 rows (25-55, 65). y_tol=5.0. Identical to 402-03.",
            "rows": t02_rows,
            "knownAnomalies": DRAFT402["tables"]["402-03"]["knownAnomalies"],
        },
        "504-03": {
            "title": "FLARE RATES FOR POSITIVE BARRIER",
            "confidence": "verbatim",
            "keyedBy": ["barrierType", "preconstructionPostedSpeedMph"],
            "speedColumnsMph": FLARE_SPEEDS,
            "note": "NEW vs 402. Ratio L:1 flare rate by barrier type and posted speed.",
            "rows": t03_rows,
        },
        "504-04": {
            "title": "CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WORK ZONES",
            "confidence": "verbatim",
            "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
            "note": "Long-term variant of 402-05. No transverse-bumps row. Columns include temporary cones/drums.",
            "columnHeaders": [c[0] for c in CHAN_COLS],
            "provisionRows": t04_provisions,
            "spacingRows": t04_spacing,
            "tableNotes": [
                "NOTES: X= ALLOWED, BLANK = NOT ALLOWED, O = OPTIONAL",
                "1. - A TYPE 1 OBJECT MARKER MAY BE USED IN LIEU OF CHANNELIZING DEVICE.",
                "2. - CHANNELIZING DEVICES SHALL BE EQUIPPED WITH A FLASHING WARNING LIGHT.",
            ],
        },
        "504-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "note": "8 entries. NYR9-11 replaces NYW8-33 from 402-06. W20-5 (not W20-5R).",
            "rows": t05_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": "18 entries: notes 1-7 plus N1-N11 nighttime block (page 1, x>=900).",
    },
    "findings": findings,
    "comparisonVs402": {
        "504-02_vs_402-03": diffs_02 if diffs_02 else "identical_all_cells",
        "504-01_vs_402-04": diffs_01 if diffs_01 else "identical_all_cells",
        "504-05_vs_402-06": diffs_05 if diffs_05 else "shared_signs_match; NYR9-11_new; NYW8-33_absent",
        "absent_vs_402": ["402-01 protective vehicle", "402-02 roll ahead distance"],
        "new_vs_402": ["504-03 positive barrier flare rates"],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
print("Row counts:")
for tid, tbl in draft["tables"].items():
    n = len(tbl.get("rows", tbl.get("provisionRows", [])))
    print(f"  {tid}: {n} rows")
print("Notes:", len(notes_printed))
print("Findings:", *findings[:8], sep="\n  ")
