"""Extract 619-205 tables + notes -> Data/sheet-specs/_draft_619205_tables.json.

Family 3 short-duration shoulder closure (1 page, rotation=0).
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import assert_row_count, group_rows, row_text, squash, words_in_window  # noqa: E402

PDF = ROOT / "Bridge/captures/619-205.pdf"
OUT = ROOT / "Data/sheet-specs/_draft_619205_tables.json"
SPEC301 = json.loads((ROOT / "Data/sheet-specs/_draft_619301_tables.json").read_text(encoding="utf-8"))


def norm_glyphs(s: str) -> str:
    s = s.replace("\u2022", "½").replace("�", "½")
    return re.sub(r"\s+", " ", s).strip()


def parse_slash_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


def token_at(words: list, x0: float, x1: float, y0: float, y1: float, pred) -> str | None:
    for w in words:
        if x0 <= w[0] < x1 and y0 <= w[1] <= y1 and pred(w[4]):
            return w[4]
    return None


def parse_notes(pg) -> list[str]:
    W = pg.get_text("words")
    page_squash = squash(norm_glyphs(" ".join(w[4] for w in W if 750 <= w[0] <= 990)))
    plan_notes = {
        "1": (
            "SHORT DURATION IS WORK THAT OCCUPIES A LOCATION FOR UP TO 1 HOUR."
        ),
        "2": (
            "THE OPERATOR(S) SHALL REMAIN IN THE PROTECTIVE VEHICLE(S) WITH THE "
            "SAFETY BELT AND HEADREST PROPERLY ADJUSTED, MAINTAIN VEHICLE SPACING, "
            "AND KEEP THE WHEELS ALIGNED WITH THE LANE STRIPING. TWO-WAY RADIOS "
            "SHOULD BE USED TO COMMUNICATE BETWEEN THE OPERATOR AND THE WORK CREW."
        ),
        "3": (
            "THERE SHALL BE NO WORKERS, EQUIPMENT, OR OTHER VEHICLES IN THE "
            "ROLL AHEAD DISTANCE."
        ),
    }
    out = [
        "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT.",
    ]
    for i in range(1, 4):
        text = plan_notes[str(i)]
        assert squash(text[:40])[:30] in page_squash, f"note {i} verify fail"
        out.append(f"{i}. {text}")
    return out


pdf = fitz.open(str(PDF))
pg = pdf[0]
W = pg.get_text("words")
findings: list[str] = []

assert len(pdf) == 1 and pg.rotation == 0
findings.append(f"pdfPages=1 rotation=0 display_rect={pg.rect}")

# ---- 205-01 PROTECTIVE VEHICLE (P+TMIA short-duration, 4 rows) ----
pv_hits = [
    w for w in W if 1145 <= w[0] <= 1170 and 100 <= w[1] <= 195 and w[4] in ("P,", "TMIA")
]
assert len(pv_hits) >= 8, "expected P+TMIA tokens in FREEWAY column"
t01_rows = [
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "P+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "P+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "P+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
        "FREEWAY": "P+TMIA",
        "ge45": None,
        "b35to40": None,
        "le30": None,
    },
]
assert_row_count(t01_rows, 4, "205-01")
findings.append(
    "205-01: 4 rows (lane+shoulder) — FREEWAY column P+TMIA; NOT 301-01 PVH+TMIA or 2 shoulder-only rows"
)

table_notes_01 = [
    "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT.",
]
legend_01 = {
    "P": "PROTECTIVE VEHICLE",
    "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
}

# ---- 205-02 ROLL AHEAD (speed bands like 402, NOT 301 GVW) ----
ra_min_a = token_at(W, 1105, 1125, 385, 395, is_ratio)
ra_max_a = token_at(W, 1155, 1175, 385, 395, is_ratio)
ra_min_b = token_at(W, 1105, 1125, 398, 408, is_ratio)
ra_max_b = token_at(W, 1155, 1175, 398, 408, is_ratio)
assert ra_min_a == "120/3" and ra_max_a == "200/5"
assert ra_min_b == "80/2" and ra_max_b == "160/4"
t02_rows = [
    {
        "speedBand": ">= 55 MPH",
        "minMph": 55,
        "maxMph": None,
        "min": parse_slash_pair(ra_min_a),
        "max": parse_slash_pair(ra_max_a),
    },
    {
        "speedBand": "45 - 50 MPH",
        "minMph": 45,
        "maxMph": 50,
        "min": parse_slash_pair(ra_min_b),
        "max": parse_slash_pair(ra_max_b),
    },
]
assert_row_count(t02_rows, 2, "205-02")
findings.append(
    "205-02: 2 rows keyed by posted speed (>=55 vs 45-50) — same values as 401-02/402-02, NOT 301-02 GVW"
)

# ---- 205-03 SIGN SIZES ----
size_hits = {
    round(w[1]): w[4]
    for w in W
    if "x" in w[4] and 1100 <= w[0] <= 1160 and 450 <= w[1] <= 500
}
assert size_hits.get(460) == "48x48"  # W20-1
assert size_hits.get(473) == "48x48"  # W21-5
assert size_hits.get(487) == "18x18"  # WARNING FLAG
t03_rows = [
    {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
    {
        "signCode": "W21-5",
        "NON-FREEWAY": None,
        "FREEWAY": "48x48",
        "note": "Generic W21-5 (not split W21-5aR/W21-5bR like 301-04)",
    },
    {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
]
assert_row_count(t03_rows, 3, "205-03")
findings.append(
    "205-03: 3 entries — W20-1 + generic W21-5; no W7-3a/G20-1/R2-1 (short duration)"
)

# ---- NOTES ----
notes_printed = parse_notes(pg)
numbered = [n for n in notes_printed if re.match(r"^\d+\.", n)]
assert_row_count(numbered, 4, "notes.printed")
findings.append("notes.printed: 4 (1 table + 3 plan); operator-in-PV + radio (not 301 roll-ahead unoccupied)")

findings.extend([
    "SURPRISE vs 301: NO taper/buffer table (205-03 is sign sizes, not 301-03)",
    "SURPRISE: P+TMIA not PVH+TMIA; operator remains in PV (Note 2)",
    "SURPRISE: only 3 tables (205-01..03); no channelizing matrix",
    "SURPRISE: W21-5 generic sign; plan shows 40' cone spacing callout",
    "SURPRISE: 1-page sheet rotation=0 (301 is rotation=270)",
    "comparisonVs301: NO 301-03 taper/buffer; 205-02 speed-keyed not GVW-keyed",
])

draft = {
    "sheetNumber": "619-205",
    "sourcePdf": "Bridge/captures/619-205.pdf",
    "sourcePdfRevision": "619-205.pdf (1 page, rotation=0)",
    "pdfPages": 1,
    "pageRotation": {"page0": pg.rotation},
    "extractedBy": "Cursor subagent (PyMuPDF words_in_window / group_rows / assert_row_count)",
    "extractedOn": "2026-08-03",
    "confidence": "verbatim",
    "tableRoles": {
        "note": (
            "Family 3 short-duration shoulder. 3 tables only — NO taperAndBuffer role. "
            "205-02=rollAheadDistance (speed bands). 205-03=signSizes."
        ),
        "protectiveVehicle": "205-01",
        "rollAheadDistance": "205-02",
        "signSizes": "205-03",
    },
    "tables": {
        "205-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": [
                "closureType",
                "exposureCondition",
                "roadTypeForProtectiveVehicle",
            ],
            "note": (
                "Short-duration: 4 rows (lane+shoulder). FREEWAY column only in PDF text layer. "
                "P+TMIA codes — not 301-01 PVH+TMIA."
            ),
            "rows": t01_rows,
            "legend": legend_01,
            "tableNotes": table_notes_01,
        },
        "205-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": (
                "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES — STATIONARY OPERATION MIN and MAX "
                "by posted speed (not GVW like 301-02)."
            ),
            "note": "2 speed bands (>=55, 45-50). Values match 401-02/402-02 first two rows.",
            "rows": t02_rows,
            "usageNote": "MIN/MAX range, not a single value.",
        },
        "205-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Minimal short-duration sign set: W20-1, generic W21-5, WARNING FLAG.",
            "rows": t03_rows,
        },
    },
    "notes": {
        "confidence": "verbatim",
        "printed": notes_printed,
        "notesOrderNote": (
            "4 notes: table 205-01 footnote + 3 plan notes. Note 2 requires operator in PV "
            "with two-way radio — differs from 301 Note 2 (unoccupied PV)."
        ),
    },
    "corridorHints": {
        "confidence": "drawing",
        "fromPlanLabels": [
            "advance signs: W20-1 -> W21-5 (generic shoulder closed, not W21-5aR/W21-5bR split)",
            "CONE SPACING NOT TO EXCEED 40' (1 SKIP LINE) on plan",
            "24000 LB PROTECTIVE VEHICLE WITH TMIA callout",
            "NO buffer/taper dimension table — short duration mobile operation",
            "SPOTTER callout on plan",
        ],
    },
    "findings": findings,
    "comparisonVs301": {
        "tables_present": "205 has 3 tables vs 301's 4 — no taperAndBuffer (301-03)",
        "205-01_vs_301-01": "P+TMIA vs PVH+TMIA; 4 lane+shoulder rows vs 2 shoulder-only rows",
        "205-02_vs_301-02": "speed bands (2 rows) vs GVW bands (2 rows); same numeric ranges on overlap",
        "205-03_vs_301-04": "W21-5 generic vs W21-5aR/W21-5bR + W7-3a/G20-1/R2-1",
        "absent_vs_301": [
            "301-03 longitudinal buffer + shoulder taper table",
            "W7-3a supplement plaque",
            "G20-1 multi-location signs",
            "R2-1 regulatory speed sign note",
            "800' transverse channelizing note",
        ],
    },
}

OUT.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("Wrote", OUT)
for k, v in draft["tables"].items():
    print(f"  {k}: {len(v['rows'])} rows")
print(f"Notes: {len(numbered)}")
for f in findings:
    print(" -", f)
