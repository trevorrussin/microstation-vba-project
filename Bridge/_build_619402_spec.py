"""Build Data/sheet-specs/619-402.json from 619-302 + _draft_619402_tables.json.

Intermediate-term one-lane Family 2 sibling. Corridor skeleton matches 302
(single MERGING TAPER + shoulder taper overlay on gap A). Genuine diffs:
2-page sheet, PVH/PVL protective-vehicle codes, Table 402-05 channelizing
matrix, Note 4 = 20' device spacing, regulatory speed mid A/B, NY9-11.
"""
from __future__ import annotations

import copy
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
ref = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
draft = json.loads((ROOT / "Data/sheet-specs/_draft_619402_tables.json").read_text(encoding="utf-8"))
s = copy.deepcopy(ref)

s["sheet"] = {
    "number": "619-402",
    "title": "WORK ZONE TRAFFIC CONTROL MULTI-LANE DIVIDED ROADWAY AND FREEWAY RIGHT LANE CLOSURE",
    "series": "WORK ZONE TRAFFIC CONTROL",
    "operation": "INTERMEDIATE TERM OPERATION",
    "units": "U.S. CUSTOMARY",
    "scale": "NOT TO SCALE",
    "approved": "2026-04-29",
    "issuedUnder": "EI / EB per sheet title block (E3 revision)",
    "signedBy": None,
    "sourceUrl": "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-402_E3.pdf",
    "localPdf": "Bridge/captures/619-402.pdf",
    "localRender": None,
    "pdfPages": 2,
    "transcribedBy": "Cursor (Family 2 intermediate sibling of 619-302; tables via _draft_619402_tables.json; corridor cloned from 302 with plan-geometry datum check)",
    "transcribedOn": "2026-08-03",
    "provenanceNote": (
        "Canonical PDF is 619-402_E3.pdf (copied to Bridge/captures/619-402.pdf). "
        "Page 1 = plan + Notes 1-8; page 2 = Tables 402-01..06. "
        "402-03/04/02 match 302-02/03/05 cell-for-cell including 65mph/12ft=800/20/21. "
        "402-01 uses PVH/PVL+TMIA (not 302's P/TMIA). 402-05 and regulatory signs are new."
    ),
}

s["applicability"]["duration"] = "Intermediate Term"
s["applicability"]["durationDefinition"] = (
    "Stationary work that occupies a location more than one daylight period up to "
    "3 consecutive days, or nighttime work lasting more than 1 hour (Note 1)."
)
s["applicability"]["closure"] = "Right lane closure"
s["applicability"]["closureNote"] = (
    "Note 2: left lane closures are symmetrical — substitute W20-5(L) and W4-2L."
)

s["tableRoles"] = draft["tableRoles"]
s["tables"] = draft["tables"]

# Fix knownAnomalies wording for this sheet
if "knownAnomalies" in s["tables"].get("402-03", {}):
    for a in s["tables"]["402-03"]["knownAnomalies"]:
        a["note"] = a.get("issue", a.get("note", "")).replace("619-302", "619-402").replace("302-02", "402-03")
        if "recommendation" in a:
            a["recommendation"] = a["recommendation"].replace("619-302", "619-402")

# inputs: remap table ids + exposure wording from 402-01
for inp in s["inputs"]:
    inp["usedBy"] = [
        u.replace("302-01", "402-01")
         .replace("302-02", "402-03")
         .replace("302-03", "402-04")
         .replace("302-04", "402-06")
         .replace("302-05", "402-02")
        for u in inp.get("usedBy", [])
    ]
    if inp["id"] == "exposureCondition":
        inp["allowed"] = [
            "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
            "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS)",
        ]

# Corridor: same skeleton as 302, table refs remapped
for z in s["corridor"]["zones"]:
    ls = z.get("lengthSource") or {}
    if isinstance(ls, dict) and "table" in ls:
        ls["table"] = (
            ls["table"]
            .replace("302-02", "402-03")
            .replace("302-03", "402-04")
            .replace("302-05", "402-02")
            .replace("302-04", "402-06")
        )
    if z.get("sheetReference"):
        z["sheetReference"] = (
            z["sheetReference"]
            .replace("302-02", "402-03")
            .replace("302-03", "402-04")
            .replace("302-05", "402-02")
        )
s["corridor"]["confidence"] = "drawing"
s["corridor"]["description"] = (
    "Same sequential skeleton as 619-302 (single MERGING TAPER L; shoulder taper "
    "overlays gap A). Confirmed via extract_plan_geometry on page 1 of E3 PDF. "
    "Intermediate-term extras (regulatory speed mid A/B, NY9-11, 20' device spacing) "
    "are symbols/rules, not corridor zones."
)

# Remap sign table refs
for item in s["signs"]["items"]:
    sub = item.get("legendSubstitution")
    if sub and sub.get("table"):
        sub["table"] = sub["table"].replace("302-03", "402-04")
    for k in ("sizeNonFreeway", "sizeFreeway"):
        pass
# Align sign codes with Table 402-06 (sheet prints W20-5, not W20-5R)
for item in s["signs"]["items"]:
    if item["signCode"] == "W20-5R":
        item["signCode"] = "W20-5"
        item["sheetNote"] = "Sheet table 402-06 prints W20-5; SignLibrary base remains W20-05R (right)."
for z in s["corridor"]["zones"]:
    if z.get("signCode") == "W20-5R":
        z["signCode"] = "W20-5"
for al in s["orderTable"]["alignments"]:
    for r in al["rows"]:
        if r.get("signCode") == "W20-5R":
            r["signCode"] = "W20-5"

s["signs"]["note"] = (
    "Base advance set matches 302 (W20-1 / W20-5 / W4-2R / G20-2). Intermediate adds "
    "regulatory R2-1 or NYR2-2/NYR2-6 per Note 8 / Table 402-06. NY9-11 (Note 7) is "
    "recommended only — tracked under symbols, not signs.items (no size row on 402-06)."
)
# Replace previous extras with codes that match 402-06 exactly
s["signs"]["items"] = [i for i in s["signs"]["items"]
                       if i["signCode"] not in ("R2-1", "NY9-11", "R2-1 OR NYR2-2", "NYR2-6")]
s["signs"]["items"].extend([
    {
        "signCode": "R2-1 OR NYR2-2",
        "sheetLegend": "SPEED LIMIT",
        "legendSubstitution": None,
        "shape": "rectangle",
        "warningFlags": False,
        "postMounted": True,
        "corridorZone": None,
        "positionNote": "Note 8: required halfway between 1st and 2nd advance warning signs unless already present.",
        "sizeNonFreeway": "30x36",
        "sizeFreeway": "36x48",
        "signLibraryKey": "R2-1",
        "required": True,
    },
    {
        "signCode": "NYR2-6",
        "sheetLegend": "SPEED LIMIT (variant)",
        "legendSubstitution": None,
        "shape": "rectangle",
        "postMounted": True,
        "corridorZone": None,
        "sizeNonFreeway": None,
        "sizeFreeway": None,
        "signLibraryKey": None,
        "required": False,
        "note": "Listed on 402-06 under the R2-1 group; no explicit size cells in the PDF text layer.",
    },
])

s["symbols"]["items"].append({
    "id": "ny911Recommended",
    "sheetLabel": "NY9-11",
    "required": False,
    "note": "Note 7 recommended: 1000' before first advance warning when speed >= 45; 300-500' when < 45.",
})

# Channelizing spacing from Note 4 / Table 402-05
for sym in s["symbols"]["items"]:
    if sym.get("id") == "channelizingDevices":
        sym["longitudinalSpacing"] = {
            "maxFt": 20,
            "sheetText": "Note 4 / Table 402-05 — not to exceed 20' in the active work space (intermediate).",
        }
        for run in sym.get("runs", []):
            dcs = run.get("deviceCountSource")
            if dcs and dcs.get("table"):
                dcs["table"] = dcs["table"].replace("302-02", "402-03")

s["symbols"]["items"].append({
    "id": "channelizingApplicationTable",
    "sheetLabel": "TABLE 402-05",
    "required": True,
    "note": "Device type matrix for intermediate-term (cones/drums/tubular/etc). See tables['402-05'].",
})

# Annotations table refs
for d in s["annotations"]["dimensions"]:
    if d.get("reference"):
        d["reference"] = (
            d["reference"]
            .replace("302-02", "402-03")
            .replace("302-03", "402-04")
            .replace("302-05", "402-02")
        )

s["details"] = {
    "402A": {
        "title": "DETAIL 402A",
        "note": "Referenced from plan (transverse channelizing / shoulder closure detail).",
    }
}

# Notes: use draft but keep only Notes 1-8 (drop N1-N8 night/misc if present)
printed = []
for n in draft["notes"]["printed"]:
    if n.startswith("N"):
        continue
    printed.append(n)
s["notes"] = {
    "confidence": "verbatim",
    "printed": printed,
    "planCallouts": [
        {"text": "THIS SIGN SHALL BE LOCATED A MINIMUM DISTANCE OF 80 FT AND MAXIMUM OF 400 FT PAST THE END OF THE DOWNSTREAM TAPER.",
         "appliesTo": "G20-2"},
    ],
    "tableNotes": draft["tables"]["402-01"].get("tableNotes", []),
}

s["rules"] = [
    {"id": "no-occupancy-buffer", "severity": "error", "source": "Note 5",
     "assert": "No work activity, equipment, vehicles, or material in the buffer space at any time.",
     "commonFailure": "Drawing work-area hatch into the buffer."},
    {"id": "sign-order", "severity": "error", "source": "Plan layout",
     "assert": "Upstream advance signs: W4-2R, then W20-5R, then W20-1 (same as 302).",
     "commonFailure": "Reversing the order or using W20-5aR from the two-lane sibling."},
    {"id": "shoulder-taper-is-an-overlay", "severity": "error", "source": "Dimension datums",
     "assert": "Shoulder taper overlays gap A; consumes no station.",
     "commonFailure": "Sequential shoulder-taper station row."},
    {"id": "device-spacing-20ft", "severity": "error", "source": "Note 4 / Table 402-05",
     "assert": "Channelizing device spacing in the active work space must not exceed 20 ft (intermediate), not 302's 40 ft.",
     "commonFailure": "Copying 40 ft from the short-term Family 2 reference."},
    {"id": "regulatory-speed-mid-AB", "severity": "error", "source": "Note 8 / Table 402-06",
     "assert": "Place R2-1 or NYR2-2/NYR2-6 halfway between 1st and 2nd advance warning signs unless already present.",
     "commonFailure": "Omitting the regulatory speed sign because 302 has none."},
    {"id": "no-invented-zones", "severity": "error", "source": "Plan layout",
     "assert": "No Vehicle Space / temp barrier / box-corr beam sequential rows on the basic plan.",
     "commonFailure": "Emitting the generic 7-row default upstream table."},
    {"id": "pvh-pvl-codes", "severity": "warning", "source": "Table 402-01",
     "assert": "Protective vehicle lookup returns PVH+TMIA / PVL+TMIA / SEE NOTE n — not 302's bare P.",
     "commonFailure": "Assuming short-term P/TMIA codes apply to intermediate sheets."},
]

s["knownCodeDeviations"] = [
    {"id": "device-spacing-default-40", "severity": "error",
     "assert": "Current placement defaults assume 40 ft skip-line spacing; intermediate sheets require 20 ft in the work space per Note 4."},
    {"id": "regulatory-sign-not-in-order-table", "severity": "warning",
     "assert": "R2-1 mid A/B is required by Note 8 but is not yet a sequential order-table row in the bridge payload."},
    {"id": "sheet-registry-wrong-signs", "severity": "warning",
     "assert": "sheet-registry.tsv must not override this spec; intermediate sheets DO include R2-1/NYR2-* (unlike 302 where those registry codes were noise)."},
]

s["knownExcerpts"] = {
    "from619-302": [
        "402-03 == 302-02 (taper+buffer) including 65/12 = 800/20/21",
        "402-04 == 302-03 (advance warning)",
        "402-02 == 302-05 (roll ahead)",
        "Corridor skeleton (single merging taper + shoulder overlay) matches 302",
    ],
    "differsFrom302": [
        "Duration Intermediate; Notes 1/4/5/7/8 differ; Note 4 = 20' spacing",
        "402-01 PVH/PVL+TMIA vs 302 P/TMIA",
        "New Table 402-05 channelizing application matrix",
        "402-06 adds R2-1/NYR2-2/NYR2-6",
        "2 PDF pages (tables on page 2)",
    ],
}

out = ROOT / "Data/sheet-specs/619-402.json"
out.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("wrote", out, "notes", len(printed), "tables", list(s["tables"]))
