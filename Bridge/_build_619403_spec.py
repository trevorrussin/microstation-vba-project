"""Build 619-403.json = 303 corridor (two-lane) + 402 intermediate extras + draft tables."""
from __future__ import annotations

import copy
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
base = json.loads((ROOT / "Data/sheet-specs/619-303.json").read_text(encoding="utf-8"))
ref402 = json.loads((ROOT / "Data/sheet-specs/619-402.json").read_text(encoding="utf-8"))
draft = json.loads((ROOT / "Data/sheet-specs/_draft_619403_tables.json").read_text(encoding="utf-8"))
s = copy.deepcopy(base)

s["sheet"].update({
    "number": "619-403",
    "title": "WORK ZONE TRAFFIC CONTROL MULTI-LANE DIVIDED ROADWAY AND FREEWAY LEFT (OR RIGHT) TWO LANE CLOSURE",
    "operation": "INTERMEDIATE TERM OPERATION",
    "approved": "2024-04-15",
    "sourceUrl": "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-403_E1_0.pdf",
    "localPdf": "Bridge/captures/619-403.pdf",
    "localRender": None,
    "pdfPages": 2,
    "transcribedBy": "Cursor (two-lane corridor from 303 + intermediate extras from 402; tables from _draft_619403_tables.json)",
    "transcribedOn": "2026-08-03",
    "provenanceNote": (
        "Hybrid of 619-303 (dual merging taper + 2L + W20-5a) and 619-402 (PVH/PVL, 20' spacing, "
        "channelizing matrix, regulatory signs). Table numbering trap: 04=channelizing, 05=sizes, "
        "06=advance spacing. FREEWAY row of 403-06: A/B=1000/1500 found in text; C=2640 not in "
        "text layer — filled from identical 402-04/011-06 with knownAnomaly flag."
    ),
})
s["applicability"]["duration"] = "Intermediate Term"
s["applicability"]["durationDefinition"] = ref402["applicability"]["durationDefinition"]
s["applicability"]["closure"] = "Left (or right) two-lane closure"
s["applicability"]["closureNote"] = (
    "Note 2: right two-lane closures are symmetrical — substitute W20-5aR and W4-2R "
    "(sheet title is LEFT OR RIGHT; plan shows both)."
)

s["tableRoles"] = draft["tableRoles"]
s["tables"] = draft["tables"]

# Ensure FREEWAY advance-warning row exists (text layer incomplete)
aw = s["tables"]["403-06"]
roads = {r["roadType"] for r in aw["rows"]}
if "FREEWAY" not in roads:
    aw["rows"].append({
        "roadType": "FREEWAY",
        "minMph": None,
        "maxMph": None,
        "A": 1000, "B": 1500, "C": 2640,
        "XX": "1 MILE", "YY": "½ MILE",
        "confidence": "inferred",
        "note": "A=1000 and B=1500 confirmed in PDF text near FREEWAY label; C=2640 and mile legends "
                "not present as '2640' tokens on this revision's text layer — values taken from "
                "identical 402-04 / 011-06 FREEWAY row. Visually confirm before trusting C.",
    })
    aw.setdefault("knownAnomalies", []).append({
        "cell": "FREEWAY.C",
        "printed": None,
        "issue": "No '2640' token anywhere in 619-403.pdf text layer; A/B 1000/1500 present.",
        "recommendation": "Visually confirm C against the printed table; until then treat as 402-04's 2640.",
    })

# Remap corridor table refs 303-* -> 403-*
remap = {
    "303-01": "403-01", "303-02": "403-02", "303-03": "403-06",
    "303-04": "403-03", "303-05": "403-05",
}
def remap_str(x: str) -> str:
    for a, b in remap.items():
        x = x.replace(a, b)
    return x

for z in s["corridor"]["zones"]:
    ls = z.get("lengthSource") or {}
    if isinstance(ls, dict) and ls.get("table"):
        ls["table"] = remap_str(ls["table"])
    if z.get("sheetReference"):
        z["sheetReference"] = remap_str(z["sheetReference"])
    # Sign code on sheet table is W20-5a (not W20-5aR)
    if z.get("signCode") == "W20-5aR":
        z["signCode"] = "W20-5a"

for al in s["orderTable"]["alignments"]:
    for r in al["rows"]:
        if r.get("signCode") == "W20-5aR":
            r["signCode"] = "W20-5a"

for item in s["signs"]["items"]:
    if item["signCode"] == "W20-5aR":
        item["signCode"] = "W20-5a"
        item["sheetNote"] = "Table 403-05 prints W20-5a; SignLibrary base remains W20-05aR for right closure."
    sub = item.get("legendSubstitution")
    if sub and sub.get("table"):
        sub["table"] = remap_str(sub["table"])

# Intermediate extras from 402
s["signs"]["items"] = [i for i in s["signs"]["items"]
                       if i["signCode"] not in ("R2-1 OR NYR2-2", "NYR2-6", "W4-2L", "NYR9-11")]
# Add size-table signs that 303 didn't have
extra_codes = {r["signCode"] for r in s["tables"]["403-05"]["rows"]}
have = {i["signCode"] for i in s["signs"]["items"]}
for row in s["tables"]["403-05"]["rows"]:
    code = row["signCode"]
    if code in have:
        continue
    if code == "WARNING FLAG":
        continue
    s["signs"]["items"].append({
        "signCode": code,
        "sheetLegend": None,
        "legendSubstitution": None,
        "shape": "rectangle" if code.startswith(("R", "NYR", "G")) else "diamond",
        "postMounted": True,
        "corridorZone": None,
        "sizeNonFreeway": row.get("NON-FREEWAY"),
        "sizeFreeway": row.get("FREEWAY"),
        "signLibraryKey": "R2-1" if code.startswith("R2") else None,
        "required": code.startswith("R2"),
        "note": "From Table 403-05 intermediate extras / left-closure symmetry.",
    })
    have.add(code)

# Ensure WARNING FLAG present
if "WARNING FLAG" not in have:
    s["signs"]["items"].append({
        "signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
        "mountedOn": "W20-1, W4-2R", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
    })

for sym in s["symbols"]["items"]:
    if sym.get("id") == "channelizingDevices":
        sym["longitudinalSpacing"] = {
            "maxFt": 20,
            "sheetText": "Note 4 / Table 403-04 — 20' max in active work space (intermediate).",
        }
        for run in sym.get("runs", []):
            dcs = run.get("deviceCountSource")
            if dcs and dcs.get("table"):
                dcs["table"] = remap_str(dcs["table"])

for d in s["annotations"]["dimensions"]:
    if d.get("reference"):
        d["reference"] = remap_str(d["reference"])

s["details"] = {"403A": {"title": "DETAIL 403A", "note": "Transverse channelizing / shoulder detail."}}

printed = [n for n in draft["notes"]["printed"] if not n.startswith("N")]
s["notes"] = {
    "confidence": "verbatim",
    "printed": printed,
    "planCallouts": s.get("notes", {}).get("planCallouts", []),
    "tableNotes": draft["tables"]["403-01"].get("tableNotes", []),
}

# Inputs exposure from 402
for inp in s["inputs"]:
    inp["usedBy"] = [remap_str(u) for u in inp.get("usedBy", [])]
    if inp["id"] == "exposureCondition":
        inp["allowed"] = [
            "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
            "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS)",
        ]

s["rules"] = [
    {"id": "two-merging-tapers-plus-2L", "severity": "error", "source": "Plan layout",
     "assert": "Upstream walk has two MERGING TAPER rows separated by 2L (= 2×L).",
     "commonFailure": "Using 402's single-taper walk."},
    {"id": "sign-is-two-lane", "severity": "error", "source": "Table 403-05",
     "assert": "Mid advance sign is W20-5a (SignLibrary W20-05aR*).",
     "commonFailure": "Emitting one-lane W20-05R* keys."},
    {"id": "device-spacing-20ft", "severity": "error", "source": "Note 4",
     "assert": "Active work space channelizing spacing <= 20 ft.",
     "commonFailure": "Copying 40 ft from short-term sheets."},
    {"id": "shoulder-taper-is-an-overlay", "severity": "error", "source": "Dimension datums",
     "assert": "Shoulder taper overlays gap A.",
     "commonFailure": "Sequential shoulder-taper station."},
    {"id": "regulatory-speed-mid-AB", "severity": "error", "source": "Note 8",
     "assert": "R2-1 or NYR2-* halfway between 1st and 2nd advance warning signs.",
     "commonFailure": "Omitting regulatory speed because short-term sheets lack it."},
    {"id": "pvh-pvl-codes", "severity": "warning", "source": "Table 403-01",
     "assert": "Protective vehicle codes are PVH/PVL+TMIA.",
     "commonFailure": "Using short-term P/TMIA codes."},
]

s["knownCodeDeviations"] = [
    {"id": "device-spacing-default-40", "severity": "error",
     "assert": "Placement still defaults to 40 ft; intermediate requires 20 ft."},
    {"id": "dual-taper-placement", "severity": "error",
     "assert": "Needs two taper runs + 2L + two arrow panels like 303."},
    {"id": "freeway-C-text-layer", "severity": "warning",
     "assert": "403-06 FREEWAY.C filled from 402-04; confirm visually on printed sheet."},
]

s["knownExcerpts"] = {
    "from619-303": ["Dual merging taper + 2L corridor", "W20-5a two-lane mid sign"],
    "from619-402": ["PVH/PVL table", "20' spacing", "channelizing matrix", "regulatory signs"],
}

out = ROOT / "Data/sheet-specs/619-403.json"
out.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("wrote", out, "signs", [i["signCode"] for i in s["signs"]["items"]])
