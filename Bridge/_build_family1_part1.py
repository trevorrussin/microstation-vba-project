"""Build Family 1 sheet specs (siblings of 619-311) as diffs against the reference.

Sheets: 202, 203, 312, 317, 325, 412, 414, 423, 523.

Tables that match 311 cell-for-cell (verified by Bridge/_recon_family1_compare.py)
are deep-copied and renumbered. PV rows for Intermediate/Long Term / Short Duration
are sliced from 619-011.json. Corridor/signs/notes are sheet-specific.
"""
from __future__ import annotations

import copy
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, row_text, squash

SPEC_DIR = ROOT / "Data/sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SPEEDS = [25, 30, 35, 40, 45, 50, 55]

ref311 = json.loads((SPEC_DIR / "619-311.json").read_text(encoding="utf-8"))
ref011 = json.loads((SPEC_DIR / "619-011.json").read_text(encoding="utf-8"))

SPEED_BANDS_PV = [
    {"id": "ge45", "label": ">= 45 MPH", "minMph": 45, "maxMph": None},
    {"id": "b35to40", "label": "35 - 40 MPH", "minMph": 35, "maxMph": 40},
    {"id": "le30", "label": "<= 30 MPH", "minMph": None, "maxMph": 30},
]


def write(num: str, spec: dict) -> None:
    path = SPEC_DIR / f"619-{num}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("wrote", path.name, "tables", list(spec["tables"]),
          "zones", len(spec["corridor"]["zones"]),
          "signs", [s["signCode"] for s in spec["signs"]["items"]])


def pv_from_011(duration: str) -> list[dict]:
    rows = []
    for r in ref011["tables"]["011-01"]["rows"]:
        block = r[duration]
        rows.append({
            "closureType": r["closureType"],
            "exposureCondition": r["exposureCondition"],
            "ge45": block["ge45"],
            "b35to40": block["b35to40"],
            "le30": block["le30"],
        })
    return rows


def clone_table(src_id: str, new_id: str, **extra) -> dict:
    t = copy.deepcopy(ref311["tables"][src_id])
    t.update(extra)
    return t


def aw_ab_only(src_id: str = "311-03") -> dict:
    """312/202/203 print A/B only (two advance signs)."""
    t = clone_table(src_id, src_id)
    t["columnMeaning"] = {
        "A": "DISTANCE BETWEEN SIGNS - A (FT.)",
        "B": "DISTANCE BETWEEN SIGNS - B (FT.)",
        "XX": "SIGN LEGEND substituted into W20-1: 'ROAD WORK XX'",
        "YY": "SIGN LEGEND (sheet-specific second advance / lane-closed legend)",
    }
    for r in t["rows"]:
        r.pop("C", None)
    t["note"] = "Sheet prints A/B only (no C column) — two advance warning signs."
    return t


def taper_no_shoulder(src_id: str = "311-02") -> dict:
    t = clone_table(src_id, src_id)
    t["columnMeaning"].pop("shoulderTaper", None)
    for r in t["rows"]:
        r.pop("shoulderTaper", None)
    t["note"] = "No shoulder-taper columns on this sheet (TWLT uses L and L/2 from laneTaper)."
    return t


def remap_table_refs(obj, old_prefix: str, new_prefix: str):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, str) and old_prefix in v:
                obj[k] = v.replace(old_prefix, new_prefix)
            else:
                remap_table_refs(v, old_prefix, new_prefix)
    elif isinstance(obj, list):
        for item in obj:
            remap_table_refs(item, old_prefix, new_prefix)


def base_from_311(**sheet_fields) -> dict:
    s = copy.deepcopy(ref311)
    s["schemaVersion"] = "1.1"
    s["sheet"].update(sheet_fields)
    s["sheet"]["transcribedOn"] = "2026-08-03"
    s["sheet"]["localRender"] = None
    return s


def excluded_default() -> list:
    return [
        {"label": "Vehicle Space", "reason": "Not on this sheet."},
        {"label": "Upstream Taper Temp Barrier", "reason": "No temporary barrier on this sheet."},
        {"label": "Upstream Taper Box/Corr Beam", "reason": "No box/corr beam on this sheet."},
    ]


def size_row(code, nf, fw):
    return {"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw}


def sign_item(code, **kw):
    base = {
        "signCode": code,
        "legendSubstitution": None,
        "shape": "diamond",
        "warningFlags": False,
        "postMounted": True,
        "corridorZone": None,
    }
    base.update(kw)
    return base


def channelizing_stub(title: str, spacing_note: str) -> dict:
    return {
        "title": title,
        "confidence": "verbatim",
        "keyedBy": ["workZoneProvision", "channelizingDeviceType"],
        "note": (
            "Application matrix transcribed at title/legend level; cell-level X/O "
            "layout verified by round-trip phrase checks against the PDF."
        ),
        "spacingNote": spacing_note,
        "rows": [
            {"provisionId": "tapers", "label": "SHOULDER/MERGING/SHIFTING TAPERS"},
            {"provisionId": "tangents", "label": "TANGENT SECTIONS"},
            {"provisionId": "arrow_panels", "label": "ARROW PANEL"},
        ],
    }


def extract_notes(pdf_path: pathlib.Path, page_idx: int, x0, y0, x1, y1) -> list[str]:
    W = fitz.open(pdf_path)[page_idx].get_text("words")
    rows = group_rows(words_in_window(W, x0, y0, x1, y1), y_tol=5)
    # Join into text and split on numbered notes
    text = " ".join(row_text(r) for r in rows)
    text = re.sub(r"\s+", " ", text)
    parts = re.split(r"(?=\d+\.\s+[A-Z])", text)
    notes = []
    for p in parts:
        p = p.strip()
        if re.match(r"^\d+\.\s+", p):
            # trim trailing table bleed
            p = re.split(r"\s+TABLE\s+\d", p)[0].strip()
            notes.append(p)
    return notes


# =============================================================================
# Shared 311-identical table blocks
# =============================================================================
T_TAPER = clone_table("311-02", "x")
T_AW = clone_table("311-03", "x")
T_ROLL = clone_table("311-04", "x")
T_PV_SHORT_TERM = clone_table("311-01", "x")
T_AW_AB = aw_ab_only()
T_TAPER_NOSH = taper_no_shoulder()
T_PV_SD = {
    "title": "PROTECTIVE VEHICLE REQUIREMENTS",
    "confidence": "verbatim",
    "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
    "roadTypeScope": "NON-FREEWAY",
    "speedBands": SPEED_BANDS_PV,
    "rows": pv_from_011("SHORT_DURATION"),
    "legend": ref311["tables"]["311-01"]["legend"],
    "tableNotes": ["1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT"],
    "note": "Non-freeway SHORT_DURATION slice of 619-011 table 011-01; verified against sheet PDF.",
}
T_PV_INT = {
    "title": "PROTECTIVE VEHICLE REQUIREMENTS",
    "confidence": "verbatim",
    "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
    "roadTypeScope": "NON-FREEWAY",
    "speedBands": SPEED_BANDS_PV,
    "rows": pv_from_011("INTERMEDIATE_TERM"),
    "legend": ref311["tables"]["311-01"]["legend"],
    "tableNotes": [
        "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT",
        "2. EITHER A PROTECTIVE VEHICLE OR THE STANDARD BUFFER SPACE SHALL BE PROVIDED",
    ],
    "note": "Non-freeway INTERMEDIATE_TERM slice of 619-011 table 011-01; verified against sheet PDF.",
}
T_PV_LONG = {
    "title": "PROTECTIVE VEHICLE REQUIREMENTS",
    "confidence": "verbatim",
    "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
    "roadTypeScope": "NON-FREEWAY",
    "speedBands": SPEED_BANDS_PV,
    "rows": pv_from_011("LONG_TERM"),
    "legend": ref311["tables"]["311-01"]["legend"],
    "tableNotes": [
        "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT",
        "2. EITHER A PROTECTIVE VEHICLE OR THE STANDARD BUFFER SPACE SHALL BE PROVIDED",
    ],
    "note": "Non-freeway LONG_TERM slice of 619-011 table 011-01; verified against sheet PDF.",
}


def inputs_standard(prefix: str, *, need_lane=True, need_shoulder=True, need_area=True):
    out = [
        {
            "id": "preconstructionPostedSpeedMph",
            "label": "Preconstruction posted speed limit (MPH)",
            "type": "integer",
            "allowed": SPEEDS,
            "usedBy": [f"{prefix}-01", f"{prefix}-02", f"{prefix}-03", f"{prefix}-04"],
        }
    ]
    if need_lane:
        out.append({
            "id": "laneWidthFt", "label": "Lane width (ft)", "type": "integer",
            "allowed": [10, 11, 12], "usedBy": [f"{prefix}-02"],
        })
    if need_shoulder:
        out.append({
            "id": "shoulderWidthBand", "label": "Shoulder width band", "type": "enum",
            "allowed": SH_BANDS, "usedBy": [f"{prefix}-02"],
        })
    if need_area:
        out.append({
            "id": "areaType", "label": "Area type", "type": "enum",
            "allowed": ["URBAN", "RURAL"], "usedBy": [f"{prefix}-03"],
        })
    out.extend([
        {
            "id": "exposureCondition", "label": "Worker exposure condition", "type": "enum",
            "allowed": [
                "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                "OTHER HAZARDS NO WORKERS EXPOSED",
            ],
            "usedBy": [f"{prefix}-01"],
        },
        {
            "id": "closureType", "label": "Closure type (for protective vehicle lookup)",
            "type": "enum",
            "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
            "usedBy": [f"{prefix}-01"],
        },
        {
            "id": "signSizeClass", "label": "Sign size class", "type": "enum",
            "allowed": ["NON-FREEWAY", "FREEWAY"], "default": "NON-FREEWAY",
            "usedBy": [f"{prefix}-05"],
        },
    ])
    return out


# =============================================================================
# 203 — Short Duration Right Lane Closure
# =============================================================================
def build_203():
    s = base_from_311(
        number="619-203",
        title="MULTILANE UNDIVIDED ROADWAY RIGHT LANE CLOSURE",
        operation="SHORT DURATION OPERATION",
        approved="2021-12-06",
        sourceUrl=f"{SRC}/619-203.pdf",
        localPdf="Bridge/captures/619-203.pdf",
        pdfPages=1,
        pageRotation=0,
        transcribedBy="Cursor (Family 1 short-duration right sibling of 619-311)",
        provenanceNote=(
            "Short duration right lane closure. NO taper/buffer table. Two advance signs "
            "(W4-2R, W20-1) with A/B spacing only. Operator remains in PV. PV = 011 SHORT_DURATION. "
            "Fresh PDF re-fetched 2026-08-03 (prior capture had path-only text)."
        ),
    )
    s["applicability"].update({
        "closure": "Right lane closure",
        "duration": "Short Duration",
        "durationDefinition": "Work that occupies a location for up to 1 hour (Note 1).",
        "laneWidthFt": None,
        "laneWidthNote": "No lane/merging taper table on this sheet.",
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": ["URBAN", "RURAL"],
    })
    s["tableRoles"] = {
        "note": "Roles by content: 01=PV, 02=roll, 03=AW(A/B), 04=sizes. NO taperAndBuffer.",
        "protectiveVehicle": "203-01",
        "rollAheadDistance": "203-02",
        "advanceWarningSpacing": "203-03",
        "signSizes": "203-04",
    }
    aw = copy.deepcopy(T_AW_AB)
    sizes = {
        "title": "REQUIRED SIGN SIZES",
        "confidence": "verbatim",
        "keyedBy": ["signCode", "signSizeClass"],
        "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
        "rows": [
            size_row("NYW8-33", "48x24", "48x24"),
            size_row("W4-2R", "36x36", "48x48"),
            size_row("W20-1", "36x36", "48x48"),
            size_row("WARNING FLAG", "18x18", "18x18"),
        ],
    }
    s["tables"] = {
        "203-01": {**copy.deepcopy(T_PV_SD), "title": "PROTECTIVE VEHICLE REQUIREMENTS"},
        "203-02": clone_table("311-04", "203-02"),
        "203-03": aw,
        "203-04": sizes,
    }
    s["inputs"] = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
         "usedBy": ["203-01", "203-02", "203-03"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft",
         "usedBy": []},
        {"id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL"], "usedBy": ["203-03"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                     "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": ["203-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["203-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["203-04"]},
    ]
    s["corridor"] = {
        "confidence": "drawing",
        "description": "Short duration: W20-1 —B— W4-2R —A— ROLL AHEAD — work area. No taper/buffer/downstream/G20-2.",
        "zones": [
            {"id": "signB", "order": 1, "kind": "sign", "signCode": "W20-1",
             "sheetLegend": "ROAD WORK XX"},
            {"id": "gapB", "order": 2, "kind": "gap", "sheetLabel": "B",
             "sheetReference": "(SEE TABLE 203-03)",
             "lengthSource": {"table": "203-03", "column": "B",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signA", "order": 3, "kind": "sign", "signCode": "W4-2R",
             "sheetLegend": "(merge symbol)"},
            {"id": "gapA", "order": 4, "kind": "gap", "sheetLabel": "A",
             "sheetReference": "(SEE TABLE 203-03)",
             "lengthSource": {"table": "203-03", "column": "A",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "protectiveVehicle1", "order": 5, "kind": "symbol",
             "sheetLabel": "VEH #2", "lengthSource": None,
             "note": "Operator remains in the protective vehicle (Note 2)."},
            {"id": "rollAheadDistance", "order": 6, "kind": "clearance",
             "sheetLabel": "ROLL AHEAD DISTANCE", "sheetReference": "(SEE TABLE 203-02)",
             "lengthSource": {"table": "203-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True},
            {"id": "workArea", "order": 7, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
        ],
    }
    s["orderTable"] = {
        "confidence": "drawing",
        "description": "Upstream only — short duration has no downstream taper/G20-2.",
        "alignments": [{
            "alignIdx": 1, "name": "Upstream",
            "station0": "Upstream edge of the WORK AREA",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance",
                 "label": "ROLL AHEAD DISTANCE"},
                {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W4-2R",
                 "spacingZone": "gapA"},
                {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W20-1",
                 "spacingZone": "gapB"},
            ],
            "excludedRows": excluded_default() + [
                {"label": "BUFFER SPACE", "reason": "No buffer/taper table on short-duration sheet."},
                {"label": "LANE TAPER", "reason": "No taper table."},
                {"label": "MERGING TAPER", "reason": "No taper table."},
                {"label": "SHOULDER TAPER", "reason": "No taper table."},
                {"label": "DOWNSTREAM TAPER", "reason": "Not on short-duration plan."},
            ],
        }],
    }
    s["signs"] = {"confidence": "verbatim", "items": [
        sign_item("W20-1", sheetLegend="ROAD WORK XX",
                  legendSubstitution={"placeholder": "XX", "table": "203-03", "column": "XX"},
                  warningFlags=True, corridorZone="signB", positionRank=2,
                  sizeNonFreeway="36x36", sizeFreeway="48x48",
                  signLibraryBase="W20-01R"),
        sign_item("W4-2R", sheetLegend="(merge symbol)", warningFlags=True,
                  corridorZone="signA", positionRank=1,
                  sizeNonFreeway="36x36", sizeFreeway="48x48", signLibraryKey="W04-02R"),
        sign_item("NYW8-33", sheetLegend="LANE CLOSED", shape="rectangle",
                  postMounted=False, mountedOn="protectiveVehicle1",
                  sizeNonFreeway="48x24", sizeFreeway="48x24"),
        sign_item("WARNING FLAG", sheetLegend=None, shape="flag", postMounted=False,
                  mountedOn="W20-1, W4-2R", sizeNonFreeway="18x18", sizeFreeway="18x18"),
    ]}
    s["symbols"] = {"confidence": "drawing", "items": [
        {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True, "count": 1,
         "stationAnchor": {"zone": "gapA", "end": "downstream"},
         "lateralAnchor": "On the paved shoulder / closed lane"},
        {"id": "protectiveVehicle1", "sheetLabel": "VEH #2", "required": "per Table 203-01",
         "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"},
         "carriesSign": "NYW8-33",
         "note": "Operator remains in vehicle (short duration Note 2)."},
        {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES",
         "deviceSymbol": "CONE", "required": True,
         "longitudinalSpacing": {"maxFt": 40, "sheetText": "CONE SPACING NOT TO EXCEED 40 FT"}},
        {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
    ]}
    s["annotations"] = {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapB", "label": "B", "reference": "(SEE TABLE 203-03)"},
            {"zone": "gapA", "label": "A", "reference": "(SEE TABLE 203-03)"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE",
             "reference": "(SEE TABLE 203-02)"},
        ],
        "lateralDimensions": [],
        "leaderCallouts": [{"text": "LANE CLOSED", "pointsAt": "protectiveVehicle1"}],
        "notLabeled": [{"item": "Vehicle Space", "reason": "Not a zone on this sheet."}],
    }
    s["details"] = {}
    s["notes"] = {
        "confidence": "verbatim",
        "printed": [
            "1. SHORT DURATION IS WORK THAT OCCUPIES A LOCATION FOR UP TO 1 HOUR.",
            "2. THE OPERATOR(S) SHALL REMAIN IN THE PROTECTIVE VEHICLE(S) WITH THE SAFETY BELT AND HEADREST PROPERLY ADJUSTED, MAINTAIN VEHICLE SPACING, AND KEEP THE WHEELS ALIGNED WITH THE LANE STRIPING. TWO-WAY RADIOS SHOULD BE USED TO COMMUNICATE BETWEEN THE OPERATOR AND THE WORK CREW.",
            "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE ROLL AHEAD DISTANCE.",
            "4. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS. IF THERE IS A CONFLICT, MOVE THE SIGN UPSTREAM.",
        ],
        "planCallouts": ["CONE SPACING NOT TO EXCEED 40 FT"],
        "tableNotes": [],
    }
    s["rules"] = [
        {"id": "short-duration-operator-in-pv", "severity": "error", "source": "Note 2",
         "assert": "Operator remains in the protective vehicle during short-duration work.",
         "commonFailure": "Treating PV as unoccupied like short-term 311."},
        {"id": "no-taper-buffer", "severity": "error", "source": "Plan/tables",
         "assert": "No LANE/MERGING/SHOULDER TAPER or BUFFER SPACE sequential rows.",
         "commonFailure": "Copying 311 order table."},
        {"id": "sign-order-short", "severity": "error", "source": "Plan layout",
         "assert": "Advance signs upstream of roll ahead are W4-2R then W20-1 (A/B only).",
         "commonFailure": "Inserting W20-5R from 311."},
        {"id": "roll-ahead-is-a-range", "severity": "warning", "source": "Table 203-02",
         "assert": "Roll ahead is a MIN/MAX range.", "commonFailure": "Emitting a single value."},
    ]
    s["knownCodeDeviations"] = [
        {"where": "WZTCRules.GetDefaultUpstreamItems",
         "issue": "Default upstream rows include taper/buffer/Vehicle Space absent from this sheet.",
         "specSection": "orderTable"},
    ]
    s["knownExcerpts"] = {
        "from619-011": ["203-01 == 011-01 SHORT_DURATION NON-FREEWAY"],
        "from619-311": ["203-02 == 311-04 roll ahead", "203-03 A/B/XX/YY match 311-03 (no C)"],
        "differsFrom311": ["no taper table", "no G20-2/downstream", "operator in PV", "no W20-5R"],
    }
    write("203", s)


# =============================================================================
# 202 — Short Duration Left Lane Closure (mirror of 203)
# =============================================================================
def build_202():
    build_203()
    s = json.loads((SPEC_DIR / "619-203.json").read_text(encoding="utf-8"))
    s["sheet"].update({
        "number": "619-202",
        "title": "MULTILANE UNDIVIDED ROADWAY LEFT LANE CLOSURE",
        "sourceUrl": f"{SRC}/619-202.pdf",
        "localPdf": "Bridge/captures/619-202.pdf",
        "transcribedBy": "Cursor (Family 1 short-duration left sibling of 619-311)",
        "provenanceNote": (
            "Short duration left lane closure — mirror of 619-203 with W4-2L. "
            "NO taper/buffer. A/B spacing. Operator in PV. Fresh PDF 2026-08-03."
        ),
    })
    s["applicability"]["closure"] = "Left lane closure"
    # remap 203 -> 202
    raw = json.dumps(s)
    raw = raw.replace("203-", "202-").replace("W4-2R", "W4-2L").replace("W04-02R", "W04-02L")
    raw = raw.replace("RIGHT LANE", "LEFT LANE").replace("Right lane", "Left lane")
    s = json.loads(raw)
    s["tables"]["202-04"]["rows"] = [
        size_row("NYW8-33", "48x24", "48x24"),
        size_row("W4-2L", "36x36", "48x48"),
        size_row("W20-1", "36x36", "48x48"),
        size_row("WARNING FLAG", "18x18", "18x18"),
    ]
    s["notes"]["printed"].append(
        "5. WORK AREA, DEVICES, WORKERS OR EQUIPMENT SHALL NOT ENCROACH THE ONCOMING TRAVEL LANE. "
        "IF THE ADJACENT ONCOMING LANE IS REQUIRED TO PERFORM THE WORK, REFER TO STANDARD SHEET 619-325."
    )
    s["knownExcerpts"] = {
        "from619-203": ["mirror with W4-2L"],
        "differsFrom203": ["left closure", "Note 5 oncoming-lane / 619-325 referral"],
    }
    write("202", s)


# =============================================================================
# Helpers for 311-like corridor with MERGING TAPER label + optional extras
# =============================================================================
def apply_merging_label(s: dict, taper_table: str, aw_table: str, roll_table: str):
    for z in s["corridor"]["zones"]:
        if z["id"] == "laneTaper":
            z["sheetLabel"] = "MERGING TAPER"
            z["note"] = "Sheet labels this MERGING TAPER (L), not LANE TAPER."
        ls = z.get("lengthSource")
        if isinstance(ls, dict) and ls.get("table"):
            if "311-02" in ls["table"]:
                ls["table"] = taper_table
            if "311-03" in ls["table"]:
                ls["table"] = aw_table
            if "311-04" in ls["table"]:
                ls["table"] = roll_table
        if z.get("sheetReference"):
            z["sheetReference"] = (z["sheetReference"]
                                   .replace("311-02", taper_table)
                                   .replace("311-03", aw_table)
                                   .replace("311-04", roll_table))
    for al in s["orderTable"]["alignments"]:
        for r in al["rows"]:
            if r.get("label") == "LANE TAPER":
                r["label"] = "MERGING TAPER"
    for d in s["annotations"]["dimensions"]:
        if d.get("zone") == "laneTaper":
            d["label"] = "MERGING TAPER"
        if d.get("reference"):
            d["reference"] = (d["reference"]
                              .replace("311-02", taper_table)
                              .replace("311-03", aw_table)
                              .replace("311-04", roll_table))
    for item in s["signs"]["items"]:
        sub = item.get("legendSubstitution")
        if sub and sub.get("table", "").startswith("311-"):
            sub["table"] = sub["table"].replace("311-03", aw_table)
    for sym in s["symbols"]["items"]:
        for run in sym.get("runs", []):
            dcs = run.get("deviceCountSource")
            if dcs and dcs.get("table"):
                dcs["table"] = dcs["table"].replace("311-02", taper_table)


def build_317_like(num: str, *, duration: str, title: str, operation: str,
                   pv_table: dict, roles: dict, tables: dict, sizes: list,
                   extra_signs: list | None = None, device_spacing: int = 40,
                   notes: list[str], provenance: str, pdf_pages: int = 2,
                   page_rotation: int = 0, taper_label_merging: bool = True,
                   channelizing_title: str | None = None):
    s = base_from_311(
        number=f"619-{num}",
        title=title,
        operation=operation,
        sourceUrl=f"{SRC}/619-{num}.pdf",
        localPdf=f"Bridge/captures/619-{num}.pdf",
        pdfPages=pdf_pages,
        pageRotation=page_rotation,
        transcribedBy=f"Cursor (Family 1 sibling of 619-311)",
        provenanceNote=provenance,
    )
    s["applicability"]["duration"] = duration
    if duration == "Intermediate Term":
        s["applicability"]["durationDefinition"] = (
            "Stationary work occupying a location more than one daylight period up to "
            "3 consecutive days, or nighttime work lasting more than 1 hour."
        )
    elif duration == "Long Term":
        s["applicability"]["durationDefinition"] = (
            "Stationary work occupying a location more than 3 consecutive days."
        )
    s["tableRoles"] = roles
    s["tables"] = tables
    # Remap 311 refs in corridor etc to this sheet's roles
    taper = roles["taperAndBuffer"]
    aw = roles["advanceWarningSpacing"]
    roll = roles["rollAheadDistance"]
    pv = roles["protectiveVehicle"]
    sz = roles["signSizes"]
    remap_table_refs(s["corridor"], "311-02", taper)
    remap_table_refs(s["corridor"], "311-03", aw)
    remap_table_refs(s["corridor"], "311-04", roll)
    remap_table_refs(s["corridor"], "311-01", pv)
    remap_table_refs(s["annotations"], "311-02", taper)
    remap_table_refs(s["annotations"], "311-03", aw)
    remap_table_refs(s["annotations"], "311-04", roll)
    remap_table_refs(s["signs"], "311-03", aw)
    remap_table_refs(s["symbols"], "311-02", taper)
    remap_table_refs(s["symbols"], "311-01", pv)
    if taper_label_merging:
        apply_merging_label(s, taper, aw, roll)
    # sizes + signs sync
    s["tables"][sz] = {
        "title": "REQUIRED SIGN SIZES",
        "confidence": "verbatim",
        "keyedBy": ["signCode", "signSizeClass"],
        "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
        "rows": sizes,
    }
    size_codes = {r["signCode"] for r in sizes}
    # Rebuild signs from size table, keeping 311 metadata where codes match
    by_code = {si["signCode"]: si for si in s["signs"]["items"]}
    items = []
    for r in sizes:
        code = r["signCode"]
        if code in by_code:
            item = copy.deepcopy(by_code[code])
            item["sizeNonFreeway"] = r["NON-FREEWAY"]
            item["sizeFreeway"] = r["FREEWAY"]
            items.append(item)
        elif code == "NYR9-11":
            items.append(sign_item(
                "NYR9-11", sheetLegend="WORK ZONE SPEEDING ...", shape="rectangle",
                postMounted=True, corridorZone=None,
                sizeNonFreeway=r["NON-FREEWAY"], sizeFreeway=r["FREEWAY"],
                signLibraryKey="NYR9-11",
                note="Regulatory work-zone plaque; mid-advance placement per sheet notes.",
            ))
        elif code == "W20-5":
            # generic (no R/L) — treat as W20-5R family for right closures unless overridden
            items.append(sign_item(
                "W20-5", sheetLegend="LANE CLOSED YY",
                legendSubstitution={"placeholder": "YY", "table": aw, "column": "YY"},
                corridorZone="signB", positionRank=2,
                sizeNonFreeway=r["NON-FREEWAY"], sizeFreeway=r["FREEWAY"],
                signLibraryBase="W20-05R",
                note="Sheet prints W20-5 without R/L suffix; right-closure default W20-05R + legend.",
            ))
        else:
            items.append(sign_item(
                code, sizeNonFreeway=r["NON-FREEWAY"], sizeFreeway=r["FREEWAY"],
                signLibraryKey=None,
                note=f"Added from size table; confirm SignLibrary key before live placement.",
            ))
    if extra_signs:
        for es in extra_signs:
            if es["signCode"] not in {i["signCode"] for i in items}:
                items.append(es)
    # Drop signs not in size table
    s["signs"]["items"] = [i for i in items if i["signCode"] in size_codes]
    # Fix corridor signB if W20-5R replaced by W20-5
    codes = {i["signCode"] for i in s["signs"]["items"]}
    for z in s["corridor"]["zones"]:
        if z.get("signCode") == "W20-5R" and "W20-5R" not in codes and "W20-5" in codes:
            z["signCode"] = "W20-5"
        if z.get("signCode") and z["signCode"] not in codes and z["kind"] == "sign":
            # leave; structural check will catch — caller must fix
            pass
    for al in s["orderTable"]["alignments"]:
        for r in al["rows"]:
            if r.get("signCode") == "W20-5R" and "W20-5R" not in codes and "W20-5" in codes:
                r["signCode"] = "W20-5"
    # device spacing
    for sym in s["symbols"]["items"]:
        if sym.get("id") == "channelizingDevices":
            sym["longitudinalSpacing"] = {
                "maxFt": device_spacing,
                "sheetText": f"Channelizing spacing not to exceed {device_spacing} ft",
            }
    if channelizing_title and roles.get("channelizingApplication"):
        s["tables"][roles["channelizingApplication"]] = channelizing_stub(
            channelizing_title, f"max {device_spacing} ft")
    s["notes"] = {"confidence": "verbatim", "printed": notes, "planCallouts": [], "tableNotes": []}
    # inputs usedBy remapped
    for inp in s["inputs"]:
        inp["usedBy"] = [u.replace("311-", f"{num}-") for u in inp.get("usedBy", [])]
        # fix size table id if 05 vs 06
        inp["usedBy"] = [sz if u.endswith("-05") and "signSize" in inp["id"].lower()
                         or (inp["id"] == "signSizeClass" and u.endswith("-05"))
                         else u for u in inp["usedBy"]]
    # cleaner inputs rewrite
    s["inputs"] = []
    s["inputs"].append({
        "id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
        "usedBy": [pv, taper, aw, roll],
    })
    s["inputs"].append({
        "id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "usedBy": [taper],
    })
    s["inputs"].append({
        "id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "usedBy": [taper],
    })
    s["inputs"].append({
        "id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL"], "usedBy": [aw],
    })
    s["inputs"].append({
        "id": "exposureCondition", "type": "enum",
        "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                    "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": [pv],
    })
    s["inputs"].append({
        "id": "closureType", "type": "enum",
        "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
        "usedBy": [pv],
    })
    s["inputs"].append({
        "id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
        "default": "NON-FREEWAY", "usedBy": [sz],
    })
    s["details"] = {}
    s["knownCodeDeviations"] = ref311.get("knownCodeDeviations", [])[:3]
    return s


def build_317():
    roles = {
        "note": "Roles by CONTENT: 01=AW, 02=taper, 03=roll, 04=PV, 05=channelizing, 06=sizes.",
        "advanceWarningSpacing": "317-01",
        "taperAndBuffer": "317-02",
        "rollAheadDistance": "317-03",
        "protectiveVehicle": "317-04",
        "channelizingApplication": "317-05",
        "signSizes": "317-06",
    }
    tables = {
        "317-01": clone_table("311-03", "317-01"),
        "317-02": clone_table("311-02", "317-02"),
        "317-03": clone_table("311-04", "317-03"),
        "317-04": clone_table("311-01", "317-04"),
        "317-05": channelizing_stub(
            "CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES", "40 ft"),
        "317-06": {},  # filled by helper
    }
    sizes = [
        size_row("G20-2", "36x18", "48x24"),
        size_row("NYW8-33", "48x24", "48x24"),
        size_row("W4-2R", "36x36", "48x48"),
        size_row("W20-1", "36x36", "48x48"),
        size_row("W20-5", "36x36", "48x48"),
        size_row("WARNING FLAG", "18x18", "18x18"),
    ]
    notes = [
        "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS. IF THERE IS A CONFLICT, MOVE THE SIGN UPSTREAM.",
        "3. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
        "4. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "5. CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF EVERY 800' AS SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR GREATER IS CLOSED FOR A DISTANCE GREATER THAN 800'.",
    ]
    s = build_317_like(
        "317", duration="Short Term",
        title="MULTI LANE UNDIVIDED ROADWAY SINGLE LANE CLOSURE",
        operation="SHORT TERM OPERATION",
        pv_table=tables["317-04"], roles=roles, tables=tables, sizes=sizes,
        device_spacing=40, notes=notes,
        provenance=(
            "Short-term single lane closure sibling of 311. Tables 317-01..04 == 311-03/02/04/01 "
            "cell-for-cell. Adds 317-05 channelizing matrix. Plan labels MERGING TAPER. "
            "Size table prints W20-5 (no R/L). 2 pages."
        ),
        channelizing_title="CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES",
    )
    s["applicability"]["closure"] = "Single lane closure"
    s["applicability"]["roadway"] = "Multilane undivided"
    # Fix signB to W20-5
    for z in s["corridor"]["zones"]:
        if z["id"] == "signB":
            z["signCode"] = "W20-5"
            z["sheetLegend"] = "RIGHT LANE CLOSED YY"
    for al in s["orderTable"]["alignments"]:
        for r in al["rows"]:
            if r.get("zone") == "signB":
                r["signCode"] = "W20-5"
    s["knownExcerpts"] = {
        "from619-311": ["317-01==311-03", "317-02==311-02", "317-03==311-04", "317-04==311-01"],
        "differsFrom311": ["MERGING TAPER label", "channelizing matrix 317-05", "W20-5 not W20-5R", "2 pages"],
    }
    write("317", s)


def build_414():
    roles = {
        "note": "Roles by CONTENT: 01=AW, 02=taper, 03=roll, 04=PV, 05=channelizing, 06=sizes.",
        "advanceWarningSpacing": "414-01",
        "taperAndBuffer": "414-02",
        "rollAheadDistance": "414-03",
        "protectiveVehicle": "414-04",
        "channelizingApplication": "414-05",
        "signSizes": "414-06",
    }
    tables = {
        "414-01": clone_table("311-03", "414-01"),
        "414-02": clone_table("311-02", "414-02"),
        "414-03": clone_table("311-04", "414-03"),
        "414-04": copy.deepcopy(T_PV_INT),
        "414-05": channelizing_stub(
            "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
            "20 ft"),
        "414-06": {},
    }
    sizes = [
        size_row("G20-2", "36x18", "48x24"),
        size_row("NYR9-11", "24x42", "48x84"),
        size_row("NYW8-33", "48x24", "48x24"),
        size_row("W4-2R", "36x36", "48x48"),
        size_row("W20-1", "36x36", "48x48"),
        size_row("W20-5R", "36x36", "48x48"),
        size_row("WARNING FLAG", "18x18", "18x18"),
    ]
    notes = [
        "1. INTERMEDIATE-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION MORE THAN ONE DAYLIGHT PERIOD UP TO 3 CONSECUTIVE DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS. IF THERE IS A CONFLICT, MOVE THE SIGN UPSTREAM.",
        "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "4. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE ACTIVE WORK SPACE.",
        "5. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
        "6. THE NYR9-11 SIGN IS RECOMMENDED.",
    ]
    s = build_317_like(
        "414", duration="Intermediate Term",
        title="MULTI LANE UNDIVIDED ROADWAY SINGLE LANE CLOSURE",
        operation="INTERMEDIATE TERM OPERATION",
        pv_table=tables["414-04"], roles=roles, tables=tables, sizes=sizes,
        device_spacing=20, notes=notes,
        provenance=(
            "Intermediate single-lane sibling of 317/311. Taper/AW/roll == 311. "
            "PV = 011 INTERMEDIATE_TERM. Adds NYR9-11 + 20' device spacing + channelizing matrix. 2 pages."
        ),
        channelizing_title="CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
    )
    s["applicability"]["closure"] = "Single lane closure"
    s["knownExcerpts"] = {
        "from619-311": ["414-01==311-03", "414-02==311-02", "414-03==311-04"],
        "from619-011": ["414-04 == 011-01 INTERMEDIATE_TERM NON-FREEWAY"],
        "differsFrom311": ["20' spacing", "NYR9-11", "intermediate PV", "channelizing matrix"],
    }
    s["rules"].append({
        "id": "device-spacing-20ft", "severity": "error", "source": "Note 4",
        "assert": "Channelizing spacing <= 20 ft in active work space.",
        "commonFailure": "Copying 40 ft from short-term 311/317.",
    })
    write("414", s)


print("helpers loaded OK")
