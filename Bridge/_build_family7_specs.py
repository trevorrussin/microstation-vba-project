"""Build Family 7 mobile sheet specs: 619-111 (ref), 619-110, 619-112.

619-113 already authored under Family 5 — cross-ref only in STATUS.md.

Family 8 (101-104) has no downloadable PDFs — blocked separately.
"""
from __future__ import annotations

import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "Data" / "sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
DATE = "2026-08-03"

# Moving-operation GVW roll-ahead (110/112) — distinct from 301 stationary values.
ROLL_GVW_MOVING = [
    {
        "gvwBand": "9,500 TO 21,999 LBS.",
        "minGvwLbs": 9500,
        "maxGvwLbs": 21999,
        "min": {"ft": 200, "skipLines": 5},
        "max": {"ft": 240, "skipLines": 6},
    },
    {
        "gvwBand": "22,000 LBS. OR GREATER",
        "minGvwLbs": 22000,
        "maxGvwLbs": None,
        "min": {"ft": 160, "skipLines": 4},
        "max": {"ft": 200, "skipLines": 5},
    },
]

# Speed-keyed moving roll-ahead (111) — same numbers as 114-02.
ROLL_SPEED_MOVING = [
    {
        "speedBand": ">= 55 MPH",
        "minMph": 55,
        "maxMph": None,
        "min": {"ft": 200, "skipLines": 5},
        "max": {"ft": 280, "skipLines": 7},
    },
    {
        "speedBand": "45 - 50 MPH",
        "minMph": 45,
        "maxMph": 50,
        "min": {"ft": 160, "skipLines": 4},
        "max": {"ft": 240, "skipLines": 6},
    },
]

PV_P_TMIA_ROWS = [
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "P, TMIA",
    },
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
        "FREEWAY": "NA",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "P, TMIA",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
        "FREEWAY": "NA",
    },
]

PV_LEGEND_P = {
    "P": (
        "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED "
        "SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY "
        "CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY"
    ),
    "TMIA": "TMIA REQUIRED",
    "NA": "NOT APPLICABLE",
}

PV_PVH_ROWS = [
    {
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
    },
    {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
    },
]

PV_LEGEND_PVH = {
    "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
    "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
}

EXCLUDED_MOBILE = [
    {"label": "BUFFER SPACE", "reason": "No buffer on mobile sheet."},
    {"label": "SHOULDER TAPER", "reason": "No taper."},
    {"label": "MERGING TAPER", "reason": "No taper."},
    {"label": "Vehicle Space", "reason": "Not on this sheet."},
]


def write(n: str, spec: dict) -> None:
    path = OUT / f"{n}.json"
    path.write_text(json.dumps(spec, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {path.relative_to(ROOT)}")


def mobile_notes(fallback: str, sheet_of: str | None = None) -> list[str]:
    of = f" ({sheet_of})" if sheet_of else ""
    return [
        "1. MOBILE WORK IS WORK THAT MOVES INTERMITTENTLY OR CONTINUOUSLY WHERE WORK AT ANY SPECIFIC LOCATION COMPLETES WITHIN 15 MINUTES.",
        f"2. SHOULD THE WORK DURATION CONTINUE ON LONGER THAN THE 15 MINUTE MAXIMUM THE WORK ZONE TRAFFIC CONTROL SETUP SHALL BE RECONFIGURED AND ADJUSTED TO MEET THE REQUIREMENTS OF STANDARD SHEET {fallback}{of}.",
        "3. THIS TYPICAL MAY BE USED FOR VEHICLE BASED OPERATIONS SUCH AS SETTING UP STATIONARY TRAFFIC CONTROL (PLACING CONES, DRUMS AND SIGNS), BUT IS NOT TO BE USED FOR OPERATIONS THAT INVOLVE WORKERS ON FOOT PERFORMING ROADWAY AND / OR APPURTENANCE REPAIRS.",
        "4. VEHICLES SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
    ]


# ============================================================================= 111
spec111 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-111",
        "title": "WORK ZONE TRAFFIC CONTROL FREEWAY RIGHT LANE ENCROACHMENT/CLOSURE",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "MOBILE OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-111.pdf",
        "localPdf": "Bridge/captures/619-111.pdf",
        "localRender": None,
        "pdfPages": 2,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Family 7 mobile)",
        "transcribedOn": DATE,
        "provenanceNote": (
            "2-sheet freeway mobile right-lane. Sheet 1 (shoulder <8'): vehicle-mounted "
            "NYW8-33+W4-2R only, 750' max between VEH#1/#2, tables 111-01..03. "
            "Sheet 2 (shoulder >=8'): adds post-mounted W20-5R at 1500' min / 1/2 mile max, "
            "tables 111-04..06. Spec primary corridor/orderTable models Sheet 2 (registry signs). "
            "Roll-ahead is SPEED-keyed moving (like 114), NOT GVW. Fallback 619-206. "
            "No taper/buffer/AW tables."
        ),
    },
    "applicability": {
        "roadType": "Freeway",
        "roadway": "Freeway",
        "closure": "Right lane encroachment/closure",
        "duration": "Mobile",
        "durationDefinition": "Work that moves intermittently or continuously where work at any specific location completes within 15 minutes (Note 1).",
        "speedRangeMph": {
            "allowed": [45, 50, 55, 65],
            "note": "Roll-ahead bands >=55 and 45-50; 65 uses >=55.",
        },
        "laneWidthFt": None,
        "laneWidthNote": "No taper table — lane width not a lookup.",
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {
            "id": "preconstructionPostedSpeedMph",
            "type": "integer",
            "allowed": [45, 50, 55, 65],
            "usedBy": ["111-02", "111-05"],
        },
        {
            "id": "shoulderWidthBand",
            "type": "enum",
            "allowed": SH_BANDS,
            "default": ">= 8 ft",
            "usedBy": [],
            "note": "Sheet 1 plan is shoulder <8'; Sheet 2 is >=8'. Primary spec uses Sheet 2.",
        },
        {
            "id": "exposureCondition",
            "type": "enum",
            "allowed": [
                "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                "OTHER HAZARDS NO WORKERS EXPOSED",
            ],
            "usedBy": ["111-01", "111-04"],
        },
        {
            "id": "closureType",
            "type": "enum",
            "allowed": [
                "LANE CLOSURE OR ENCROACHMENT",
                "SHOULDER CLOSURE OR ENCROACHMENT",
            ],
            "default": "LANE CLOSURE OR ENCROACHMENT",
            "usedBy": ["111-01", "111-04"],
        },
        {
            "id": "signSizeClass",
            "type": "enum",
            "allowed": ["NON-FREEWAY", "FREEWAY"],
            "default": "FREEWAY",
            "usedBy": ["111-03", "111-06"],
        },
    ],
    "tableRoles": {
        "note": (
            "6 tables across 2 pages. Primary roles point at Sheet 2 (111-04/05/06). "
            "NO taperAndBuffer / advanceWarningSpacing."
        ),
        "protectiveVehicle": "111-04",
        "rollAheadDistance": "111-05",
        "signSizes": "111-06",
    },
    "tables": {
        "111-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Sheet 1. Identical cells to 111-04.",
            "rows": PV_P_TMIA_ROWS,
            "legend": PV_LEGEND_P,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT."
            ],
        },
        "111-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "Sheet 1. MOVING OPERATION (PV speed 15 MPH). Identical to 111-05.",
            "rows": ROLL_SPEED_MOVING,
            "usageNote": "MIN/MAX range.",
        },
        "111-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Sheet 1 — no W20-5R (narrow-shoulder / encroachment variant).",
            "rows": [
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
                {"signCode": "W4-2R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
            ],
        },
        "111-04": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Sheet 2 primary. Workers-exposed rows P, TMIA; OTHER HAZARDS rows NA.",
            "rows": PV_P_TMIA_ROWS,
            "legend": PV_LEGEND_P,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT."
            ],
        },
        "111-05": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "Sheet 2 primary. MOVING OPERATION values — same as 114-02.",
            "rows": ROLL_SPEED_MOVING,
            "usageNote": "MIN/MAX range.",
        },
        "111-06": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Sheet 2 primary — adds W20-5R.",
            "rows": [
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
                {"signCode": "W4-2R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W20-5R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": (
            "Sheet 2 primary: ROLL AHEAD (111-05) + W20-5R at 1500' min / 1/2 mile max. "
            "NYW8-33 and W4-2R vehicle-mounted. No tapers. Sheet 1 omits W20-5R."
        ),
        "zones": [
            {
                "id": "signA",
                "order": 1,
                "kind": "sign",
                "signCode": "W20-5R",
                "sheetLegend": "RIGHT LANE CLOSED AHEAD",
            },
            {
                "id": "gapA",
                "order": 2,
                "kind": "gap",
                "sheetLabel": "A",
                "lengthSource": {"fixedRange": {"minFt": 1500, "maxFt": 2640}},
                "dimensioned": True,
                "note": "Plan Sheet 2: 1500' (MIN.) / 1/2 MILE (MAX.).",
            },
            {
                "id": "protectiveVehicle",
                "order": 3,
                "kind": "symbol",
                "sheetLabel": "WORK VEHICLE / VEH #1..#4",
                "lengthSource": None,
            },
            {
                "id": "rollAheadDistance",
                "order": 4,
                "kind": "clearance",
                "sheetLabel": "ROLL AHEAD DISTANCE",
                "sheetReference": "(SEE TABLE 111-05)",
                "lengthSource": {
                    "table": "111-05",
                    "column": "range",
                    "lookupBy": ["preconstructionPostedSpeedMph"],
                },
                "dimensioned": True,
            },
            {
                "id": "workArea",
                "order": 5,
                "kind": "workArea",
                "sheetLabel": "WORK AREA",
                "lengthSource": None,
                "hatched": True,
                "dimensioned": False,
            },
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Sheet 2: Roll Ahead + W20-5R. Vehicle-mounted NYW8-33/W4-2R are not order-table sign rows.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {
                        "rowNum": 1,
                        "type": "Non-Sign",
                        "zone": "rollAheadDistance",
                        "label": "ROLL AHEAD DISTANCE",
                    },
                    {
                        "rowNum": 2,
                        "type": "Sign",
                        "zone": "signA",
                        "signCode": "W20-5R",
                        "spacingZone": "gapA",
                    },
                ],
                "excludedRows": EXCLUDED_MOBILE,
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {
                "signCode": "W20-5R",
                "sheetLegend": "RIGHT LANE CLOSED AHEAD",
                "shape": "diamond",
                "postMounted": True,
                "corridorZone": "signA",
                "sizeNonFreeway": "36x36",
                "sizeFreeway": "48x48",
                "signLibraryKey": "W20-05RA",
            },
            {
                "signCode": "NYW8-33",
                "sheetLegend": "LANE CLOSED",
                "shape": "rectangle",
                "postMounted": False,
                "mountedOn": "protectiveVehicle",
                "sizeNonFreeway": "48x24",
                "sizeFreeway": "48x24",
                "signLibraryKey": None,
            },
            {
                "signCode": "W4-2R",
                "sheetLegend": "RIGHT LANE ENDS (SYMBOL)",
                "shape": "diamond",
                "postMounted": False,
                "mountedOn": "protectiveVehicle",
                "sizeNonFreeway": "36x36",
                "sizeFreeway": "48x48",
                "signLibraryKey": "W04-02R",
            },
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {
                "id": "protectiveVehicle",
                "sheetLabel": "VEH #1..#4 / WORK VEHICLE",
                "required": True,
                "carriesSign": "NYW8-33",
            },
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapA", "label": "1500' (MIN.) … 1/2 MILE (MAX.)"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": mobile_notes("619-206", "SHEET 2 OF 2"),
    },
    "rules": [
        {
            "id": "fifteen-minute-cap",
            "severity": "error",
            "source": "Notes 1-2",
            "assert": "Mobile duration <=15 min; else use 619-206.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No merging/shoulder/downstream taper rows.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "sheet2-advance-w20-5r",
            "severity": "warning",
            "source": "Plan Sheet 2",
            "assert": "Advance W20-5R spacing is 1500' min / 1/2 mile max (Sheet 2 / shoulder >=8').",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
    ],
    "knownCodeDeviations": [],
}
write("619-111", spec111)


# ============================================================================= 110
spec110 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-110",
        "title": "WORK ZONE TRAFFIC CONTROL FREEWAY SHOULDER CLOSURE",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "MOBILE OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-110.pdf",
        "localPdf": "Bridge/captures/619-110.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 270,
        "transcribedBy": "Cursor (Family 7 mobile)",
        "transcribedOn": DATE,
        "provenanceNote": (
            "Freeway mobile shoulder closure (no lane encroachment, shoulder >=8'). "
            "3 tables. W8-23 only. PV matrix uses PVH+TMIA (like 301), NOT P,TMIA. "
            "Roll-ahead is GVW-keyed MOVING values (200/5-240/6 light; 160/4-200/5 heavy) — "
            "header '45-60 / w 55' is context like 301, not row keys. Fallback 619-205. "
            "Errata 1 Eff. 09/01/23 (EB 23-016)."
        ),
    },
    "applicability": {
        "roadType": "Freeway",
        "roadway": "Freeway",
        "closure": "Shoulder closure without lane encroachment",
        "duration": "Mobile",
        "durationDefinition": "Work that moves intermittently or continuously where work at any specific location completes within 15 minutes (Note 1).",
        "speedRangeMph": {
            "allowed": [45, 50, 55, 65],
            "note": "Header context 45-60 / >=55; roll-ahead keyed by GVW not speed.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {
            "id": "preconstructionPostedSpeedMph",
            "type": "integer",
            "allowed": [45, 50, 55, 65],
            "usedBy": [],
            "note": "Not used by roll-ahead (GVW-keyed).",
        },
        {
            "id": "protectiveVehicleGvwLbs",
            "type": "integer",
            "allowed": [9500, 22000],
            "default": 22000,
            "usedBy": ["110-02"],
            "note": "Roll-ahead keyed by PV GVW band.",
        },
        {
            "id": "shoulderWidthBand",
            "type": "enum",
            "allowed": SH_BANDS,
            "default": ">= 8 ft",
            "usedBy": [],
        },
        {
            "id": "exposureCondition",
            "type": "enum",
            "allowed": [
                "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
            ],
            "usedBy": ["110-01"],
        },
        {
            "id": "closureType",
            "type": "enum",
            "allowed": [
                "LANE CLOSURE OR ENCROACHMENT",
                "SHOULDER CLOSURE OR ENCROACHMENT",
            ],
            "default": "SHOULDER CLOSURE OR ENCROACHMENT",
            "usedBy": ["110-01"],
        },
        {
            "id": "signSizeClass",
            "type": "enum",
            "allowed": ["NON-FREEWAY", "FREEWAY"],
            "default": "FREEWAY",
            "usedBy": ["110-03"],
        },
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing. Roll-ahead GVW-keyed.",
        "protectiveVehicle": "110-01",
        "rollAheadDistance": "110-02",
        "signSizes": "110-03",
    },
    "tables": {
        "110-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "PVH+TMIA for both lane and shoulder workers-exposed rows. Single FREEWAY column.",
            "rows": PV_PVH_ROWS,
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
            ],
        },
        "110-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["protectiveVehicleGvwLbs"],
            "columnMeaning": (
                "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES — MOVING OPERATION MIN and MAX, "
                "by protective-vehicle gross vehicle weight (not posted speed)."
            ),
            "note": (
                "Header prints PRECONSTRUCTION POSTED SPEED LIMIT 45-60 / w 55 context "
                "but data rows are GVW bands (same structural pattern as 301-02, moving values)."
            ),
            "rows": ROLL_GVW_MOVING,
            "usageNote": "MIN/MAX range.",
        },
        "110-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "rows": [
                {"signCode": "W8-23", "NON-FREEWAY": None, "FREEWAY": "48x48"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": "Mobile shoulder: ROLL AHEAD (GVW table) + W8-23 on shoulder. PVH+TMIA + arrow panel caution. No tapers.",
        "zones": [
            {
                "id": "signA",
                "order": 1,
                "kind": "sign",
                "signCode": "W8-23",
                "sheetLegend": "NO SHOULDER",
            },
            {
                "id": "gapA",
                "order": 2,
                "kind": "gap",
                "sheetLabel": "A",
                "lengthSource": {"fixedFt": 0},
                "dimensioned": False,
                "note": "W8-23 sits at the protective vehicle — no numbered advance-gap callout.",
            },
            {
                "id": "protectiveVehicle",
                "order": 3,
                "kind": "symbol",
                "sheetLabel": "PVH WITH TMIA",
                "lengthSource": None,
            },
            {
                "id": "rollAheadDistance",
                "order": 4,
                "kind": "clearance",
                "sheetLabel": "ROLL AHEAD DISTANCE",
                "sheetReference": "(SEE TABLE 110-02)",
                "lengthSource": {
                    "table": "110-02",
                    "column": "range",
                    "lookupBy": ["protectiveVehicleGvwLbs"],
                },
                "dimensioned": True,
            },
            {
                "id": "workArea",
                "order": 5,
                "kind": "workArea",
                "sheetLabel": "WORK AREA",
                "lengthSource": None,
                "hatched": True,
                "dimensioned": False,
            },
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Roll Ahead + W8-23 only.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Protective vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {
                        "rowNum": 1,
                        "type": "Non-Sign",
                        "zone": "rollAheadDistance",
                        "label": "ROLL AHEAD DISTANCE",
                    },
                    {
                        "rowNum": 2,
                        "type": "Sign",
                        "zone": "signA",
                        "signCode": "W8-23",
                        "spacingZone": "gapA",
                    },
                ],
                "excludedRows": EXCLUDED_MOBILE,
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {
                "signCode": "W8-23",
                "sheetLegend": "NO SHOULDER",
                "shape": "diamond",
                "postMounted": True,
                "corridorZone": "signA",
                "sizeNonFreeway": "36x36",
                "sizeFreeway": "48x48",
                "signLibraryKey": "W08-23",
            },
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {
                "id": "protectiveVehicle",
                "sheetLabel": "PVH WITH TMIA",
                "required": True,
            },
            {
                "id": "workVehicle",
                "sheetLabel": "WORK VEHICLE",
                "required": True,
            },
            {
                "id": "arrowPanel",
                "sheetLabel": "ARROW PANEL (CAUTION MODE)",
                "required": True,
            },
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE (SEE TABLE 110-02)"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": mobile_notes("619-205"),
    },
    "rules": [
        {
            "id": "fifteen-minute-cap",
            "severity": "error",
            "source": "Notes 1-2",
            "assert": "Mobile duration <=15 min; else use 619-205.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No merging/shoulder/downstream taper rows.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "roll-ahead-by-gvw",
            "severity": "error",
            "source": "Table 110-02",
            "assert": "Roll ahead MIN/MAX comes from PV GVW band, not posted speed.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
    ],
    "knownCodeDeviations": [],
}
write("619-110", spec110)


# ============================================================================= 112
spec112 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-112",
        "title": "WORK ZONE TRAFFIC CONTROL FREEWAY RIGHT TWO LANE CLOSURE",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "MOBILE OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-112.pdf",
        "localPdf": "Bridge/captures/619-112.pdf",
        "localRender": None,
        "pdfPages": 2,
        "pageRotation": 270,
        "transcribedBy": "Cursor (Family 7 mobile)",
        "transcribedOn": DATE,
        "provenanceNote": (
            "2-sheet freeway mobile right TWO-lane closure. Sheet 1 (shoulder <8'): "
            "W20-5AR+NYW8-33, gaps 1000'/750'. Sheet 2 (shoulder >=8'): adds W4-2R, "
            "advance gap 1500' min / 1/2 mile max. Primary corridor models Sheet 2. "
            "PV = PVH+TMIA; roll-ahead GVW-keyed MOVING (same cells as 110-02). "
            "Fallback 619-207 / 619-209 per notes. Errata 1 Eff. 09/01/23."
        ),
    },
    "applicability": {
        "roadType": "Freeway",
        "roadway": "Freeway",
        "closure": "Right two lane closure",
        "duration": "Mobile",
        "durationDefinition": "Work that moves intermittently or continuously where work at any specific location completes within 15 minutes (Note 1).",
        "speedRangeMph": {
            "allowed": [45, 50, 55, 65],
            "note": "Header context 45-60 / >=55; roll-ahead keyed by GVW not speed.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {
            "id": "preconstructionPostedSpeedMph",
            "type": "integer",
            "allowed": [45, 50, 55, 65],
            "usedBy": [],
            "note": "Not used by roll-ahead (GVW-keyed).",
        },
        {
            "id": "protectiveVehicleGvwLbs",
            "type": "integer",
            "allowed": [9500, 22000],
            "default": 22000,
            "usedBy": ["112-02", "112-05"],
        },
        {
            "id": "shoulderWidthBand",
            "type": "enum",
            "allowed": SH_BANDS,
            "default": ">= 8 ft",
            "usedBy": [],
        },
        {
            "id": "exposureCondition",
            "type": "enum",
            "allowed": [
                "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
            ],
            "usedBy": ["112-01", "112-04"],
        },
        {
            "id": "closureType",
            "type": "enum",
            "allowed": [
                "LANE CLOSURE OR ENCROACHMENT",
                "SHOULDER CLOSURE OR ENCROACHMENT",
            ],
            "default": "LANE CLOSURE OR ENCROACHMENT",
            "usedBy": ["112-01", "112-04"],
        },
        {
            "id": "signSizeClass",
            "type": "enum",
            "allowed": ["NON-FREEWAY", "FREEWAY"],
            "default": "FREEWAY",
            "usedBy": ["112-03", "112-06"],
        },
    ],
    "tableRoles": {
        "note": (
            "6 tables across 2 pages. Primary roles = Sheet 2 (112-04/05/06). "
            "Roll-ahead GVW-keyed. NO taperAndBuffer / advanceWarningSpacing."
        ),
        "protectiveVehicle": "112-04",
        "rollAheadDistance": "112-05",
        "signSizes": "112-06",
    },
    "tables": {
        "112-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Sheet 1. Identical cells to 112-04.",
            "rows": PV_PVH_ROWS,
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
            ],
        },
        "112-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["protectiveVehicleGvwLbs"],
            "note": "Sheet 1. Identical moving-GVW cells to 110-02 / 112-05.",
            "rows": ROLL_GVW_MOVING,
            "usageNote": "MIN/MAX range.",
        },
        "112-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Sheet 1 — W20-5AR + NYW8-33 only (no W4-2R).",
            "rows": [
                {"signCode": "W20-5AR", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
            ],
        },
        "112-04": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Sheet 2 primary.",
            "rows": PV_PVH_ROWS,
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
            ],
        },
        "112-05": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["protectiveVehicleGvwLbs"],
            "note": "Sheet 2 primary. Plan callout SEE TABLE 112-05. Same cells as 112-02/110-02.",
            "rows": ROLL_GVW_MOVING,
            "usageNote": "MIN/MAX range.",
        },
        "112-06": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "Sheet 2 primary — adds W4-2R.",
            "rows": [
                {"signCode": "W20-5AR", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W4-2R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": (
            "Sheet 2: ROLL AHEAD (GVW) + W20-5AR at 1500' min / 1/2 mile max. "
            "NYW8-33/W4-2R vehicle-mounted. No tapers."
        ),
        "zones": [
            {
                "id": "signA",
                "order": 1,
                "kind": "sign",
                "signCode": "W20-5AR",
                "sheetLegend": "RIGHT TWO LANES CLOSED AHEAD",
            },
            {
                "id": "gapA",
                "order": 2,
                "kind": "gap",
                "sheetLabel": "A",
                "lengthSource": {"fixedRange": {"minFt": 1500, "maxFt": 2640}},
                "dimensioned": True,
                "note": "Plan Sheet 2: 1500' (MIN.) / 1/2 MILE (MAX.).",
            },
            {
                "id": "protectiveVehicle",
                "order": 3,
                "kind": "symbol",
                "sheetLabel": "PVH / WORK VEHICLES",
                "lengthSource": None,
            },
            {
                "id": "rollAheadDistance",
                "order": 4,
                "kind": "clearance",
                "sheetLabel": "ROLL AHEAD DISTANCE",
                "sheetReference": "(SEE TABLE 112-05)",
                "lengthSource": {
                    "table": "112-05",
                    "column": "range",
                    "lookupBy": ["protectiveVehicleGvwLbs"],
                },
                "dimensioned": True,
            },
            {
                "id": "workArea",
                "order": 5,
                "kind": "workArea",
                "sheetLabel": "WORK AREA",
                "lengthSource": None,
                "hatched": True,
                "dimensioned": False,
            },
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Sheet 2: Roll Ahead + W20-5AR. Vehicle-mounted signs excluded from order rows.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Protective vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {
                        "rowNum": 1,
                        "type": "Non-Sign",
                        "zone": "rollAheadDistance",
                        "label": "ROLL AHEAD DISTANCE",
                    },
                    {
                        "rowNum": 2,
                        "type": "Sign",
                        "zone": "signA",
                        "signCode": "W20-5AR",
                        "spacingZone": "gapA",
                    },
                ],
                "excludedRows": EXCLUDED_MOBILE,
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {
                "signCode": "W20-5AR",
                "sheetLegend": "RIGHT TWO LANES CLOSED AHEAD",
                "shape": "diamond",
                "postMounted": True,
                "corridorZone": "signA",
                "sizeNonFreeway": "36x36",
                "sizeFreeway": "48x48",
                "signLibraryKey": "W20-05aRA",
            },
            {
                "signCode": "W4-2R",
                "sheetLegend": "RIGHT LANE ENDS (SYMBOL)",
                "shape": "diamond",
                "postMounted": False,
                "mountedOn": "protectiveVehicle",
                "sizeNonFreeway": "36x36",
                "sizeFreeway": "48x48",
                "signLibraryKey": "W04-02R",
            },
            {
                "signCode": "NYW8-33",
                "sheetLegend": "LANE CLOSED",
                "shape": "rectangle",
                "postMounted": False,
                "mountedOn": "protectiveVehicle",
                "sizeNonFreeway": "48x24",
                "sizeFreeway": "48x24",
                "signLibraryKey": None,
            },
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH WITH TMIA", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapA", "label": "1500' (MIN.) … 1/2 MILE (MAX.)"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. MOBILE WORK IS WORK THAT MOVES INTERMITTENTLY OR CONTINUOUSLY WHERE WORK AT ANY SPECIFIC LOCATION COMPLETES WITHIN 15 MINUTES.",
            "2. SHOULD THE WORK DURATION CONTINUE ON LONGER THAN THE 15 MINUTE MAXIMUM THE WORK ZONE TRAFFIC CONTROL SETUP SHALL BE RECONFIGURED AND ADJUSTED TO MEET THE REQUIREMENTS OF STANDARD SHEET 619-207 / 619-209.",
            "3. THIS TYPICAL MAY BE USED FOR VEHICLE BASED OPERATIONS SUCH AS SETTING UP STATIONARY TRAFFIC CONTROL (PLACING CONES, DRUMS AND SIGNS), BUT IS NOT TO BE USED FOR OPERATIONS THAT INVOLVE WORKERS ON FOOT PERFORMING ROADWAY AND / OR APPURTENANCE REPAIRS.",
            "4. VEHICLES SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
            "5. IF AN ENTRANCE OR EXIT RAMP CONFLICTS WITH THE OPERATION, THE SETUP SHALL BE ADJUSTED.",
        ],
    },
    "rules": [
        {
            "id": "fifteen-minute-cap",
            "severity": "error",
            "source": "Notes 1-2",
            "assert": "Mobile duration <=15 min; else use 619-207/619-209.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No merging/shoulder/downstream taper rows.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
        {
            "id": "roll-ahead-by-gvw",
            "severity": "error",
            "source": "Table 112-05",
            "assert": "Roll ahead MIN/MAX comes from PV GVW band, not posted speed.",
            "commonFailure": "Ignoring this sheet-specific rule and using a generic default from another family.",
        },
    ],
    "knownCodeDeviations": [],
}
write("619-112", spec112)

print("Family 7 specs written.")
