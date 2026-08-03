"""Build Family 9 (mowing/mulching/marking) sheet specs.

Reference: 619-023. Pattern ≈ 619-041 (PV + roll + sizes, minimal order).
050/051 blocked separately (no PDF).
"""
from __future__ import annotations

import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "Data" / "sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"
DATE = "2026-08-03"
SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]

EXCLUDED_MOBILE = [
    {"label": "BUFFER SPACE", "reason": "No buffer on mowing/mulching sheet."},
    {"label": "SHOULDER TAPER", "reason": "No taper."},
    {"label": "MERGING TAPER", "reason": "No taper."},
    {"label": "Vehicle Space", "reason": "Not on this sheet."},
]

PV_LEGEND_PVH = {
    "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
    "PVL": "PROTECTIVE VEHICLE LIGHT (MINIMUM GROSS WEIGHT 9,500 LBS. OR GREATER) (SEE NOTE 3)",
    "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
}

PV_LEGEND_P = {
    "P": (
        "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED "
        "SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY "
        "CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY"
    ),
    "TMIA": "TMIA REQUIRED",
    "NA": "NOT APPLICABLE",
}

SPEED_BANDS_NF = [
    {"id": "ge45", "label": ">= 45 MPH", "minMph": 45, "maxMph": None},
    {"id": "b35to40", "label": "35 - 40 MPH", "minMph": 35, "maxMph": 40},
    {"id": "le30", "label": "<= 30 MPH", "minMph": None, "maxMph": 30},
]

# GVW x speed matrix for 022/023/032/060 — encoded as speed-keyed min/max
# (min=heavy GVW, max=light GVW) so sheet_spec.resolve works unchanged.
# Verbatim light/heavy cells kept for round-trip.
ROLL_GVW_SPEED_NF = [
    {
        "speedBand": "45 - 55 MPH",
        "minMph": 45,
        "maxMph": 55,
        "lightGvw": {"ft": 200, "skipLines": 5},
        "heavyGvw": {"ft": 160, "skipLines": 4},
        "min": {"ft": 160, "skipLines": 4},
        "max": {"ft": 200, "skipLines": 5},
    },
    {
        "speedBand": "<= 40 MPH",
        "minMph": None,
        "maxMph": 40,
        "lightGvw": {"ft": 120, "skipLines": 3},
        "heavyGvw": {"ft": 120, "skipLines": 3},
        "min": {"ft": 120, "skipLines": 3},
        "max": {"ft": 120, "skipLines": 3},
    },
]

ROLL_GVW_SPEED_FW = [
    {
        "speedBand": ">= 60 MPH",
        "minMph": 60,
        "maxMph": None,
        "lightGvw": {"ft": 240, "skipLines": 6},
        "heavyGvw": {"ft": 200, "skipLines": 5},
        "min": {"ft": 200, "skipLines": 5},
        "max": {"ft": 240, "skipLines": 6},
    },
    {
        "speedBand": "45 - 55 MPH",
        "minMph": 45,
        "maxMph": 55,
        "lightGvw": {"ft": 200, "skipLines": 5},
        "heavyGvw": {"ft": 160, "skipLines": 4},
        "min": {"ft": 160, "skipLines": 4},
        "max": {"ft": 200, "skipLines": 5},
    },
]

ROLL_SPEED_MOVING_3 = [
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
    {
        "speedBand": "<= 40 MPH",
        "minMph": None,
        "maxMph": 40,
        "min": {"ft": 120, "skipLines": 3},
        "max": {"ft": 200, "skipLines": 5},
    },
]


def write(n: str, spec: dict) -> None:
    path = OUT / f"{n}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {path.relative_to(ROOT)}")


def sheet_meta(num, title, operation, pdf_pages, rotation, note, approved="2021-12-02",
               ei="EI 21-028", local=None):
    return {
        "number": num,
        "title": title,
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": operation,
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": approved,
        "issuedUnder": ei,
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/{num}.pdf",
        "localPdf": local or f"Bridge/captures/{num}.pdf",
        "localRender": None,
        "pdfPages": pdf_pages,
        "pageRotation": rotation,
        "transcribedBy": "Cursor (Family 9 mowing/mulching/marking)",
        "transcribedOn": DATE,
        "provenanceNote": note,
    }


def pvh_speed_row(closure, ge45, b35, le30, exposure="WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"):
    return {
        "closureType": closure,
        "exposureCondition": exposure,
        "ge45": ge45,
        "b35to40": b35,
        "le30": le30,
    }


def mobile_corridor(sign_code, gap_note="500 FT. (MINIMUM) AND 2 MILE (MAXIMUM)"):
    return {
        "confidence": "drawing",
        "description": f"Moving operation: ROLL AHEAD + {sign_code}. No tapers.",
        "zones": [
            {
                "id": "signA",
                "order": 1,
                "kind": "sign",
                "signCode": sign_code,
            },
            {
                "id": "gapA",
                "order": 2,
                "kind": "gap",
                "sheetLabel": "A",
                "lengthSource": {"fixedFt": 500},
                "dimensioned": False,
                "note": gap_note,
            },
            {
                "id": "protectiveVehicle",
                "order": 3,
                "kind": "symbol",
                "sheetLabel": "PROTECTIVE VEHICLE / WORK VEHICLE",
                "lengthSource": None,
            },
            {
                "id": "rollAheadDistance",
                "order": 4,
                "kind": "clearance",
                "sheetLabel": "ROLL AHEAD DISTANCE",
                "lengthSource": {
                    "table": None,  # filled by caller via role
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
    }


def order_roll_sign(sign_code, roll_table):
    return {
        "confidence": "drawing",
        "description": f"Roll Ahead + {sign_code} only.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work / protective vehicle",
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
                        "signCode": sign_code,
                        "spacingZone": "gapA",
                    },
                ],
                "excludedRows": EXCLUDED_MOBILE,
            }
        ],
    }


def fix_roll_table(corridor, table_id):
    for z in corridor["zones"]:
        if z["id"] == "rollAheadDistance":
            z["lengthSource"]["table"] = table_id
            z["sheetReference"] = f"(SEE TABLE {table_id})"


# ============================================================================= 023 (family ref)
corr023 = mobile_corridor("W21-8")
fix_roll_table(corr023, "023-02")
# 023 also places NYW/W23 set on plan; primary advance is W21-8 / W23-1 cluster.
# Order-table primary walk uses W21-8 (mowing ahead) like registry primary; extra
# signs documented in signs.items for sizes.
spec023 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-023",
        "WORK ZONE TRAFFIC CONTROL LANE CLOSURE/ENCROACHMENT NON-FREEWAY - SHOULDER < 8' MOWING/MULCHING",
        "MOWING/MULCHING OPERATION",
        1,
        270,
        "Family 9 reference. Two-lane mowing/mulching lane encroachment, shoulder <8'. "
        "PVH/PVL×speed PV matrix; GVW×speed moving roll-ahead; W21-8/W23-1/NYW cluster. "
        "Fallback to 619-022 when vehicle stays on shoulder.",
        approved="2022-08-11",
        ei="EI 22-019",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Two-lane two-way, paved shoulder < 8 ft",
        "closure": "Lane closure / encroachment",
        "duration": "Mowing/Mulching (special)",
        "durationDefinition": "Special mowing/mulching operation; daylight shifts; suspend in poor visibility (Note 1).",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "PV bands <=30/35-40/>=45; roll-ahead <=40 / 45-55.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["023-01", "023-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": "<= 4 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"],
         "usedBy": ["023-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["023-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["023-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing. Roll is GVW×speed matrix encoded as speed-keyed min/max.",
        "protectiveVehicle": "023-01",
        "rollAheadDistance": "023-02",
        "signSizes": "023-03",
    },
    "tables": {
        "023-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "note": "NON-FREEWAY speed bands. LANE closure / workers-exposed only (no OTHER HAZARDS rows printed).",
            "speedBands": SPEED_BANDS_NF,
            "rows": [
                pvh_speed_row("LANE CLOSURE OR ENCROACHMENT", "PVH+TMIA", "PVL+TMIA", "PVL"),
            ],
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
                "3. THE USE OF A PROTECTIVE VEHICLE LIGHT (PVL) AS A SHADOW VEHICLE IS LIMITED TO NON-FREEWAY ROADWAYS WHERE THE POSTED SPEED LIMITS IS <= 40 MPH UNLESS OTHERWISE AUTHORIZED BY THE ENGINEER.",
            ],
        },
        "023-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": (
                "Printed as GVW columns (9,500-21,999 / 22,000+) × speed rows (45-55 / <=40), "
                "MOVING OPERATION 15 MPH MAX. Encoded speed-keyed with min=heavy / max=light for resolve(); "
                "lightGvw/heavyGvw preserve verbatim cells."
            ),
            "rows": ROLL_GVW_SPEED_NF,
            "usageNote": "MIN/MAX = heavy/light GVW for the speed band.",
        },
        "023-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W21-8", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W23-1", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW8-32", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "NYW8-35", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW23-1", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": corr023,
    "orderTable": order_roll_sign("W21-8", "023-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W21-8", "sheetLegend": "MOWING AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W21-08"},
            {"signCode": "W23-1", "sheetLegend": "SLOW TRAFFIC AHEAD", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01"},
            {"signCode": "NYW8-32", "sheetLegend": "SLOW MOVING VEHICLE AHEAD", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": None, "note": "Not in SignLibrary.bas — cell TBD."},
            {"signCode": "NYW8-35", "sheetLegend": "DO NOT PASS", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01xNY"},
            {"signCode": "NYW23-1", "sheetLegend": "DO NOT PASS", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01xNY"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH OR PVL", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "mower", "sheetLabel": "MOWER / MULCHING VEHICLE", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THIS MOWING/MULCHING OPERATION IS A SPECIAL OPERATION AND IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS AND HAVE LITTLE OR NO INTERFERENCE WITH TRAFFIC. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. THIS SHEET SHALL BE USED ON ROADWAYS WHERE THE VEHICLE ENCROACHES THE TRAVEL LANE. IF THE WORK VEHICLE REMAINS ENTIRELY ON THE SHOULDER OR PROVIDES A 10' MINIMUM TRAVEL LANE FOR THE DURATION OF THE OPERATION STANDARD SHEET 619-022 MAY BE USED.",
            "3. MOWERS/MULCHING VEHICLES SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES. IF IT IS NECESSARY FOR THE MOWER/MULCHING VEHICLES TO ENCROACH ONTO THE TRAVEL LANE, IT SHALL BE FOLLOWED BY A PROTECTIVE VEHICLE WITH OPERATING FLASHING LIGHTS.",
            "4. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON MOWERS OR WORK VEHICLE NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING TRACTOR WITHIN RIGHT OF WAY.",
            "5. REGARDLESS OF THE EXISTANCE OF A PASSING OR NO PASSING ZONE, THE WORK AND PROTECTIVE VEHICLES SHOULD PULL OVER PERIODICALLY WHERE POSSIBLE TO PROVIDE A 10' MINIMUM LANE WIDTH FOR VEHICULAR TRAFFIC TO PASS.",
            "6. THE WORK VEHICLE AND VEH #2 SHALL OPERATE FROM THE SHOULDER WHEREVER POSSIBLE. WHEN IT IS NECESSARY FOR THE WORK VEHICLE TO ENCROACH THE TRAVEL LANE, VEH #2 SHALL REMAIN IN THE TRAVEL LANE UNTIL THE WORK VEHICLE CAN CLEAR THE TRAVEL LANE.",
            "7. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL REMAIN WITHIN THE ALLOWABLE ROLL-AHEAD DISTANCE LIMITS.",
            "8. VERBAL COMMUNICATION SHALL BE ESTABLISHED AND MAINTAINED BETWEEN THE WORK VEHICLE AND PROTECTIVE VEHICLE(S) FOR SPACING AND CONTROL OF TRAFFIC QUEUES.",
        ],
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        },
        {
            "id": "fallback-022",
            "severity": "warning",
            "source": "Note 2",
            "assert": "If work stays on shoulder / 10' travel lane remains, use 619-022.",
            "commonFailure": "Using 023 when 022 applies.",
        },
    ],
    "knownCodeDeviations": [],
}
write("619-023", spec023)


# ============================================================================= 022
corr022 = mobile_corridor("W21-8")
# Also W8-23 on PV — order primary = W21-8; W8-23 is on vehicle.
fix_roll_table(corr022, "022-02")
spec022 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-022",
        "WORK ZONE TRAFFIC CONTROL SHOULDER CLOSURE/LANE ENCROACHMENT NON-FREEWAY MOWING",
        "MOWING OPERATION",
        1,
        270,
        "Non-freeway mowing shoulder closure / lane encroachment. PVH/PVL matrix for both "
        "LANE and SHOULDER closures; GVW×speed roll-ahead; W21-8 + W8-23.",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Non-freeway",
        "closure": "Shoulder closure / lane encroachment",
        "duration": "Mowing (special)",
        "durationDefinition": "Special mowing operation; daylight; suspend in poor visibility (Note 1).",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "PV <=30/35-40/>=45; roll <=40 / 45-55.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["022-01", "022-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"],
         "usedBy": ["022-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": ["022-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["022-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing.",
        "protectiveVehicle": "022-01",
        "rollAheadDistance": "022-02",
        "signSizes": "022-03",
    },
    "tables": {
        "022-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "note": "Both LANE and SHOULDER workers-exposed rows; same PVH+TMIA/PVL+TMIA/PVL pattern. No OTHER HAZARDS value rows.",
            "speedBands": SPEED_BANDS_NF,
            "rows": [
                pvh_speed_row("LANE CLOSURE OR ENCROACHMENT", "PVH+TMIA", "PVL+TMIA", "PVL"),
                pvh_speed_row("SHOULDER CLOSURE OR ENCROACHMENT", "PVH+TMIA", "PVL+TMIA", "PVL"),
            ],
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
                "3. THE USE OF A PROTECTIVE VEHICLE LIGHT (PVL) AS A SHADOW VEHICLE IS LIMITED TO NON-FREEWAY ROADWAYS WHERE THE POSTED SPEED LIMITS IS <= 40 MPH UNLESS OTHERWISE AUTHORIZED BY THE ENGINEER.",
            ],
        },
        "022-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "Same GVW×speed matrix as 023-02.",
            "rows": ROLL_GVW_SPEED_NF,
            "usageNote": "MIN/MAX = heavy/light GVW for the speed band.",
        },
        "022-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W21-8", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W8-23", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": corr022,
    "orderTable": order_roll_sign("W21-8", "022-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W21-8", "sheetLegend": "MOWING AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W21-08"},
            {"signCode": "W8-23", "sheetLegend": "NO SHOULDER", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": "W08-23"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH OR PVL", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "mower", "sheetLabel": "MOWER", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [{"label": "10' MIN.", "note": "Minimum open width beside work."}],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THE MOWING OPERATION IS A SPECIAL OPERATION. IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. TRACTOR MOWERS SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES.",
            "3. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON TRACTORS NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING TRACTOR WITHIN RIGHT OF WAY.",
            "4. IF SHOULDER AREA BECOMES TOO NARROW FOR THE PROTECTIVE VEHICLE(S) TO BE COMPLETELY ON THE SHOULDER, THE VEHICLES SHALL STAY ON THE WIDER SHOULDER AREA UNTIL OPERATORS CAN SAFELY DRIVE AROUND THE NARROW SHOULDER TO NEW SET-UP POINT.",
            "5. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
        ],
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        }
    ],
    "knownCodeDeviations": [],
}
write("619-022", spec022)


# ============================================================================= 021 — work beyond shoulder (minimal)
spec021 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-021",
        "WORK ZONE TRAFFIC CONTROL WORK BEYOND SHOULDER NON-FREEWAY MOWING",
        "MOWING OPERATION",
        1,
        270,
        "Work beyond shoulder mowing. AW spacing table + W21-8 sizes. No PV/roll. "
        "Plan places W21-8 at 500' min / 2 mile max. live-build n/a (sign-only, no roll row).",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Non-freeway, work beyond shoulder",
        "closure": "Work beyond shoulder",
        "duration": "Mowing (special)",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "AW URBAN bands + RURAL.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": ["URBAN", "RURAL"],
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["021-01"]},
        {"id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL"], "default": "RURAL",
         "usedBy": ["021-01"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["021-02"]},
    ],
    "tableRoles": {
        "note": "2 tables. NO PV/roll/taper. Plan gap is fixed 500'-2mi (not AW A/B/C).",
        "advanceWarningSpacing": "021-01",
        "signSizes": "021-02",
    },
    "tables": {
        "021-01": {
            "title": "ADVANCE WARNING SIGN SPACING",
            "confidence": "verbatim",
            "keyedBy": ["roadType"],
            "note": "Same URBAN/RURAL shape as 011-06 (no FREEWAY row on this non-freeway sheet).",
            "rows": [
                {"roadType": "URBAN", "speedBand": "<= 30 MPH", "minMph": None, "maxMph": 30,
                 "A": 100, "B": 100, "C": 100, "XX": "AHEAD", "YY": "AHEAD"},
                {"roadType": "URBAN", "speedBand": "35-40 MPH", "minMph": 35, "maxMph": 40,
                 "A": 200, "B": 200, "C": 200, "XX": "AHEAD", "YY": "AHEAD"},
                {"roadType": "URBAN", "speedBand": ">= 45 MPH", "minMph": 45, "maxMph": None,
                 "A": 350, "B": 350, "C": 350, "XX": "AHEAD", "YY": "AHEAD"},
                {"roadType": "RURAL", "speedBand": "ALL", "minMph": None, "maxMph": None,
                 "A": 500, "B": 500, "C": 500, "XX": "1500 FT.", "YY": "1000 FT."},
            ],
        },
        "021-02": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W21-8", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": "Work beyond shoulder: W21-8 only at 500' min / 2 mi max. No PV/roll/tapers.",
        "zones": [
            {"id": "signA", "order": 1, "kind": "sign", "signCode": "W21-8"},
            {"id": "gapA", "order": 2, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 500}, "dimensioned": False,
             "note": "500 FT. (MINIMUM) AND 2 MILE (MAXIMUM)"},
            {"id": "workArea", "order": 3, "kind": "workArea", "sheetLabel": "MOWER / WORK AREA",
             "lengthSource": None, "hatched": False, "dimensioned": False},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "W21-8 only — no roll/taper. live-build n/a (sign-only payload).",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Mower",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Sign", "zone": "signA", "signCode": "W21-8",
                     "spacingZone": "gapA"},
                ],
                "excludedRows": EXCLUDED_MOBILE + [
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No protective vehicle / roll-ahead on this sheet."},
                ],
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W21-8", "sheetLegend": "MOWING AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W21-08"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "mower", "sheetLabel": "MOWER", "required": True},
        ],
    },
    "annotations": {"confidence": "drawing", "dimensions": [], "lateralDimensions": []},
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THE MOWING OPERATION IS A SPECIAL OPERATION. IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. TRACTOR MOWERS SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES. (2' MIN. FROM FACE OF CURB OR 6' MIN. FROM EDGE OF SHOULDER)",
            "3. \"MOWING AHEAD\" SIGN IS NECESSARY ON BOTH SIDES SIMULTANEOUSLY IF THE WORK WILL OCCUR ON BOTH SIDES OF THE ROAD.",
            "4. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON TRACTORS NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING TRACTOR WITHIN RIGHT OF WAY.",
        ],
    },
    "rules": [
        {
            "id": "no-roll",
            "severity": "error",
            "source": "Plan",
            "assert": "No ROLL AHEAD / taper rows.",
            "commonFailure": "Copying mobile roll-ahead order from 022/023.",
        }
    ],
    "knownCodeDeviations": [],
}
write("619-021", spec021)


# ============================================================================= 031
corr031 = mobile_corridor("W20-1")
# Also W8-23 on plan — primary advance W20-1
fix_roll_table(corr031, "031-02")
spec031 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-031",
        "WORK ZONE TRAFFIC CONTROL SHOULDER CLOSURE/LANE ENCROACHMENT TWO-LANE TWO-WAY MULCHING/HERBICIDE",
        "MULCHING/HERBICIDE OPERATION",
        1,
        0,
        "Two-lane mulching/herbicide. Classic P,TMIA NON-FREEWAY PV (like 041); speed-keyed "
        "moving roll-ahead 3-band; W20-1 + W8-23. Registry title 'Freeway Mowing' is wrong — PDF is two-lane mulching.",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Two-lane two-way",
        "closure": "Shoulder closure / lane encroachment",
        "duration": "Mulching/Herbicide (special)",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "PV <=30/35-40/>=45; roll <=40/45-50/>=55.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["031-01", "031-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS NO WORKERS EXPOSED",
         ], "usedBy": ["031-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": ["031-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["031-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing. PV is P,TMIA (not PVH/PVL).",
        "protectiveVehicle": "031-01",
        "rollAheadDistance": "031-02",
        "signSizes": "031-03",
    },
    "tables": {
        "031-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "note": "NON-FREEWAY only. Same shape as 041-01.",
            "speedBands": SPEED_BANDS_NF,
            "rows": [
                {"closureType": "LANE CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                 "ge45": "P, TMIA", "b35to40": "P, TMIA", "le30": "P"},
                {"closureType": "LANE CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
                 "ge45": "NA", "b35to40": "NA", "le30": "NA"},
                {"closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                 "ge45": "P, TMIA", "b35to40": "P", "le30": "P"},
                {"closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
                 "ge45": "NA", "b35to40": "NA", "le30": "NA"},
            ],
            "legend": PV_LEGEND_P,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT",
            ],
        },
        "031-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "MOVING OPERATION (based on protective vehicle speed of 15 MPH). Classic 3-band min/max.",
            "rows": ROLL_SPEED_MOVING_3,
            "usageNote": "MIN/MAX range.",
        },
        "031-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W8-23", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": corr031,
    "orderTable": order_roll_sign("W20-1", "031-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "sheetLegend": "ROAD WORK AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RA"},
            {"signCode": "W8-23", "sheetLegend": "NO SHOULDER", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": "W08-23"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "24000 LB PROTECTIVE VEHICLE", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "workVehicle", "sheetLabel": "WORK VEHICLE (MULCHING/HERBICIDE)", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [{"label": "10' (MIN.)", "note": "Minimum open width."}],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THE HERBICIDE/MULCHING OPERATION IS A SPECIAL OPERATION. IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. WORK VEHICLE SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES. IF IT IS NECESSARY FOR THE MOWER TO ENCROACH ONTO THE TRAVEL LANE, THE WORK VEHICLE SHALL BE FOLLOWED BY A PROTECTIVE VEHICLE WITH OPERATING FLASHING LIGHTS.",
            "3. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON WORK VEHICLES NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING WORK VEHICLE WITHIN RIGHT OF WAY. PESTICIDE APPLICATORS ARE REQUIRED TO WEAR APPROVED PPE THAT ADHERES TO THE PRODUCT LABEL.",
            "4. IF SHOULDER AREA BECOMES TOO NARROW FOR VEH #1 TO BE COMPLETELY ON THE SHOULDER, THE VEHICLE SHALL STAY ON THE WIDER SHOULDER AREA UNTIL OPERATOR CAN SAFELY DRIVE AROUND THE NARROW SHOULDER TO NEW SET-UP POINT. VEH #1 SHALL STAY AS FAR TO THE RIGHT AS PRACTICAL.",
            "5. WHERE PRACTICAL AND AS NEEDED, THE WORK AND PROTECTIVE VEHICLES SHOULD PULL OVER PERIODICALLY TO ALLOW VEHICULAR TRAFFIC TO PASS.",
            "6. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION. AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
        ],
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        }
    ],
    "knownCodeDeviations": [],
}
write("619-031", spec031)


# ============================================================================= 032
corr032 = mobile_corridor("W23-1")
fix_roll_table(corr032, "032-02")
spec032 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-032",
        "WORK ZONE TRAFFIC CONTROL LANE CLOSURE/ENCROACHMENT NON-FREEWAY - SHOULDER < 8' HERBICIDE",
        "HERBICIDE OPERATION",
        1,
        270,
        "Herbicide lane encroachment, shoulder <8'. Sibling of 023 with herbicide notes; "
        "same PVH/PVL + GVW×speed roll shape. Fallback to 619-031 when work stays on shoulder.",
        approved="2022-08-11",
        ei="EI 22-019",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Non-freeway, paved shoulder < 8 ft",
        "closure": "Lane closure / encroachment",
        "duration": "Herbicide (special)",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "PV <=30/35-40/>=45; roll <=40 / 45-55.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["032-01", "032-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": "<= 4 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"],
         "usedBy": ["032-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["032-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["032-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing.",
        "protectiveVehicle": "032-01",
        "rollAheadDistance": "032-02",
        "signSizes": "032-03",
    },
    "tables": {
        "032-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "speedBands": SPEED_BANDS_NF,
            "rows": [
                pvh_speed_row("LANE CLOSURE OR ENCROACHMENT", "PVH+TMIA", "PVL+TMIA", "PVL"),
            ],
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
                "3. THE USE OF A PROTECTIVE VEHICLE LIGHT (PVL) AS A SHADOW VEHICLE IS LIMITED TO NON-FREEWAY ROADWAYS WHERE THE POSTED SPEED LIMITS IS <= 40 MPH UNLESS OTHERWISE AUTHORIZED BY THE ENGINEER.",
            ],
        },
        "032-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "Same GVW×speed matrix as 023-02.",
            "rows": ROLL_GVW_SPEED_NF,
            "usageNote": "MIN/MAX = heavy/light GVW for the speed band.",
        },
        "032-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W23-1", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW8-32", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "NYW8-35", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW23-1", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": corr032,
    "orderTable": order_roll_sign("W23-1", "032-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W23-1", "sheetLegend": "SLOW TRAFFIC AHEAD", "shape": "rectangle",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": "W23-01"},
            {"signCode": "NYW8-32", "sheetLegend": "SLOW MOVING VEHICLE AHEAD", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": None},
            {"signCode": "NYW8-35", "sheetLegend": "DO NOT PASS", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01xNY"},
            {"signCode": "NYW23-1", "sheetLegend": "DO NOT PASS", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01xNY"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH OR PVL", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "workVehicle", "sheetLabel": "WORK VEHICLE (HERBICIDE)", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THE HERBICIDE OPERATION IS A SPECIAL OPERATION. IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS AND HAVE LITTLE OR NO INTERFERENCE WITH TRAFFIC. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. THIS SHEET SHALL BE USED ON ROADWAYS WHERE THE WORK VEHICLE ENCROACHES THE TRAVEL LANE. IF THE WORK REMAINS ENTIRELY ON THE SHOULDER OR PROVIDES A 10' MINIMUM TRAVEL LANE FOR THE DURATION OF THE OPERATION STANDARD SHEET 619-031 MAY BE USED.",
            "3. WORK VEHICLE SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES. IF IT IS NECESSARY FOR THE WORK VEHICLE TO ENCROACH ONTO THE TRAVEL LANE, THE WORK VEHICLE SHALL BE FOLLOWED BY A PROTECTIVE VEHICLE WITH OPERATING FLASHING LIGHTS.",
            "4. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON WORK VEHICLES NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING WORK VEHICLE WITHIN RIGHT OF WAY. PESTICIDE APPLICATORS ARE REQUIRED TO WEAR APPROVED PPE THAT ADHERES TO THE PRODUCT LABEL.",
            "5. REGARDLESS OF THE EXISTANCE OF A PASSING OR NO PASSING ZONE, THE WORK AND PROTECTIVE VEHICLES SHOULD PULL OVER PERIODICALLY WHERE POSSIBLE TO PROVIDE A 10' MINIMUM LANE WIDTH FOR VEHICULAR TRAFFIC TO PASS.",
            "6. THE WORK VEHICLE AND VEH #2 SHALL OPERATE FROM THE SHOULDER WHEREVER POSSIBLE. WHEN IT IS NECESSARY FOR THE WORK VEHICLE TO ENCROACH THE TRAVEL LANE, VEH #2 SHALL REMAIN IN THE TRAVEL LANE UNTIL THE WORK VEHICLE CAN CLEAR THE TRAVEL LANE.",
            "7. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL REMAIN WITHIN THE ALLOWABLE ROLL-AHEAD DISTANCE LIMITS.",
            "8. VERBAL COMMUNICATION SHALL BE ESTABLISHED AND MAINTAINED BETWEEN THE WORK VEHICLE AND PROTECTIVE VEHICLE(S) FOR SPACING AND CONTROL OF TRAFFIC QUEUES.",
        ],
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        },
        {
            "id": "fallback-031",
            "severity": "warning",
            "source": "Note 2",
            "assert": "If work stays on shoulder / 10' travel lane remains, use 619-031.",
            "commonFailure": "Using 032 when 031 applies.",
        },
    ],
    "knownCodeDeviations": [],
}
write("619-032", spec032)


# ============================================================================= 033
corr033 = mobile_corridor("W20-1")
fix_roll_table(corr033, "033-02")
spec033 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-033",
        "WORK ZONE TRAFFIC CONTROL SHOULDER CLOSURE FREEWAY MULCHING/HERBICIDE",
        "MULCHING/HERBICIDE OPERATION",
        1,
        270,
        "Freeway shoulder closure mulching/herbicide. FREEWAY-only PVH+TMIA; GVW×speed "
        "roll-ahead with >=60 and 45-55 bands; W20-1 + W8-23. Plan typo SEE TABLE 034-02 → 033-02.",
    ),
    "applicability": {
        "roadType": "Freeway",
        "roadway": "Freeway",
        "closure": "Shoulder closure",
        "duration": "Mulching/Herbicide (special)",
        "speedRangeMph": {
            "allowed": [45, 50, 55, 60, 65],
            "note": "Roll-ahead bands 45-55 / >=60.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [45, 50, 55, 60, 65], "usedBy": ["033-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"],
         "usedBy": ["033-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": ["033-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["FREEWAY"],
         "default": "FREEWAY", "usedBy": ["033-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. FREEWAY-only PV. NO taperAndBuffer / advanceWarningSpacing.",
        "protectiveVehicle": "033-01",
        "rollAheadDistance": "033-02",
        "signSizes": "033-03",
    },
    "tables": {
        "033-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Single FREEWAY column. PVH+TMIA for both closures (workers-exposed). No OTHER HAZARDS rows.",
            "rows": [
                {"closureType": "LANE CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
                 "FREEWAY": "PVH+TMIA"},
                {"closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
                 "exposureCondition": "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
                 "FREEWAY": "PVH+TMIA"},
            ],
            "legend": {
                "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
                "TMIA": "TRUCK/TRAILER MOUNTED IMPACT ATTENUATOR",
            },
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
            ],
        },
        "033-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "GVW×speed matrix with >=60 and 45-55 bands. Encoded speed-keyed min=heavy/max=light.",
            "rows": ROLL_GVW_SPEED_FW,
            "usageNote": "MIN/MAX = heavy/light GVW for the speed band.",
        },
        "033-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "FREEWAY column only on this sheet.",
            "rows": [
                {"signCode": "W20-1", "FREEWAY": "48x48"},
                {"signCode": "W8-23", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": corr033,
    "orderTable": order_roll_sign("W20-1", "033-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "sheetLegend": "ROAD WORK AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeFreeway": "48x48", "signLibraryKey": "W20-01RA"},
            {"signCode": "W8-23", "sheetLegend": "NO SHOULDER", "shape": "diamond",
             "postMounted": True, "sizeFreeway": "48x48", "signLibraryKey": "W08-23"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeFreeway": "18x18", "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "workVehicle", "sheetLabel": "WORK VEHICLE", "required": True},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THE HERBICIDE/MULCHING OPERATIONS IS A SPECIAL OPERATION. IT SHALL BE SCHEDULED AND COMPLETED DURING DAYLIGHT WORK SHIFTS. THE WORK SHALL BE SUSPENDED DURING PERIODS OF POOR VISIBILITY.",
            "2. WORK VEHICLE SHALL HAVE AN AMBER BEACON OPERATING AT ALL TIMES. IF IT IS NECESSARY FOR THE VEHICLE TO ENCROACH ONTO THE TRAVEL LANE, THE WORK VEHICLE SHALL BE FOLLOWED BY A PROTECTIVE VEHICLE WITH OPERATING FLASHING LIGHTS.",
            "3. APPROVED PERSONAL PROTECTIVE EQUIPMENT (PPE) SHALL BE WORN WHILE ON WORK VEHICLES NOT EQUIPPED WITH AN ENCLOSED CAB. PPE IS REQUIRED WHEN EXITING WORK VEHICLE WITHIN RIGHT OF WAY. PESTICIDE APPLICATIONS ARE REQUIRED TO WEAR APPROVED PPE THAT ADHERES TO THE PRODUCT LABEL.",
            "4. IF SHOULDER AREA BECOMES TOO NARROW FOR VEH #1 TO BE COMPLETELY ON THE SHOULDER, THE VEHICLE SHALL STAY ON THE WIDER SHOULDER AREA UNTIL OPERATOR CAN SAFELY DRIVE AROUND THE NARROW SHOULDER TO NEW SET-UP POINT.",
            "5. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION. AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
        ],
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        }
    ],
    "knownCodeDeviations": [
        {
            "id": "plan-table-typo-034",
            "note": "Plan callout prints 'SEE TABLE 034-02' but sheet tables are 033-*; transcribed as 033-02.",
        }
    ],
}
write("619-033", spec033)


# ============================================================================= 060
corr060 = mobile_corridor("W23-1")
fix_roll_table(corr060, "060-02")
spec060 = {
    "schemaVersion": "1.1",
    "sheet": sheet_meta(
        "619-060",
        "WORK ZONE TRAFFIC CONTROL PAVEMENT MARKING OPERATIONS NON-FREEWAY",
        "PAVEMENT MARKING OPERATIONS",
        2,
        270,
        "2-sheet pavement marking train. Sheet 1 = plan + notes; Sheet 2 = tables 060-01..04. "
        "PVH/PVL + GVW×speed roll like 023; signs W23-1/NYW8-*/W3-4. Primary order = ROLL + W23-1.",
        approved="2022-08-11",
        ei="EI 22-019",
    ),
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Non-freeway (two-lane / multi-lane pavement marking)",
        "closure": "Lane closure (pavement marking train)",
        "duration": "Pavement Marking (special)",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55],
            "note": "PV <=30/35-40/>=45; roll <=40 / 45-55.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55], "usedBy": ["060-01", "060-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC"],
         "usedBy": ["060-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["060-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["060-03"]},
    ],
    "tableRoles": {
        "note": "4 tables on sheet 2. 060-04 is PVMS messaging (not a spacing role).",
        "protectiveVehicle": "060-01",
        "rollAheadDistance": "060-02",
        "signSizes": "060-03",
        "pvmsMessaging": "060-04",
    },
    "tables": {
        "060-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "speedBands": SPEED_BANDS_NF,
            "rows": [
                pvh_speed_row("LANE CLOSURE OR ENCROACHMENT", "PVH+TMIA", "PVL+TMIA", "PVL"),
            ],
            "legend": PV_LEGEND_PVH,
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUME THERE IS NO POSITIVE PROTECTION PRESENT.",
                "2. TRUCK/TRAILER MOUNTED IMPACT ATTENUATORS (TMIA) SHALL NOT BE MOUNTED/INSTALLED ON VEHICLES WITH A GROSS VEHICLE WEIGHT (GVW) LESS THAN WHAT IS MINIMALLY REQUIRED BY THE MANUFACTURER OF THE TMIA.",
                "3. THE USE OF A PROTECTIVE VEHICLE LIGHT (PVL) AS A SHADOW VEHICLE IS LIMITED TO NON-FREEWAY ROADWAYS WHERE THE POSTED SPEED LIMITS IS <= 40 MPH UNLESS OTHERWISE AUTHORIZED BY THE ENGINEER.",
            ],
        },
        "060-02": {
            "title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "Same GVW×speed matrix as 023-02.",
            "rows": ROLL_GVW_SPEED_NF,
            "usageNote": "MIN/MAX = heavy/light GVW for the speed band.",
        },
        "060-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W23-1", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "W3-4", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "NYW23-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "NYW8-32", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW8-31", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "NYW8-30", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
        "060-04": {
            "title": "PAVEMENT MARKING PVMS MESSAGING",
            "confidence": "drawing",
            "keyedBy": ["scenario"],
            "note": "Multi-phase PVMS messages for wet paint / caution conditions. Partial transcription of scenario headers; full phase text on PDF sheet 2.",
            "rows": [
                {"scenario": "CAUTION/HAZARD", "condition": "YELLOW LINES WET",
                 "note": "See PDF sheet 2 for PHASE 1/2/3 message cells."},
                {"scenario": "CAUTION/HAZARD", "condition": "WHITE LINES WET",
                 "note": "See PDF sheet 2 for PHASE 1/2/3 message cells."},
                {"scenario": "CAUTION/HAZARD", "condition": "YELLOW & WHITE LINES WET",
                 "note": "See PDF sheet 2 for PHASE 1/2/3 message cells."},
            ],
        },
    },
    "corridor": corr060,
    "orderTable": order_roll_sign("W23-1", "060-02"),
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W23-1", "sheetLegend": "SLOW TRAFFIC AHEAD", "shape": "rectangle",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": "W23-01"},
            {"signCode": "W3-4", "sheetLegend": "BE PREPARED TO STOP", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": "W03-04",
             "note": "Conditional per Special Note S1 (queue relief)."},
            {"signCode": "NYW23-1", "sheetLegend": "SLOW TRAFFIC AHEAD (static)", "shape": "diamond",
             "postMounted": True, "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": "W23-01"},
            {"signCode": "NYW8-32", "sheetLegend": "DO NOT PASS / STAY IN LANE family", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": None},
            {"signCode": "NYW8-31", "sheetLegend": "STAY IN LANE (front mount)", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W23-01wNY"},
            {"signCode": "NYW8-30", "sheetLegend": "WET PAINT", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
             "signLibraryKey": "W21-02NY"},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "workVehicle", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
             "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "PVH OR PVL (VEH #1/#2/#3)", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL (CAUTION MODE)", "required": True},
            {"id": "pavementMarkingVehicle", "sheetLabel": "PAVEMENT MARKING VEHICLE", "required": True},
            {"id": "pvms", "sheetLabel": "PORTABLE VARIABLE MESSAGE SIGN", "required": False},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [{"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"}],
        "lateralDimensions": [],
    },
    "notes": {
        "confidence": "drawing",
        "printed": [
            "1. TRAFFIC QUEUES SHALL BE CONTINUOUSLY MONITORED. TRAFFIC SHALL BE RELIEVED AS SOON AS PRACTICABLE BY PULLING OFF ON A SHOULDER THAT SUPPORTS THE WIDTH AND WEIGHT OF THE VEHICLES IN THE PAVEMENT MARKING TRAIN.",
            "2. VEH #1 SHALL NOT ENCROACH INTO TRAVEL LANE, STAY AS FAR TO THE RIGHT AS POSSIBLE AND ADJUST ITS SPACING TO ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION.",
            "3. VERBAL COMMUNICATION MUST BE MAINTAINED BETWEEN ALL VEHICLES IN TRAIN.",
            "6. PORTABLE VARIABLE MESSAGE SIGNS (PVMS) MAY BE USED TO SUPPLEMENT REQUIRED WARNING SIGNS USING APPROVED MESSAGES (SEE TABLE 060-04).",
            "7. VEH #3 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL REMAIN WITHIN THE ALLOWABLE ROLL-AHEAD DISTANCE LIMITS.",
        ],
        "note": "Full Notes 1-10 + Special Notes S1-S9 + Use of Cones C1-C3 on sheet 1 — see PDF for complete verbatim set.",
    },
    "rules": [
        {
            "id": "no-tapers",
            "severity": "error",
            "source": "Plan",
            "assert": "No taper/buffer rows in default order.",
            "commonFailure": "Emitting default MERGING/SHOULDER TAPER from another family.",
        }
    ],
    "knownCodeDeviations": [],
}
write("619-060", spec060)

print("Family 9 corridor specs written.")
