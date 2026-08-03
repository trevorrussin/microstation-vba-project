"""Build complete Family 4 sheet specs from _draft_*_tables.json + corridor geometry."""
from __future__ import annotations

import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
SPEC_DIR = ROOT / "Data/sheet-specs"

d306 = json.loads((SPEC_DIR / "_draft_619306_tables.json").read_text(encoding="utf-8"))
d212 = json.loads((SPEC_DIR / "_draft_619212_tables.json").read_text(encoding="utf-8"))
d114 = json.loads((SPEC_DIR / "_draft_619114_tables.json").read_text(encoding="utf-8"))
d041 = json.loads((SPEC_DIR / "_draft_619041_tables.json").read_text(encoding="utf-8"))

SH_BANDS = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"


def write(name: str, spec: dict) -> None:
    path = SPEC_DIR / f"{name}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("Wrote", path)


# ============================================================================= 306
spec306 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-306",
        "title": "WORK ZONE TRAFFIC CONTROL RIGHT LANE CLOSURE PARKWAY - SHOULDER < 8 FOOT",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "SHORT TERM OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-306.pdf",
        "localPdf": "Bridge/captures/619-306.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Family 4 parkway reference)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": (
            "Family 4 reference — parkway right lane closure with shoulder < 8 ft. "
            "Hybrid of Family 2 corridor (MERGING + DOWNSTREAM) and Family 3 spacing "
            "(fixed plan gaps 1000/1500/2640, no advance-warning table). "
            "NO shoulder-taper dimension on plan despite L/3 columns in table 306-03. "
            "NO NYW8-33. Speeds 45/50/55/65 only. Tables verified vs 302-02 overlap identical."
        ),
    },
    "applicability": {
        "roadType": "Parkway",
        "roadway": "Parkway with paved shoulder less than 8 feet",
        "closure": "Right lane closure",
        "duration": "Short Term",
        "durationDefinition": "Daytime work that occupies a location for more than 1 hour within a single daylight period (Note 1).",
        "speedRangeMph": {
            "allowed": [45, 50, 55, 65],
            "note": "Table 306-03 covers 45/50/55/65 only — same 4-speed set as Family 3 301-03.",
        },
        "laneWidthFt": [10, 11, 12],
        "shoulderWidthBands": SH_BANDS,
        "shoulderWidthBandNote": "Sheet title restricts applicability to shoulder < 8 ft; table still prints all three L/3 bands.",
        "areaTypes": None,
        "areaTypeNote": "No advance-warning spacing table; gaps are fixed plan callouts 1000'/1500'/2640'.",
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": [45, 50, 55, 65],
         "usedBy": ["306-01", "306-02", "306-03"]},
        {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "usedBy": ["306-03"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": "<= 4 ft",
         "usedBy": ["306-03"], "note": "Parkway sheet is for shoulder < 8; default narrow band."},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS NO WORKERS EXPOSED",
         ], "usedBy": ["306-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["306-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": ["306-04"]},
    ],
    "tableRoles": {
        "note": (
            "Family 4 parkway reference. Roles by CONTENT: 306-01=PV, 306-02=rollAhead, "
            "306-03=taperAndBuffer, 306-04=signSizes. NO advanceWarningSpacing role."
        ),
        "protectiveVehicle": "306-01",
        "rollAheadDistance": "306-02",
        "taperAndBuffer": "306-03",
        "signSizes": "306-04",
    },
    "tables": {
        "306-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "roadTypeForProtectiveVehicle"],
            "note": "FREEWAY column only in PDF text layer (parkway uses freeway PV column). All 4 cells P, TMIA.",
            "rows": d306["tables"]["306-01"]["rows"],
            "legend": d306["tables"]["306-01"]["legend"],
            "tableNotes": [
                "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT"
            ],
        },
        "306-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES — STATIONARY OPERATION MIN and MAX",
            "note": "2 bands (>=55, 45-50). Identical to 302-05 first two rows. No <=40 row.",
            "rows": d306["tables"]["306-02"]["rows"],
            "usageNote": "MIN/MAX range, not a single value.",
        },
        "306-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "laneWidthFt", "shoulderWidthBand"],
            "columnMeaning": {
                "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
                "laneTaper": "TAPER LENGTH: L (FT.)/ # OF SKIP LINES/ # OF CHANNELIZING DEVICES, FOR LANE WIDTH IN FT.",
                "shoulderTaper": "SHOULDER TAPER LENGTH: L/3 FOR SHOULDER WIDTH — printed in table; NOT dimensioned on this plan",
            },
            "note": "4 speeds 45/50/55/65. All cells identical to 302-02 on overlapping speeds. Plan uses MERGING TAPER only (no shoulder-taper dimension).",
            "rows": d306["tables"]["306-03"]["rows"],
        },
        "306-04": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "note": "5 rows — NO NYW8-33 (unlike 302). FREEWAY sizes only extracted from text layer.",
            "rows": [
                {"signCode": "G20-2", "NON-FREEWAY": None, "FREEWAY": "48x24"},
                {"signCode": "W4-2R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W20-5R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": (
            "From extract_plan_geometry.py: DOWNSTREAM 50-100 → ROLL AHEAD → BUFFER → "
            "MERGING TAPER → 1000' → 1500' → 2640'. No shoulder-taper overlay (absent from plan)."
        ),
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK 1 MILE"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
             "lengthSource": {"fixedFt": 2640}, "dimensioned": True, "spans": "W20-1 to W20-5R",
             "note": "Plan callout 2640' — no AW spacing table."},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-5R", "sheetLegend": "RIGHT LANE CLOSED ½ MILE"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"fixedFt": 1500}, "dimensioned": True, "spans": "W20-5R to W4-2R"},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W4-2R", "sheetLegend": "(merge symbol)"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 1000}, "dimensioned": True,
             "spans": "W4-2R to upstream end of MERGING TAPER",
             "note": "Unlike 302, shoulder taper is NOT an overlay inside gap A on this sheet."},
            {"id": "laneTaper", "order": 7, "kind": "taper", "sheetLabel": "MERGING TAPER",
             "sheetReference": "(SEE TABLE 306-03)",
             "lengthSource": {"table": "306-03", "column": "laneTaper",
                              "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
             "dimensioned": True},
            {"id": "bufferSpace", "order": 8, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "sheetReference": "(SEE TABLE 306-03)",
             "lengthSource": {"table": "306-03", "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True, "mustBeEmptyReason": "Note 3."},
            {"id": "protectiveVehicle", "order": 9, "kind": "symbol", "sheetLabel": "VEH #1",
             "lengthSource": None},
            {"id": "rollAheadDistance", "order": 10, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": "(SEE TABLE 306-02)",
             "lengthSource": {"table": "306-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True, "mustBeEmptyReason": "Note 3."},
            {"id": "workArea", "order": 11, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
            {"id": "downstreamTaper", "order": 12, "kind": "taper", "sheetLabel": "DOWNSTREAM TAPER",
             "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}},
             "sheetText": "50'-100'", "dimensioned": True},
            {"id": "gapEndRoadWork", "order": 13, "kind": "gap", "sheetLabel": None,
             "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False},
            {"id": "signEndRoadWork", "order": 14, "kind": "sign", "signCode": "G20-2",
             "sheetLegend": "END ROAD WORK"},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Upstream: Roll Ahead, Buffer, Merging Taper, W4-2R, W20-5R, W20-1. Downstream: Downstream Taper + G20-2. No shoulder-taper row.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Upstream edge of the WORK AREA",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                    {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                    {"rowNum": 3, "type": "Non-Sign", "zone": "laneTaper", "label": "MERGING TAPER"},
                    {"rowNum": 4, "type": "Sign", "zone": "signA", "signCode": "W4-2R", "spacingZone": "gapA"},
                    {"rowNum": 5, "type": "Sign", "zone": "signB", "signCode": "W20-5R", "spacingZone": "gapB"},
                    {"rowNum": 6, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
                ],
                "excludedRows": [
                    {"label": "SHOULDER TAPER", "reason": "Not dimensioned on parkway shoulder<8 plan (table cols only)."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                    {"label": "Upstream Taper Temp Barrier", "reason": "No temporary barrier."},
                    {"label": "Upstream Taper Box/Corr Beam", "reason": "No box/corr beam."},
                ],
            },
            {
                "alignIdx": 2,
                "name": "Downstream",
                "station0": "Downstream edge of the WORK AREA",
                "walkDirection": "Downstream, with traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper", "label": "DOWNSTREAM TAPER"},
                    {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                     "spacingZone": "gapEndRoadWork"},
                ],
            },
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "sheetLegend": "ROAD WORK 1 MILE", "shape": "diamond",
             "warningFlags": True, "postMounted": True, "corridorZone": "signC",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RM"},
            {"signCode": "W20-5R", "sheetLegend": "RIGHT LANE CLOSED ½ MILE", "shape": "diamond",
             "postMounted": True, "corridorZone": "signB",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-05RM"},
            {"signCode": "W4-2R", "sheetLegend": "(merge symbol)", "shape": "diamond",
             "warningFlags": True, "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W04-02R"},
            {"signCode": "G20-2", "sheetLegend": "END ROAD WORK", "shape": "rectangle",
             "postMounted": True, "corridorZone": "signEndRoadWork",
             "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02"},
            {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
             "mountedOn": "W20-1, W4-2R", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True,
             "stationAnchor": {"zone": "laneTaper", "end": "upstream"}},
            {"id": "protectiveVehicle", "sheetLabel": "VEH #1", "required": True,
             "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"}},
            {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
             "longitudinalSpacing": {"maxFt": 40},
             "runs": [
                 {"id": "laneTaperRun", "zone": "laneTaper"},
                 {"id": "longitudinalRun", "zone": "bufferSpace..workArea"},
                 {"id": "downstreamRun", "zone": "downstreamTaper"},
             ]},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapC", "label": "2640'"},
            {"zone": "gapB", "label": "1500'"},
            {"zone": "gapA", "label": "1000'"},
            {"zone": "laneTaper", "label": "MERGING TAPER"},
            {"zone": "bufferSpace", "label": "BUFFER SPACE"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
            {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": d306["notes"]["printed"],
        "notesOrderNote": "Exactly 3 numbered notes — fewer than 302's 8 (no left-lane / VEH#2 / transverse / 40' notes on this parkway sheet).",
    },
    "rules": [
        {"id": "no-occupancy-buffer-rollahead", "severity": "error", "source": "Note 3",
         "assert": "No workers/equipment/vehicles in buffer or roll ahead."},
        {"id": "sign-order", "severity": "error", "source": "Plan",
         "assert": "Upstream signs W4-2R then W20-5R then W20-1."},
        {"id": "no-shoulder-taper-row", "severity": "error", "source": "Plan geometry",
         "assert": "Do not emit SHOULDER TAPER as a sequential station — not dimensioned on this sheet."},
        {"id": "fixed-freeway-gaps", "severity": "error", "source": "Plan callouts",
         "assert": "Gaps are fixed 1000/1500/2640 — do not look up an AW spacing table."},
    ],
    "knownCodeDeviations": [],
}
write("619-306", spec306)


# ============================================================================= 212
spec212 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-212",
        "title": "WORK ZONE TRAFFIC CONTROL RIGHT/LEFT LANE CLOSURE PARKWAY",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "SHORT DURATION OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-212.pdf",
        "localPdf": "Bridge/captures/619-212.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Family 4 short-duration)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": (
            "Short-duration parkway lane closure. Plan: SHOULDER TAPER only (no MERGING/DOWNSTREAM). "
            "Table 212-03 still has full lane+shoulder columns (identical to 302-02 overlap). "
            "Fixed gaps 500'/1500'/2640'(printed as 1/2 MILE). NYW8-33 on PV. Operator stays in vehicle. "
            "Roll-ahead speeds 45-65 only — no <=40 row despite earlier recon guess."
        ),
    },
    "applicability": {
        "roadType": "Parkway",
        "roadway": "Parkway",
        "closure": "Right/left lane closure",
        "duration": "Short Duration",
        "durationDefinition": "Work that occupies a location for up to 1 hour (Note 1).",
        "speedRangeMph": {"allowed": [45, 50, 55, 65],
                          "note": "Taper/roll tables cover 45-65; no 40 mph row in PDF."},
        "laneWidthFt": [10, 11, 12],
        "laneWidthNote": "Lane taper columns exist in table 212-03 but plan uses SHOULDER TAPER only.",
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
        "areaTypeNote": "No AW spacing table; fixed plan gaps 500/1500/2640.",
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": [45, 50, 55, 65],
         "usedBy": ["212-02", "212-03"]},
        {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "default": 12,
         "usedBy": ["212-03"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft",
         "usedBy": ["212-03"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS NO WORKERS EXPOSED",
         ], "usedBy": ["212-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["212-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": ["212-04"]},
    ],
    "tableRoles": {
        "note": "212-01=PV, 212-02=rollAhead, 212-03=taperAndBuffer (shoulder used on plan), 212-04=signSizes. No AW role.",
        "protectiveVehicle": "212-01",
        "rollAheadDistance": "212-02",
        "taperAndBuffer": "212-03",
        "signSizes": "212-04",
    },
    "tables": {
        "212-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "FREEWAY column only; all P, TMIA.",
            "rows": d212["tables"]["212-01"]["rows"],
            "legend": {
                "P": "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY",
                "TMIA": "TMIA REQUIRED",
            },
        },
        "212-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "rows": d212["tables"]["212-02"]["rows"],
            "usageNote": "MIN/MAX range. STATIONARY OPERATION.",
        },
        "212-03": {
            "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph", "laneWidthFt", "shoulderWidthBand"],
            "note": "Full lane+shoulder columns (match 302-02). Plan dimensions SHOULDER TAPER only; BUFFER not dimensioned on plan.",
            "rows": d212["tables"]["212-03"]["rows"],
        },
        "212-04": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "rows": [
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
                {"signCode": "W4-2R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "W20-5R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": (
            "extract_plan_geometry: ROLL AHEAD → PV → SHOULDER TAPER → 500' → 1500' → 1/2 MILE(=2640'). "
            "No MERGING/DOWNSTREAM/BUFFER dimensions."
        ),
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
             "lengthSource": {"fixedFt": 2640}, "dimensioned": True,
             "note": "Plan prints 1/2 MILE (=2640 ft) as furthest gap."},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-5R", "sheetLegend": "RIGHT LANE CLOSED"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"fixedFt": 1500}, "dimensioned": True},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W4-2R", "sheetLegend": "(merge symbol)"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 500}, "dimensioned": True,
             "note": "500' (1 REFERENCE MARKER) — shorter than freeway A=1000."},
            {"id": "shoulderTaper", "order": 7, "kind": "taper", "sheetLabel": "SHOULDER TAPER",
             "sheetReference": "(SEE TABLE 212-03)",
             "lengthSource": {"table": "212-03", "column": "shoulderTaper",
                              "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
             "dimensioned": True, "consumesStation": True,
             "note": "Plan label TAPER LENGTH → shoulder L/3; no merging taper."},
            {"id": "protectiveVehicle", "order": 8, "kind": "symbol", "sheetLabel": "WORK VEHICLE",
             "lengthSource": None, "note": "Operator remains in vehicle; carries NYW8-33."},
            {"id": "rollAheadDistance", "order": 9, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": "(SEE TABLE 212-02)",
             "lengthSource": {"table": "212-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 10, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Roll Ahead, Shoulder Taper, W4-2R, W20-5R, W20-1. No buffer/merging/downstream rows.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work vehicle / work area",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                    {"rowNum": 2, "type": "Non-Sign", "zone": "shoulderTaper", "label": "SHOULDER TAPER"},
                    {"rowNum": 3, "type": "Sign", "zone": "signA", "signCode": "W4-2R", "spacingZone": "gapA"},
                    {"rowNum": 4, "type": "Sign", "zone": "signB", "signCode": "W20-5R", "spacingZone": "gapB"},
                    {"rowNum": 5, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
                ],
                "excludedRows": [
                    {"label": "BUFFER SPACE", "reason": "Buffer column in table 212-03 but not dimensioned on plan."},
                    {"label": "MERGING TAPER", "reason": "Short-duration parkway plan shows shoulder taper only."},
                    {"label": "DOWNSTREAM TAPER", "reason": "Not on this sheet."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "shape": "diamond", "postMounted": True, "corridorZone": "signC",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RM"},
            {"signCode": "W20-5R", "shape": "diamond", "postMounted": True, "corridorZone": "signB",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-05RM"},
            {"signCode": "W4-2R", "shape": "diamond", "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W04-02R"},
            {"signCode": "NYW8-33", "sheetLegend": "LANE CLOSED", "shape": "rectangle",
             "postMounted": False, "mountedOn": "protectiveVehicle",
             "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": None,
             "signLibraryNote": "Vehicle-mounted; not emitted as order-table Sign row."},
            {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
             "mountedOn": "W20-1", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "WORK VEHICLE", "required": True,
             "carriesSign": "NYW8-33",
             "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"}},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": False},
            {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
             "longitudinalSpacing": {"maxFt": 40},
             "runs": [{"id": "shoulderTaperRun", "zone": "shoulderTaper"}]},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapC", "label": "1/2 MILE"},
            {"zone": "gapB", "label": "1500'"},
            {"zone": "gapA", "label": "500'"},
            {"zone": "shoulderTaper", "label": "TAPER LENGTH"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. SHORT DURATION IS WORK THAT OCCUPIES A LOCATION FOR UP TO 1 HOUR.",
            "2. THE OPERATOR(S) SHALL REMAIN IN THE PROTECTIVE VEHICLE(S) WITH THE SAFETY BELT AND HEADREST PROPERLY ADJUSTED, MAINTAIN VEHICLE SPACING, AND KEEP THE WHEELS ALIGNED WITH THE LANE STRIPING. TWO-WAY RADIOS SHOULD BE USED TO COMMUNICATE BETWEEN THE OPERATOR AND THE WORK CREW.",
            "3. THERE SHALL BE NO WORKERS, EQUIPMENT, OR OTHER VEHICLES IN THE ROLL AHEAD DISTANCE.",
            "4. THIS TYPICAL SHOWS SET UP FOR RIGHT LANE CLOSURE. SIMILAR SET UP CAN BE USED FOR LEFT LANE CLOSURE WITH APPROPRIATE LANE CLOSURE SIGNS.",
        ],
    },
    "rules": [
        {"id": "operator-in-vehicle", "severity": "error", "source": "Note 2",
         "assert": "Operator remains in the protective vehicle (short duration)."},
        {"id": "shoulder-taper-only", "severity": "error", "source": "Plan",
         "assert": "Emit SHOULDER TAPER, not MERGING TAPER."},
        {"id": "fixed-gaps-500-1500-2640", "severity": "error", "source": "Plan",
         "assert": "Gaps are fixed 500/1500/2640 — no AW table."},
    ],
    "knownCodeDeviations": [],
}
write("619-212", spec212)


# ============================================================================= 114
spec114 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-114",
        "title": "WORK ZONE TRAFFIC CONTROL LANE CLOSURE PARKWAY",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "MOBILE OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-114.pdf",
        "localPdf": "Bridge/captures/619-114.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Family 4 mobile)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": (
            "Mobile parkway lane closure. 3 tables only — NO taper/buffer. "
            "Signs NYW8-33 + W20-5R. Moving-operation roll-ahead (higher than stationary). "
            "Plan spacing to W20-5R: 500' min / 2 mile max. If duration >15 min → reconfigure to 619-212."
        ),
    },
    "applicability": {
        "roadType": "Parkway",
        "roadway": "Parkway",
        "closure": "Lane closure",
        "duration": "Mobile",
        "durationDefinition": "Work that moves intermittently or continuously where work at any specific location completes within 15 minutes (Note 1).",
        "speedRangeMph": {"allowed": [45, 50, 55, 65],
                          "note": "Roll-ahead bands >=55 and 45-50; 65 uses >=55."},
        "laneWidthFt": None,
        "laneWidthNote": "No taper table — lane width not a lookup.",
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": [45, 50, 55, 65],
         "usedBy": ["114-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS NO WORKERS EXPOSED",
         ], "usedBy": ["114-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["114-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": ["114-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing.",
        "protectiveVehicle": "114-01",
        "rollAheadDistance": "114-02",
        "signSizes": "114-03",
    },
    "tables": {
        "114-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition"],
            "note": "Workers-exposed rows P, TMIA; OTHER HAZARDS rows NA.",
            "rows": d114["tables"]["114-01"]["rows"],
            "legend": {
                "P": "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY",
                "TMIA": "TMIA REQUIRED",
                "NA": "NOT APPLICABLE",
            },
        },
        "114-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "MOVING OPERATION values (based on protective vehicle speed of 15 MPH) — higher than stationary 306/212.",
            "rows": d114["tables"]["114-02"]["rows"],
            "usageNote": "MIN/MAX range.",
        },
        "114-03": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "rows": [
                {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
                {"signCode": "W20-5R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": "Mobile: ROLL AHEAD (table) + W20-5R at 500' min / 2 mi max. NYW8-33 on work vehicle. No tapers.",
        "zones": [
            {"id": "signA", "order": 1, "kind": "sign", "signCode": "W20-5R",
             "sheetLegend": "RIGHT LANE CLOSED AHEAD"},
            {"id": "gapA", "order": 2, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedRange": {"minFt": 500, "maxFt": 10560}},
             "dimensioned": True,
             "note": "Plan: 500' (MINIMUM) / 2 MILE (MAXIMUM)."},
            {"id": "protectiveVehicle", "order": 3, "kind": "symbol", "sheetLabel": "WORK VEHICLE",
             "lengthSource": None},
            {"id": "rollAheadDistance", "order": 4, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": "(SEE TABLE 114-02)",
             "lengthSource": {"table": "114-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 5, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Roll Ahead + W20-5R only. NYW8-33 is vehicle-mounted, not an order-table sign row.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                    {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W20-5R", "spacingZone": "gapA"},
                ],
                "excludedRows": [
                    {"label": "BUFFER SPACE", "reason": "No buffer on mobile sheet."},
                    {"label": "SHOULDER TAPER", "reason": "No taper."},
                    {"label": "MERGING TAPER", "reason": "No taper."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-5R", "sheetLegend": "RIGHT LANE CLOSED AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-05RA"},
            {"signCode": "NYW8-33", "sheetLegend": "LANE CLOSED", "shape": "rectangle",
             "postMounted": False, "mountedOn": "protectiveVehicle",
             "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": None},
            {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
             "mountedOn": "W20-5R", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "WORK VEHICLE / VEH #1", "required": True,
             "carriesSign": "NYW8-33"},
            {"id": "protectiveVehicle2", "sheetLabel": "VEH #2", "required": False,
             "note": "Notes 4-5: shoulder placement / visibility optimization."},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": False},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapA", "label": "500' (MINIMUM) … 2 MILE (MAXIMUM)"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. MOBILE WORK IS WORK THAT MOVES INTERMITTENTLY OR CONTINUOUSLY WHERE WORK AT ANY SPECIFIC LOCATION COMPLETES WITHIN 15 MINUTES.",
            "2. SHOULD THE WORK DURATION CONTINUE ON LONGER THAN THE 15 MINUTE MAXIMUM THE WORK ZONE TRAFFIC CONTROL SETUP SHALL BE RECONFIGURED AND ADJUSTED TO MEET THE REQUIREMENTS OF STANDARD SHEET 619-212.",
            "3. THIS TYPICAL MAY BE USED FOR VEHICLE BASED OPERATIONS SUCH AS SETTING UP STATIONARY TRAFFIC CONTROL (PLACING CONES, DRUMS AND SIGNS), BUT IS NOT TO BE USED FOR OPERATIONS THAT INVOLVE WORKERS ON FOOT PERFORMING ROADWAY AND / OR APPURTENANCE REPAIRS.",
            "4. IF SHOULDER AREA BECOMES TOO NARROW FOR VEH #1 VEHICLE TO BE COMPLETELY ON THE SHOULDER, THE VEHICLE SHALL STAY ON THE WIDER SHOULDER AREA UNTIL THE OPERATOR CAN SAFELY DRIVE AROUND THE NARROW SHOULDER TO NEW SET-UP POINT.",
            "5. VEH #2 SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
        ],
    },
    "rules": [
        {"id": "fifteen-minute-cap", "severity": "error", "source": "Notes 1-2",
         "assert": "Mobile duration <=15 min; else use 619-212."},
        {"id": "no-tapers", "severity": "error", "source": "Plan",
         "assert": "No merging/shoulder/downstream taper rows."},
        {"id": "roll-ahead-min-500-spacing", "severity": "warning", "source": "Plan",
         "assert": "Advance W20-5R spacing is 500' min / 2 mile max."},
    ],
    "knownCodeDeviations": [],
}
write("619-114", spec114)


# ============================================================================= 041
spec041 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-041",
        "title": "WORK ZONE TRAFFIC CONTROL SHOULDER CLOSURE/LANE ENCROACHMENT NON-FREEWAY",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "MOVING OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-041.pdf",
        "localPdf": "Bridge/captures/619-041.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Family 4 mowing / parkway-adjacent)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": (
            "Mowing / continuously-moving shoulder closure or lane encroachment on NON-FREEWAY. "
            "3 tables. W8-23 only. NON-FREEWAY PV matrix with speed bands. "
            "Moving roll-ahead including <=40. No tapers. Work area <=40 ft (Note 4). "
            "If duration >5 min → reconfigure to 619-201."
        ),
    },
    "applicability": {
        "roadType": "Non-Freeway",
        "roadway": "Non-freeway (parkway-adjacent / mowing)",
        "closure": "Shoulder closure / lane encroachment",
        "duration": "Moving",
        "durationDefinition": "Continuously moving or operations stopping for no more than 5 minutes (Note 1).",
        "speedRangeMph": {"allowed": [30, 35, 40, 45, 50, 55, 65],
                          "note": "PV speed bands cover <=30/35-40/>=45; roll-ahead has <=40/45-50/>=55."},
        "laneWidthFt": None,
        "shoulderWidthBands": SH_BANDS,
        "areaTypes": None,
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55, 65], "usedBy": ["041-01", "041-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft", "usedBy": []},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS NO WORKERS EXPOSED",
         ], "usedBy": ["041-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": ["041-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["041-03"]},
    ],
    "tableRoles": {
        "note": "3 tables. NO taperAndBuffer / advanceWarningSpacing. 041-01 uses NON-FREEWAY speedBands.",
        "protectiveVehicle": "041-01",
        "rollAheadDistance": "041-02",
        "signSizes": "041-03",
    },
    "tables": {
        "041-01": {
            "title": "PROTECTIVE VEHICLE REQUIREMENTS",
            "confidence": "verbatim",
            "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
            "note": "NON-FREEWAY only (no FREEWAY column). OTHER HAZARDS rows all NA.",
            "speedBands": d041["tables"]["041-01"]["speedBands"],
            "rows": d041["tables"]["041-01"]["rows"],
            "legend": {
                "P": "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY",
                "TMIA": "TMIA REQUIRED",
                "NA": "NOT APPLICABLE",
            },
        },
        "041-02": {
            "title": "ROLL AHEAD DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "note": "MOVING OPERATION (based on protective vehicle speed of 15 MPH). 3 bands including <=40.",
            "rows": d041["tables"]["041-02"]["rows"],
            "usageNote": "MIN/MAX range.",
        },
        "041-03": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W8-23", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
            ],
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": "Moving mowing: ROLL AHEAD + W8-23. Lateral 10'-0\" MIN. Work area <=40 ft. No tapers.",
        "zones": [
            {"id": "signA", "order": 1, "kind": "sign", "signCode": "W8-23",
             "sheetLegend": "NO SHOULDER / LOW SHOULDER"},
            {"id": "gapA", "order": 2, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 500}, "dimensioned": False,
             "note": "No numbered advance-gap callout; nominal spacing for order-table resolve."},
            {"id": "protectiveVehicle", "order": 3, "kind": "symbol", "sheetLabel": "WORK VEHICLE",
             "lengthSource": None},
            {"id": "rollAheadDistance", "order": 4, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": "(SEE TABLE 041-02)",
             "lengthSource": {"table": "041-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 5, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False,
             "maxLengthFt": 40,
             "note": "Note 4: work area shall not exceed 40 feet in length."},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "Roll Ahead + W8-23 only.",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                    {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W8-23", "spacingZone": "gapA"},
                ],
                "excludedRows": [
                    {"label": "BUFFER SPACE", "reason": "No buffer."},
                    {"label": "SHOULDER TAPER", "reason": "No taper."},
                    {"label": "MERGING TAPER", "reason": "No taper."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W8-23", "sheetLegend": "NO SHOULDER", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W08-23"},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "WORK VEHICLE / 24000 LB PROTECTIVE", "required": True},
            {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": False},
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        ],
        "lateralDimensions": [
            {"label": "10'-0\" (MIN.)", "note": "Minimum open width beside work."},
        ],
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THIS TYPICAL APPLIES TO CONTINUOUSLY MOVING OR OPERATIONS STOPPING FOR NO MORE THAN 5 MINUTES.",
            "2. SHOULD THE WORK DURATION EXCEED 5 MINUTES, THE WZTC SETUP SHALL BE RECONFIGURED TO MEET THE REQUIREMENTS OF STANDARD SHEET 619-201.",
            "3. THIS STANDARD SHEET MAY BE USED FOR OPERATIONS SUCH AS SETTING UP STATIONARY TRAFFIC CONTROL (E.G., PLACING CONES, DRUMS AND SIGNS) AND DEBRIS REMOVAL.",
            "4. WORK AREA SHALL NOT EXCEED 40 FEET IN LENGTH.",
            "5. THE PROTECTIVE VEHICLE SHALL BE PLACED TO OPTIMIZE AND ENHANCE VISIBILITY FROM THE REAR OF THE OPERATION. AND SHALL NOT EXCEED THE APPROPRIATE ROLL AHEAD DISTANCE VALUES.",
            "6. THE PROTECTIVE VEHICLE SHALL STAY AS FAR RIGHT AS POSSIBLE AND SHALL ADJUST ITS SPACING TO ACCOMMODATE CHANGING SIGHT DISTANCE AND OTHER FIELD CONDITIONS.",
            "7. UNLESS THIS SETUP IS BEING USED FOR AN EMERGENCY SITUATION, WORK SHOULD BE SCHEDULED DURING NON-PEAK HOURS.",
        ],
    },
    "rules": [
        {"id": "five-minute-cap", "severity": "error", "source": "Notes 1-2",
         "assert": "Moving duration <=5 min; else use 619-201."},
        {"id": "work-area-max-40", "severity": "error", "source": "Note 4",
         "assert": "Work area length <= 40 ft."},
        {"id": "no-tapers", "severity": "error", "source": "Plan",
         "assert": "No taper/buffer rows."},
    ],
    "knownCodeDeviations": [],
}
write("619-041", spec041)

print("All 4 specs written.")
