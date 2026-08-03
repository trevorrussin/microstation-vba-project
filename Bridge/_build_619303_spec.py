"""Build Data/sheet-specs/619-303.json as a Family-2 sibling of 619-302.

Structural differences (NOT a blind clone — AUTHORING.md slow-down case):
- TWO-LANE closure: W20-5aR, two successive MERGING TAPER L with a 2L span between
- Table numbering: 02=roll ahead, 04=taper+buffer, 05=sign sizes (reversed roles vs 302's 02/05)
- 9 plan notes; Note 5 is about W20-5a / W4-2 symmetry; Note 9 = VEH #2 shoulder rule
- Two arrow panels (one at each merging taper)
- Five protective-vehicle callouts (VEH #1..#5)
Tables 303-01/02/03/04 verified cell-identical to 302-01/05/03/02; only 303-05
swaps W20-5R -> W20-5aR.
"""
from __future__ import annotations

import copy
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
ref = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
s = copy.deepcopy(ref)

# ---- identity ----
s["sheet"] = {
    "number": "619-303",
    "title": "WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT TWO-LANE CLOSURE",
    "series": "WORK ZONE TRAFFIC CONTROL",
    "operation": "SHORT TERM OPERATION",
    "units": "U.S. CUSTOMARY",
    "scale": "NOT TO SCALE",
    "approved": "2021-12-02",
    "issuedUnder": "EI 21-028",
    "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
    "sourceUrl": "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-303.pdf",
    "localPdf": "Bridge/captures/619-303.pdf",
    "localRender": "Bridge/captures/sheet_619303_plan.png",
    "pdfPages": 1,
    "transcribedBy": "Cursor (Family 2 sibling of 619-302; corridor/anchors/notes direct from extract_plan_geometry + plan crop; tables cross-checked cell-identical to 302 except sign-size W20-5aR)",
    "transcribedOn": "2026-08-03",
    "provenanceNote": "Two-lane closure variant of Family 2 reference 619-302. Same DATUM SHARING (shoulder taper inside gap A). Dual merging tapers with a dimensioned 2L span between them — not a single MERGING TAPER. Tables 303-01/02/03/04 verified identical to 302-01/05/03/02; 303-05 replaces W20-5R with W20-5aR.",
}

s["applicability"]["closure"] = "Right two-lane closure"
s["applicability"]["closureNote"] = "Closes the two rightmost travel lanes. Note 5: left two-lane closures are symmetrical — substitute W20-5aL and W4-2L."

# ---- tableRoles (content, not suffix) ----
s["tableRoles"] = {
    "note": "Numbering trap vs 302: on 303, 02=ROLL AHEAD, 04=BUFFER+TAPER, 05=SIGN SIZES (302 had 02=taper, 04=sizes, 05=roll ahead).",
    "protectiveVehicle": "303-01",
    "rollAheadDistance": "303-02",
    "advanceWarningSpacing": "303-03",
    "taperAndBuffer": "303-04",
    "signSizes": "303-05",
}

# Remap tables: copy content, rename keys
old = s["tables"]
s["tables"] = {
    "303-01": copy.deepcopy(old["302-01"]),
    "303-02": copy.deepcopy(old["302-05"]),
    "303-03": copy.deepcopy(old["302-03"]),
    "303-04": copy.deepcopy(old["302-02"]),
    "303-05": copy.deepcopy(old["302-04"]),
}
s["tables"]["303-01"]["title"] = "PROTECTIVE VEHICLE REQUIREMENTS"
s["tables"]["303-02"]["title"] = "ROLL AHEAD DISTANCE"
s["tables"]["303-03"]["title"] = "ADVANCE WARNING SIGN SPACING"
s["tables"]["303-04"]["title"] = "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS"
s["tables"]["303-05"]["title"] = "REQUIRED SIGN SIZES"
# Fix knownAnomalies table refs inside copied taper table
if "knownAnomalies" in s["tables"]["303-04"]:
    for a in s["tables"]["303-04"]["knownAnomalies"]:
        a["note"] = a.get("note", "").replace("302-02", "303-04").replace("619-302", "619-303")
# Sign sizes: W20-5R -> W20-5aR
for row in s["tables"]["303-05"]["rows"]:
    if row["signCode"] == "W20-5R":
        row["signCode"] = "W20-5aR"
s["tables"]["303-05"]["note"] = (
    "Identical to 619-302's 302-04 except W20-5R is replaced by W20-5aR "
    "(RIGHT TWO LANES CLOSED) — verified from PDF word extraction. WARNING FLAG row present."
)

# Update inputs usedBy table ids
for inp in s["inputs"]:
    ub = inp.get("usedBy", [])
    inp["usedBy"] = [u.replace("302-", "303-") for u in ub]

# ---- corridor ----
s["corridor"] = {
    "confidence": "drawing",
    "description": (
        "Zones in DIRECTION OF TRAVEL. Derived from scripts/extract_plan_geometry.py "
        "(main dim column x=281.8 + left taper column x=154.5) and confirmed on "
        "Bridge/captures/sheet_619303_plan.png. TWO successive MERGING TAPER L runs "
        "with a dimensioned 2L straight span between them; shoulder taper overlays gap A."
    ),
    "zones": [
        {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK XX"},
        {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C", "sheetReference": "(SEE TABLE 303-03)",
         "lengthSource": {"table": "303-03", "column": "C", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True, "spans": "W20-1 to W20-5aR"},
        {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-5aR",
         "sheetLegend": "2 RIGHT LANES CLOSED YY"},
        {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B", "sheetReference": "(SEE TABLE 303-03)",
         "lengthSource": {"table": "303-03", "column": "B", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True, "spans": "W20-5aR to W4-2R"},
        {"id": "signA", "order": 5, "kind": "sign", "signCode": "W4-2R",
         "sheetLegend": "(merge symbol)", "note": "Note 5: substitute W20-5aL / W4-2L for left two-lane closure."},
        {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A", "sheetReference": "(SEE TABLE 303-03)",
         "lengthSource": {"table": "303-03", "column": "A", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True,
         "spans": "W4-2R to the upstream end of the first (upstream) MERGING TAPER L",
         "containsOverlay": "shoulderTaper",
         "note": "DATUM SHARING at y=575.8: SHOULDER TAPER (x=154.5, 575.8-600.5) shares datum with A (x=281.8, 575.8-640.3) — overlay inside A, same pattern as 302/311."},
        {"id": "shoulderTaper", "order": 7, "kind": "taper", "sheetLabel": "SHOULDER TAPER",
         "sheetReference": "(SEE TABLE 303-04)",
         "lengthSource": {"table": "303-04", "column": "shoulderTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
         "dimensioned": True, "consumesStation": False, "containedIn": "gapA",
         "stationAnchor": {"zone": "mergingTaperUpstream", "end": "upstream"},
         "note": "L/3 overlay on gap A."},
        {"id": "mergingTaperUpstream", "order": 8, "kind": "taper", "sheetLabel": "MERGING TAPER",
         "sheetReference": "(SEE TABLE 303-04)",
         "lengthSource": {"table": "303-04", "column": "laneTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
         "dimensioned": True,
         "note": "First MERGING TAPER L a driver encounters — closes the outer (rightmost) travel lane. Length L from Table 303-04. Vector segment y 505.1-575.8 on x=154.5."},
        {"id": "taperGap2L", "order": 9, "kind": "gap", "sheetLabel": "2L",
         "sheetReference": "(SEE TABLE 303-04)",
         "lengthSource": {"table": "303-04", "column": "laneTaper", "scale": 2,
                          "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
         "dimensioned": True,
         "note": "Straight channelizing span between the two merging tapers. Sheet labels it 2L — length is 2× the L value from Table 303-04 (not a separate table column). Vector segment y 382.8-505.1 on x=281.8."},
        {"id": "mergingTaperDownstream", "order": 10, "kind": "taper", "sheetLabel": "MERGING TAPER",
         "sheetReference": "(SEE TABLE 303-04)",
         "lengthSource": {"table": "303-04", "column": "laneTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
         "dimensioned": True,
         "note": "Second MERGING TAPER L — closes the next travel lane into the work area. Same L as the upstream taper. Vector segment y 321.5-382.8 on x=154.5."},
        {"id": "bufferSpace", "order": 11, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
         "sheetReference": "(SEE TABLE 303-04)",
         "lengthSource": {"table": "303-04", "column": "longitudinalBufferSpace",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True, "mustBeEmptyReason": "Note 3."},
        {"id": "protectiveVehicleCluster", "order": 12, "kind": "symbol",
         "sheetLabel": "VEH #3 / #4 / #5", "lengthSource": None,
         "note": "Three protective vehicles spanning the two closed lanes at the upstream end of the roll ahead distance (plan labels VEH #3/#4/#5 in the roll-ahead band)."},
        {"id": "rollAheadDistance", "order": 13, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
         "sheetReference": "(SEE TABLE 303-02)",
         "lengthSource": {"table": "303-02", "column": "range",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True, "mustBeEmptyReason": "Note 3.",
         "spans": "Protective vehicle cluster to the upstream edge of the WORK AREA"},
        {"id": "workArea", "order": 14, "kind": "workArea", "sheetLabel": "WORK AREA",
         "lengthSource": None, "lengthNote": "Project-specific.", "hatched": True, "dimensioned": False},
        {"id": "spotter", "order": 15, "kind": "symbol", "sheetLabel": "SPOTTER RECOMMENDED",
         "lengthSource": None, "required": False},
        {"id": "downstreamTaper", "order": 16, "kind": "taper", "sheetLabel": "DOWNSTREAM TAPER",
         "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}}, "sheetText": "50'-100'",
         "dimensioned": True},
        {"id": "gapEndRoadWork", "order": 17, "kind": "gap", "sheetLabel": None,
         "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False,
         "sheetText": "THIS SIGN SHALL BE LOCATED A MINIMUM DISTANCE OF 80 FT AND MAXIMUM OF 400 FT PAST THE END OF THE DOWNSTREAM TAPER.",
         "spans": "End of the downstream taper to G20-2"},
        {"id": "signEndRoadWork", "order": 18, "kind": "sign", "signCode": "G20-2",
         "sheetLegend": "END ROAD WORK"},
    ],
}

s["orderTable"] = {
    "confidence": "drawing",
    "description": "Upstream walk includes BOTH merging tapers and the 2L span between them. Shoulder taper remains an overlay on gap A.",
    "alignments": [
        {
            "alignIdx": 1,
            "name": "Upstream",
            "station0": "Upstream edge of the WORK AREA",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                {"rowNum": 3, "type": "Non-Sign", "zone": "mergingTaperDownstream", "label": "MERGING TAPER"},
                {"rowNum": 4, "type": "Non-Sign", "zone": "taperGap2L", "label": "2L"},
                {"rowNum": 5, "type": "Non-Sign", "zone": "mergingTaperUpstream", "label": "MERGING TAPER"},
                {"rowNum": 6, "type": "Sign", "zone": "signA", "signCode": "W4-2R", "spacingZone": "gapA"},
                {"rowNum": 7, "type": "Sign", "zone": "signB", "signCode": "W20-5aR", "spacingZone": "gapB"},
                {"rowNum": 8, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
            ],
            "overlayZones": [
                {"zone": "shoulderTaper", "anchor": {"zone": "mergingTaperUpstream", "end": "upstream"},
                 "direction": "upstream",
                 "note": "Overlay inside gap A — same DATUM SHARING pattern as 302/311."}
            ],
            "excludedRows": [
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
}

s["signs"] = {
    "confidence": "verbatim",
    "note": "Legend substitution from Table 303-03 XX/YY. Mid advance sign is W20-5aR (two-lane), SignLibrary base W20-05aR.",
    "items": [
        {
            "signCode": "W20-1", "sheetLegend": "ROAD WORK XX",
            "legendSubstitution": {"placeholder": "XX", "table": "303-03", "column": "XX"},
            "shape": "diamond", "warningFlags": True, "postMounted": True,
            "corridorZone": "signC", "positionRank": 3,
            "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
            "signLibraryKey": None, "signLibraryBase": "W20-01R",
        },
        {
            "signCode": "W20-5aR", "sheetLegend": "2 RIGHT LANES CLOSED YY",
            "legendSubstitution": {"placeholder": "YY", "table": "303-03", "column": "YY"},
            "shape": "diamond", "warningFlags": False, "postMounted": True,
            "corridorZone": "signB", "positionRank": 2,
            "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
            "signLibraryKey": None, "signLibraryBase": "W20-05aR",
            "signLibraryNote": "Two-lane family: W20-05aRA / W20-05aRF / W20-05aRM. Note 5: left closure uses W20-5aL.",
        },
        {
            "signCode": "W4-2R", "sheetLegend": "(merge symbol)", "legendSubstitution": None,
            "shape": "diamond", "warningFlags": True, "postMounted": True,
            "corridorZone": "signA", "positionRank": 1,
            "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W04-02R",
        },
        {
            "signCode": "G20-2", "sheetLegend": "END ROAD WORK", "legendSubstitution": None,
            "shape": "rectangle", "warningFlags": False, "postMounted": True,
            "corridorZone": "signEndRoadWork",
            "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02",
        },
        {
            "signCode": "NYW8-33", "sheetLegend": "LANE CLOSED", "legendSubstitution": None,
            "shape": "rectangle", "warningFlags": False, "postMounted": False,
            "mountedOn": "protectiveVehicleCluster", "signLibraryKey": None,
            "sizeNonFreeway": "48x24", "sizeFreeway": "48x24",
        },
        {
            "signCode": "WARNING FLAG", "sheetLegend": None, "shape": "flag", "postMounted": False,
            "mountedOn": "W20-1, W4-2R", "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
        },
    ],
}

s["symbols"] = {
    "confidence": "drawing",
    "items": [
        {
            "id": "arrowPanelUpstream", "sheetLabel": "ARROW PANEL", "required": True, "count": 1,
            "stationAnchor": {"zone": "mergingTaperUpstream", "end": "upstream"},
            "lateralAnchor": "Shoulder / outer edge at head of first merging taper",
            "alternative": {"sheetText": "OR", "option": "VEH #1",
                            "description": "Truck-mounted arrow panel alternative at same station."},
        },
        {
            "id": "arrowPanelDownstream", "sheetLabel": "ARROW PANEL", "required": True, "count": 1,
            "stationAnchor": {"zone": "mergingTaperDownstream", "end": "upstream"},
            "lateralAnchor": "At head of second merging taper (in the already-closed outer lane)",
            "note": "Second arrow panel — two-lane closures show one panel per merging taper.",
        },
        {
            "id": "protectiveVehicle1", "sheetLabel": "VEH #1",
            "required": "only when truck-mounted arrow panel option used at first taper",
            "stationAnchor": {"zone": "mergingTaperUpstream", "end": "upstream"},
            "cellHint": "TWZWVA_P",
        },
        {
            "id": "protectiveVehicle2", "sheetLabel": "VEH #2",
            "required": "per Note 9 — only when shoulder width is >= 8 ft",
            "stationAnchor": {"zone": "mergingTaperDownstream", "end": "upstream",
                              "note": "Plan places VEH #2 near the second taper / buffer; Note 9 is the shoulder-width gate."},
            "lateralAnchor": "Closed paved shoulder", "cellHint": "TWZWVA_P",
            "confidenceNote": "Medium — label and Note 9 confirmed; exact station vs second taper tip should be visually confirmed before live placement relies on it.",
        },
        {
            "id": "protectiveVehicleCluster", "sheetLabel": "VEH #3 / #4 / #5",
            "required": "per Table 303-01 (one PV per closed lane + conditional shoulder)",
            "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"},
            "lateralAnchor": "Across the two closed travel lanes", "cellHint": "TWZWVA_P",
            "carriesSign": "NYW8-33",
        },
        {
            "id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "deviceSymbol": "CONE",
            "required": True,
            "longitudinalSpacing": {"maxFt": 40, "sheetText": "Note 7 — not to exceed 40' in the active work space."},
            "runs": [
                {"id": "shoulderTaperRun", "zone": "shoulderTaper",
                 "deviceCountSource": {"table": "303-04", "column": "shoulderTaper.devices"}},
                {"id": "mergingTaperUpstreamRun", "zone": "mergingTaperUpstream",
                 "deviceCountSource": {"table": "303-04", "column": "laneTaper.devices"}},
                {"id": "taperGap2LRun", "zone": "taperGap2L", "deviceCountSource": None},
                {"id": "mergingTaperDownstreamRun", "zone": "mergingTaperDownstream",
                 "deviceCountSource": {"table": "303-04", "column": "laneTaper.devices"}},
                {"id": "longitudinalRun", "zone": "bufferSpace..workArea", "deviceCountSource": None},
                {"id": "downstreamRun", "zone": "downstreamTaper", "deviceCountSource": None},
            ],
            "transverse": {
                "required": "conditional",
                "condition": "Paved shoulder 8' or wider closed > 800'",
                "maxSpacingFt": 800, "sheetText": "Note 6", "detail": "303A",
            },
        },
        {"id": "spotter", "sheetLabel": "SPOTTER RECOMMENDED", "required": False,
         "stationAnchor": {"zone": "workArea", "end": "downstream"}, "detail": "303A"},
        {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True,
         "stationAnchor": {"zone": "workArea", "end": "both"},
         "lateralAnchor": "Both closed travel lanes (plus closed shoulder if applicable)", "hatched": True},
    ],
}

s["annotations"] = {
    "confidence": "drawing",
    "dimensions": [
        {"zone": "gapC", "label": "C", "reference": "(SEE TABLE 303-03)"},
        {"zone": "gapB", "label": "B", "reference": "(SEE TABLE 303-03)"},
        {"zone": "gapA", "label": "A", "reference": "(SEE TABLE 303-03)"},
        {"zone": "shoulderTaper", "label": "SHOULDER TAPER", "reference": "(SEE TABLE 303-04)"},
        {"zone": "mergingTaperUpstream", "label": "MERGING TAPER", "reference": "(SEE TABLE 303-04)"},
        {"zone": "taperGap2L", "label": "2L", "reference": "(SEE TABLE 303-04)"},
        {"zone": "mergingTaperDownstream", "label": "MERGING TAPER", "reference": "(SEE TABLE 303-04)"},
        {"zone": "bufferSpace", "label": "BUFFER SPACE", "reference": "(SEE TABLE 303-04)"},
        {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": "(SEE TABLE 303-02)"},
        {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER", "reference": None},
    ],
    "labels": [
        {"text": "WORK AREA", "zone": "workArea"},
        {"text": "SPOTTER RECOMMENDED", "zone": "spotter"},
        {"text": "ARROW PANEL", "zone": "mergingTaperUpstream"},
        {"text": "ARROW PANEL", "zone": "mergingTaperDownstream"},
    ],
}

s["details"] = {
    "303A": {
        "title": "DETAIL 303A",
        "note": "Referenced by Note 6 (transverse channelizing when shoulder >= 8' closed > 800'). Spotter also shown.",
    }
}

s["notes"] = {
    "confidence": "verbatim",
    "printed": [
        "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS. IF THERE IS A CONFLICT, MOVE THE SIGN UPSTREAM.",
        "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "4. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
        "5. LEFT LANE CLOSURES ARE SYMMETRICAL TO RIGHT LANE CLOSURES. SUBSTITUTE THE W20-5a SIGN AND THE CORRESPONDING W4-2 SIGN.",
        "6. CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF EVERY 800' AS SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR GREATER IS CLOSED FOR A DISTANCE GREATER THAN 800'.",
        "7. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN THE ACTIVE WORK SPACE.",
        "8. THE MINIMUM LANE WIDTH SHALL BE 11' FOR FREEWAYS AND 10' FOR NON-FREEWAYS.",
        "9. VEH #2 IS ONLY NEEDED WHEN THE SHOULDER WIDTH IS >= 8'.",
    ],
    "planCallouts": [
        {"text": "THIS SIGN SHALL BE LOCATED A MINIMUM DISTANCE OF 80 FT AND MAXIMUM OF 400 FT PAST THE END OF THE DOWNSTREAM TAPER.", "appliesTo": "G20-2"},
        {"text": "CONE SPACING NOT TO EXCEED 40 FT. (1 SKIP LINE)", "appliesTo": "channelizingDevices"},
    ],
    "tableNotes": {
        "303-01": [
            "1. THE EXPOSURE CONDITIONS ASSUMES THERE IS NO POSITIVE PROTECTION PRESENT",
            "2. EITHER A PROTECTIVE VEHICLE OR THE STANDARD BUFFER SPACE SHALL BE PROVIDED",
        ]
    },
}

s["rules"] = [
    {"id": "no-occupancy-buffer-rollahead", "severity": "error", "source": "Note 3",
     "assert": "No work area hatch, protective vehicle, equipment or other symbol may overlap the bufferSpace or rollAheadDistance zones.",
     "commonFailure": "Drawing the work area hatch from station 0 through the roll ahead distance."},
    {"id": "sign-order", "severity": "error", "source": "Plan layout",
     "assert": "Walking upstream from the first merging taper, advance signs appear W4-2R, then W20-5aR, then W20-1.",
     "commonFailure": "Using W20-5R (one-lane) instead of W20-5aR (two-lane), or reversing the order."},
    {"id": "two-merging-tapers-plus-2L", "severity": "error", "source": "Plan layout",
     "assert": "Upstream walk contains two MERGING TAPER rows separated by a 2L row whose length equals 2 * L from Table 303-04.",
     "commonFailure": "Cloning 619-302's single MERGING TAPER walk and omitting the 2L span / second taper."},
    {"id": "sign-is-two-lane", "severity": "error", "source": "Plan + Table 303-05",
     "assert": "Mid advance sign is W20-5aR with SignLibrary base W20-05aR (A/F/M variants).",
     "commonFailure": "Emitting W20-05R* keys from the one-lane Family 2 reference."},
    {"id": "shoulder-taper-is-an-overlay", "severity": "error", "source": "Dimension line datums (PDF vector layer)",
     "assert": "The shoulder taper starts where the upstream merging taper ends and runs upstream INSIDE gap A. It consumes no station of its own.",
     "commonFailure": "Treating the shoulder taper as a sequential station row, which pushes every advance sign upstream by L/3."},
    {"id": "no-invented-zones", "severity": "error", "source": "Plan layout",
     "assert": "No Vehicle Space, temporary barrier, or box/corrugated beam row in the order table.",
     "commonFailure": "Emitting the generic 7-row default upstream table."},
    {"id": "veh2-is-conditional", "severity": "error", "source": "Note 9",
     "assert": "VEH #2 is only placed when the closed shoulder width is >= 8 ft.",
     "commonFailure": "Always placing VEH #2, or treating Note 9 as optional guidance."},
    {"id": "two-arrow-panels", "severity": "error", "source": "Plan layout",
     "assert": "Two arrow panels — one at the upstream end of each merging taper.",
     "commonFailure": "Placing only the single arrow panel that 619-302 shows."},
    {"id": "left-closure-sign-substitution", "severity": "warning", "source": "Note 5",
     "assert": "For a LEFT two-lane closure, substitute W20-5aL and W4-2L.",
     "commonFailure": "Hardcoding right-lane sign codes regardless of side."},
    {"id": "cone-spacing", "severity": "warning", "source": "Note 7",
     "assert": "Channelizing device spacing (center to center) must not exceed 40 ft in the active work space.",
     "commonFailure": "Drawing a bare polyline with no spacing annotation or device count."},
    {"id": "transverse-devices", "severity": "warning", "source": "Note 6 / Detail 303A",
     "assert": "When a paved shoulder 8 ft or wider is closed for more than 800 ft, place transverse channelizing devices at least every 800 ft.",
     "commonFailure": "Omitting Detail 303A on long shoulder closures."},
]

s["knownCodeDeviations"] = [
    {"id": "dual-taper-placement", "severity": "error",
     "assert": "Current PlaceOrderTableChannelizing / symbol placement assumes a single merging taper; 619-303 needs two taper runs + 2L longitudinal run + two arrow panels."},
    {"id": "sheet-registry-wrong-signs", "severity": "warning",
     "assert": "Data/sheet-registry.tsv for 619-303 must not be trusted over this spec (book-PDF noise pattern seen on 302)."},
]

s["knownExcerpts"] = {
    "from619-302": [
        "303-01 == 302-01 (protective vehicle)",
        "303-02 == 302-05 (roll ahead)",
        "303-03 == 302-03 (advance warning)",
        "303-04 == 302-02 (taper+buffer), including 65mph/12ft = 800/20/21",
    ],
    "differsFrom302": [
        "303-05 has W20-5aR instead of W20-5R",
        "Corridor has two MERGING TAPER L + 2L span",
        "9 notes vs 8; Note 5/9 content differ",
        "Two arrow panels; VEH #1..#5",
    ],
}

out = ROOT / "Data/sheet-specs/619-303.json"
out.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("wrote", out, "bytes", out.stat().st_size)
