"""Build complete Family 5 sheet specs from drafts + 302/301 corridor patterns."""
from __future__ import annotations

import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
SPEC = ROOT / "Data/sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"
SH3 = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SH7 = ["<= 4 ft", "5 - 7 ft", ">= 8 ft", "9 ft", "10 ft", "11 ft", "12 ft"]
SPEEDS = [45, 50, 55, 65]


def load_draft(n: int) -> dict:
    return json.loads((SPEC / f"_draft_619{n}_tables.json").read_text(encoding="utf-8"))


def write(name: str, spec: dict) -> None:
    path = SPEC / f"{name}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("Wrote", path.name)


def sheet_meta(n: int, title: str, op: str, pages: int, rot: int, note: str) -> dict:
    return {
        "number": f"619-{n}",
        "title": title,
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": op,
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2022-04-08" if n >= 300 else "2021-12-02",
        "issuedUnder": "EI 22-008" if n >= 300 else "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-{n}.pdf",
        "localPdf": f"Bridge/captures/619-{n}.pdf",
        "localRender": None,
        "pdfPages": pages,
        "pageRotation": rot,
        "transcribedBy": "Cursor (Family 5 ramp-adjacent)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": note,
    }


def inputs_lane(prefix: str, bands=None, has_lane=True):
    bands = bands or SH3
    out = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
         "usedBy": [f"{prefix}-01", f"{prefix}-02"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": bands, "default": ">= 8 ft",
         "usedBy": [f"{prefix}-01"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                     "OTHER HAZARDS NO WORKERS EXPOSED"],
         "usedBy": [f"{prefix}-04" if prefix == "318" else f"{prefix}-pv"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT",
         "usedBy": [f"{prefix}-04" if prefix == "318" else f"{prefix}-pv"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": [f"{prefix}-06" if prefix == "318" else f"{prefix}-sign"]},
    ]
    if has_lane:
        out.insert(1, {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12],
                       "usedBy": [f"{prefix}-01"]})
    return out


def size_row(code, fw, nf=None):
    return {"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw}


def sign_item(code, legend, key, zone, size="48x48", post=True, mounted=None, flags=False, note=None):
    s = {
        "signCode": code,
        "sheetLegend": legend,
        "legendSubstitution": None,
        "shape": "diamond" if code.startswith("W") and code != "WARNING FLAG" else "rectangle",
        "warningFlags": flags,
        "postMounted": post,
        "corridorZone": zone,
        "sizeNonFreeway": size if nf_default(size) else None,
        "sizeFreeway": size,
        "signLibraryKey": key,
    }
    if mounted:
        s["mountedOn"] = mounted
    if note:
        s["note"] = note
    if code == "WARNING FLAG":
        s["shape"] = "flag"
        s["sizeNonFreeway"] = "18x18"
    if code == "G20-2":
        s["sizeNonFreeway"] = "36x18"
        s["sizeFreeway"] = "48x24"
    if code == "NYW8-33":
        s["sizeNonFreeway"] = "48x24"
        s["sizeFreeway"] = "48x24"
        s["shape"] = "rectangle"
    return s


def nf_default(size):
    return True


def corridor_merging_fixed(gaps: dict, taper_table: str, roll_table: str,
                           two_w20: bool = True, taper_label: str = "MERGING TAPER",
                           shoulder_overlay: bool = True):
    """302-like corridor with fixed A/B/C(/D) gaps (no AW spacing table)."""
    zones = [
        {"id": "signD" if two_w20 else "signC", "order": 1, "kind": "sign", "signCode": "W20-1",
         "sheetLegend": "ROAD WORK 1 MILE"},
        {"id": "gapD" if two_w20 else "gapC", "order": 2, "kind": "gap",
         "sheetLabel": "D" if two_w20 else "C",
         "lengthSource": {"fixedFt": gaps.get("D", gaps.get("C", 1320))},
         "dimensioned": True,
         "spans": "W20-1 (1 MILE) to W20-1 (½ MILE)" if two_w20 else "W20-1 to W20-5"},
    ]
    order = 3
    if two_w20:
        zones += [
            {"id": "signC", "order": order, "kind": "sign", "signCode": "W20-1",
             "sheetLegend": "ROAD WORK ½ MILE",
             "note": "Second W20-1; SignLibrary key W20-01RPM preferred — order table reuses W20-1."},
            {"id": "gapC", "order": order + 1, "kind": "gap", "sheetLabel": "C",
             "lengthSource": {"fixedFt": gaps["C"]}, "dimensioned": True,
             "spans": "W20-1 (½ MILE) to W20-5"},
        ]
        order += 2
    zones += [
        {"id": "signB", "order": order, "kind": "sign", "signCode": "W20-5",
         "sheetLegend": "RIGHT LANE CLOSED ½ MILE"},
        {"id": "gapB", "order": order + 1, "kind": "gap", "sheetLabel": "B",
         "lengthSource": {"fixedFt": gaps["B"]}, "dimensioned": True,
         "spans": "W20-5 to W4-2R"},
        {"id": "signA", "order": order + 2, "kind": "sign", "signCode": "W4-2R",
         "sheetLegend": "(merge symbol)"},
        {"id": "gapA", "order": order + 3, "kind": "gap", "sheetLabel": "A",
         "lengthSource": {"fixedFt": gaps["A"]}, "dimensioned": True,
         "spans": f"W4-2R to upstream end of {taper_label}",
         "containsOverlay": "shoulderTaper" if shoulder_overlay else None,
         "note": "Fixed plan callout — no advance-warning spacing table."},
    ]
    order = order + 4
    if shoulder_overlay:
        zones.append({
            "id": "shoulderTaper", "order": order, "kind": "taper",
            "sheetLabel": "SHOULDER TAPER", "sheetReference": f"(SEE TABLE {taper_table})",
            "lengthSource": {"table": taper_table, "column": "shoulderTaper",
                             "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
            "dimensioned": True, "consumesStation": False, "containedIn": "gapA",
            "stationAnchor": {"zone": "laneTaper", "end": "upstream"},
        })
        order += 1
    zones += [
        {"id": "laneTaper", "order": order, "kind": "taper", "sheetLabel": taper_label,
         "sheetReference": f"(SEE TABLE {taper_table})",
         "lengthSource": {"table": taper_table, "column": "laneTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
         "dimensioned": True},
        {"id": "bufferSpace", "order": order + 1, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
         "sheetReference": f"(SEE TABLE {taper_table})",
         "lengthSource": {"table": taper_table, "column": "longitudinalBufferSpace",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True},
        {"id": "protectiveVehicle", "order": order + 2, "kind": "symbol",
         "sheetLabel": "VEH #1", "lengthSource": None},
        {"id": "rollAheadDistance", "order": order + 3, "kind": "clearance",
         "sheetLabel": "ROLL AHEAD DISTANCE", "sheetReference": f"(SEE TABLE {roll_table})",
         "lengthSource": {"table": roll_table, "column": "range",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True},
        {"id": "workArea", "order": order + 4, "kind": "workArea", "sheetLabel": "WORK AREA",
         "lengthSource": None, "hatched": True, "dimensioned": False},
        {"id": "downstreamTaper", "order": order + 5, "kind": "taper",
         "sheetLabel": "DOWNSTREAM TAPER",
         "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}},
         "sheetText": "50'-100'", "dimensioned": True},
        {"id": "gapEndRoadWork", "order": order + 6, "kind": "gap", "sheetLabel": None,
         "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False},
        {"id": "signEndRoadWork", "order": order + 7, "kind": "sign", "signCode": "G20-2",
         "sheetLegend": "END ROAD WORK"},
    ]
    # strip null containsOverlay
    for z in zones:
        if z.get("containsOverlay") is None:
            z.pop("containsOverlay", None)
    return zones


def order_merging(two_w20=True, shoulder_overlay=True, taper_label="MERGING TAPER"):
    rows = [
        {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
        {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
        {"rowNum": 3, "type": "Non-Sign", "zone": "laneTaper", "label": taper_label},
        {"rowNum": 4, "type": "Sign", "zone": "signA", "signCode": "W4-2R", "spacingZone": "gapA"},
        {"rowNum": 5, "type": "Sign", "zone": "signB", "signCode": "W20-5", "spacingZone": "gapB"},
    ]
    n = 6
    if two_w20:
        rows.append({"rowNum": n, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                     "spacingZone": "gapC"})
        rows.append({"rowNum": n + 1, "type": "Sign", "zone": "signD", "signCode": "W20-1",
                     "spacingZone": "gapD"})
    else:
        rows.append({"rowNum": n, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                     "spacingZone": "gapC"})
    al1 = {
        "alignIdx": 1, "name": "Upstream",
        "station0": "Upstream edge of the WORK AREA",
        "walkDirection": "Upstream, against traffic",
        "rows": rows,
        "excludedRows": [
            {"label": "Vehicle Space", "reason": "Not on this sheet."},
            {"label": "SHOULDER TAPER", "reason": "Overlay inside gap A, not a sequential row."}
            if shoulder_overlay else
            {"label": "Vehicle Space", "reason": "Not on this sheet."},
        ],
    }
    if shoulder_overlay:
        al1["overlayZones"] = [{
            "zone": "shoulderTaper",
            "anchor": {"zone": "laneTaper", "end": "upstream"},
            "direction": "upstream",
        }]
        # fix excludedRows duplicate
        al1["excludedRows"] = [
            {"label": "Vehicle Space", "reason": "Not on this sheet."},
            {"label": "SHOULDER TAPER", "reason": "Overlay inside gap A."},
        ]
    else:
        al1["excludedRows"] = [{"label": "Vehicle Space", "reason": "Not on this sheet."}]
    al2 = {
        "alignIdx": 2, "name": "Downstream",
        "station0": "Downstream edge of the WORK AREA",
        "walkDirection": "Downstream, with traffic",
        "rows": [
            {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper", "label": "DOWNSTREAM TAPER"},
            {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
             "spacingZone": "gapEndRoadWork"},
        ],
    }
    return [al1, al2]


def dims_merging(two_w20=True, shoulder_overlay=True):
    d = [
        {"zone": "gapA", "label": "A", "reference": None},
        {"zone": "gapB", "label": "B", "reference": None},
        {"zone": "gapC", "label": "C", "reference": None},
        {"zone": "laneTaper", "label": "MERGING TAPER L", "reference": None},
        {"zone": "bufferSpace", "label": "BUFFER SPACE", "reference": None},
        {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": None},
        {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER", "reference": None},
    ]
    if two_w20:
        d.insert(0, {"zone": "gapD", "label": "D", "reference": None})
    if shoulder_overlay:
        d.append({"zone": "shoulderTaper", "label": "SHOULDER TAPER", "reference": None})
    return d


def base_rules(sheet_n: str):
    return [
        {"id": "no-occupancy-buffer-rollahead", "severity": "error", "source": "Notes",
         "assert": "No workers/equipment in bufferSpace or rollAheadDistance.",
         "commonFailure": "Hatching through roll ahead."},
        {"id": "sign-order", "severity": "error", "source": "Plan layout",
         "assert": "Upstream advance signs: W4-2R, then W20-5, then W20-1 (furthest).",
         "commonFailure": "Reversed advance warning order."},
        {"id": "fixed-gaps", "severity": "error", "source": "Plan callouts",
         "assert": "A/B/C(/D) gaps are fixed plan callouts, not an AW spacing table.",
         "commonFailure": "Applying 302-03 FREEWAY 1000/1500/2640 when sheet prints 1320."},
        {"id": "end-road-work-side", "severity": "error", "source": "G20-2 callout",
         "assert": "G20-2 downstream 80-400 ft past downstream taper.",
         "commonFailure": "G20-2 on upstream alignment."},
        {"id": "cone-spacing", "severity": "warning", "source": "Notes",
         "assert": "Channelizing spacing <= 40 ft in active work space (short-term) or per duration note.",
         "commonFailure": "Bare polyline without spacing."},
    ]


def build_318():
    d = load_draft(318)
    size_rows = d["tables"]["318-06"]["rows"]
    # Ensure WARNING FLAG has NON-FREEWAY for sync
    for r in size_rows:
        if r["signCode"] == "WARNING FLAG":
            r["NON-FREEWAY"] = "18x18"
            r["FREEWAY"] = "18x18"

    taper_id, roll_id, pv_id, sign_id = "318-01", "318-02", "318-04", "318-06"
    gaps = d["planGapsFt"]
    zones = corridor_merging_fixed(gaps, taper_id, roll_id, two_w20=True)

    signs = [
        sign_item("W20-1", "ROAD WORK 1 MILE / ½ MILE", "W20-01RM", "signD", flags=True,
                  note="Two plan placements; nearer may use W20-01RPM."),
        sign_item("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
        sign_item("W4-2R", "(merge symbol)", "W04-02R", "signA", flags=True),
        sign_item("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", size="48x24"),
        sign_item("NYW8-33", "LANE CLOSED", None, None, size="48x24", post=False,
                  mounted="protectiveVehicle"),
        sign_item("W4-1R", "(ramp merge)", "W04-01R", None, note="Entrance ramp merge — not in upstream walk."),
        sign_item("R1-2", "YIELD", "R01-02", None, note="Ramp yield."),
        sign_item("W3-2", "YIELD AHEAD", "W03-02", None, note="Ramp yield ahead."),
        sign_item("WARNING FLAG", None, None, None, size="18x18", post=False, mounted="W20-1, W4-2R"),
        sign_item("R2-1 OR NYR2-2/NYR2-6", "SPEED LIMIT", None, None, size="36x48",
                  note="Note 9 — halfway between 1st and 2nd advance warning signs."),
    ]
    # fix NYW8-33 key
    for s in signs:
        if s["signCode"] == "NYW8-33":
            s["signLibraryKey"] = None
            s["signLibraryNote"] = "Vehicle-mounted — not a post-mounted SignLibrary placement."
        if s["signCode"] == "WARNING FLAG":
            s["signLibraryKey"] = None
        if s["signCode"] == "R2-1 OR NYR2-2/NYR2-6":
            s["signLibraryKey"] = "R02-01"
        if s["signCode"] == "R1-2":
            s["signLibraryKey"] = "R01-02"
        if s["signCode"] == "W3-2":
            s["signLibraryKey"] = "W03-02"

    notes = [
        "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD.",
        "2. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN THE ACTIVE WORK SPACE.",
        "3. NO WORK ACTIVITY, EQUIPMENT, OR STORAGE OF VEHICLES, OR MATERIAL SHALL OCCUR WITHIN THE BUFFER SPACE AT ANY TIME.",
        "4. CHANNELIZING DEVICES SHALL BE PLACED TRANSVERSELY A MINIMUM OF EVERY 800' AS SHOWN WHEN A PAVED SHOULDER HAVING A WIDTH OF 8' OR GREATER IS CLOSED FOR A DISTANCE GREATER THAN 800'.",
        "5. MAINLINE MERGING TAPER WITH THE ARROW PANEL AT ITS STARTING POINT SHALL BE LOCATED SUFFICIENTLY IN ADVANCE SO THAT THE ARROW PANEL IS NOT VISIBLE TO DRIVERS ON THE ENTRANCE RAMP, AND SO THAT THE MAINLINE MERGING TRAFFIC FROM THE LANE CLOSURE HAS THE OPPORTUNITY TO STABILIZE BEFORE ENCOUNTERING THE VEHICULAR TRAFFIC MERGING FROM THE RAMP.",
        "6. IF THE RAMP CURVES SHARPLY TO THE RIGHT, WARNING SIGNS LOCATED IN ADVANCE OF THE ENTRANCE TERMINAL SHALL BE PLACED IN PAIRS (ONE ON EACH SIDE OF THE RAMP).",
        "7. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS/ENGINE OFF) OR PARK/NEUTRAL (AUTOMATIC TRANSMISSION) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
        "8. VEH # 3 IS ONLY NEEDED WHEN THE SHOULDER WIDTH IS >= 8'.",
        "9. A REGULATORY SPEED LIMIT SIGN IS REQUIRED HALFWAY BETWEEN THE 1ST AND 2ND ADVANCE WARNING SIGNS UNLESS A REGULATORY SPEED LIMIT SIGN IS ALREADY PRESENT BETWEEN THOSE ADVANCED WARNING SIGNS OR A REGULATORY SPEED LIMIT REDUCTION IS AUTHORIZED AND THOSE SIGNS HAVE BEEN INSTALLED. ONE R2-1 OR NYR2-2 THROUGH NYR2-6 SHALL BE PROVIDED AS APPROPRIATE DEPENDING ON THE LOCATION. SEE STANDARD SHEET 619-012 FOR SIGN FACE AND SIZE.",
    ]

    # Fix inputs usedBy for 318
    inp = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
         "usedBy": ["318-01", "318-02", "318-03"]},
        {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "usedBy": ["318-01"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH3, "default": ">= 8 ft",
         "usedBy": ["318-01"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                     "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": ["318-04"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": ["318-04"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": ["318-06"]},
    ]

    spec = {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(
            318,
            "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR ENTRANCE RAMP",
            "SHORT TERM OPERATION", 2, 0,
            "Family 5 reference. Corridor like 302 (MERGING+DOWNSTREAM+shoulder overlay) with "
            "fixed plan gaps 1000/1500/1320/1320 (two W20-1) — NO advance-warning spacing table. "
            "Speeds 45/50/55/65. Table 318-01 == 302-02 overlap. PVH+TMIA freeway-only. "
            "Ramp signs W4-1R/R1-2/W3-2. Advance-placement table 318-03 (NY2C-4) is advisory, not A/B/C.",
        ),
        "applicability": {
            "roadType": "Freeway",
            "roadway": "Freeway near entrance ramp",
            "closure": "Single lane closure near entrance ramp",
            "duration": "Short Term",
            "durationDefinition": "Daytime work that occupies a location for more than 1 hour within a single daylight period (Note 1).",
            "speedRangeMph": {"allowed": SPEEDS, "note": "Table 318-01 covers 45/50/55/65 only."},
            "laneWidthFt": [10, 11, 12],
            "shoulderWidthBands": SH3,
            "areaTypes": None,
            "areaTypeNote": "No AW spacing table; gaps are fixed plan callouts.",
        },
        "inputs": inp,
        "tableRoles": {
            "note": "Roles by CONTENT. 318-03 advance placement is NOT advanceWarningSpacing. "
                    "318-05 channelizing omitted from roles (matrix present on PDF).",
            "taperAndBuffer": "318-01",
            "rollAheadDistance": "318-02",
            "protectiveVehicle": "318-04",
            "signSizes": "318-06",
        },
        "tables": {
            "318-01": {
                "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph", "laneWidthFt", "shoulderWidthBand"],
                "note": "Identical to 302-02 on 45/50/55/65.",
                "rows": d["tables"]["318-01"]["rows"],
            },
            "318-02": {
                "title": "ROLL AHEAD DISTANCE",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph"],
                "rows": d["tables"]["318-02"]["rows"],
                "usageNote": "MIN/MAX range.",
            },
            "318-03": {
                "title": "ADVANCE PLACEMENT OF WARNING SIGN",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph"],
                "note": "NY2C-4 advisory distances — NOT the A/B/C plan gaps.",
                "rows": d["tables"]["318-03"]["rows"],
            },
            "318-04": {
                "title": "PROTECTIVE VEHICLE REQUIREMENTS",
                "confidence": "verbatim",
                "keyedBy": ["closureType", "exposureCondition"],
                "note": "FREEWAY column only; all PVH+TMIA.",
                "rows": d["tables"]["318-04"]["rows"],
                "legend": {
                    "PVH": "PROTECTIVE VEHICLE HEAVY (MINIMUM GROSS WEIGHT 22,000 LBS. OR GREATER)",
                    "TMIA": "TMIA REQUIRED",
                },
            },
            "318-06": {
                "title": "REQUIRED SIGN SIZES",
                "confidence": "verbatim",
                "keyedBy": ["signCode", "signSizeClass"],
                "rows": size_rows,
            },
        },
        "corridor": {
            "confidence": "drawing",
            "description": "Fixed gaps 1000/1500/1320/1320; MERGING+DOWNSTREAM; shoulder overlay in A; ramp branch with W4-1R/YIELD.",
            "zones": zones,
        },
        "orderTable": {
            "confidence": "drawing",
            "description": "Roll Ahead, Buffer, Merging Taper, W4-2R, W20-5, W20-1 x2. Downstream taper + G20-2.",
            "alignments": order_merging(two_w20=True),
        },
        "signs": {"confidence": "verbatim", "items": signs},
        "symbols": {
            "confidence": "drawing",
            "items": [
                {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True,
                 "stationAnchor": {"zone": "laneTaper", "end": "upstream"}},
                {"id": "protectiveVehicle", "sheetLabel": "VEH #1", "required": True,
                 "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"},
                 "carriesSign": "NYW8-33"},
                {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES",
                 "deviceSymbol": "CONE", "required": True,
                 "runs": [
                     {"id": "shoulderTaperRun", "zone": "shoulderTaper",
                      "deviceCountSource": {"table": "318-01", "column": "shoulderTaper.devices"}},
                     {"id": "laneTaperRun", "zone": "laneTaper",
                      "deviceCountSource": {"table": "318-01", "column": "laneTaper.devices"}},
                     {"id": "longitudinalRun", "zone": "bufferSpace..workArea", "deviceCountSource": None},
                     {"id": "downstreamRun", "zone": "downstreamTaper", "deviceCountSource": None},
                 ]},
                {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
            ],
        },
        "annotations": {
            "confidence": "drawing",
            "dimensions": dims_merging(True, True),
            "lateralDimensions": [],
            "leaderCallouts": [
                {"text": "SEE DETAIL 318A", "pointsAt": "downstream taper"},
                {"text": "VEH #3 (SEE NOTE 8)", "pointsAt": "shoulder PV when >=8 ft"},
            ],
            "notLabeled": [{"item": "WORK AREA length", "reason": "Hatched; no length dimension."}],
        },
        "details": {
            "318A": {
                "title": "DETAIL 318A",
                "confidence": "drawing",
                "purpose": "Transverse channelizing when paved shoulder >=8 ft closed >800 ft.",
            }
        },
        "notes": {"confidence": "verbatim", "printed": notes, "planCallouts": [
            "THIS SIGN SHALL BE LOCATED A MINIMUM DISTANCE OF 80 FT AND MAXIMUM OF 400 FT PAST THE END OF THE DOWNSTREAM TAPER.",
        ]},
        "rules": base_rules("318") + [
            {"id": "ramp-arrow-panel-visibility", "severity": "error", "source": "Note 5",
             "assert": "Arrow panel at mainline merging taper start must not be visible from entrance ramp.",
             "commonFailure": "Placing arrow panel where ramp traffic can see it."},
            {"id": "shoulder-taper-is-an-overlay", "severity": "error", "source": "Plan datums",
             "assert": "Shoulder taper lies inside gap A; consumes no station.",
             "commonFailure": "Sequential shoulder taper row pushing signs upstream."},
        ],
        "knownCodeDeviations": [
            {"where": "order table dual W20-1", "issue": "Both W20-1 rows resolve to W20-01RM; nearer should be W20-01RPM.",
             "specSection": "signs.items / orderTable"},
        ],
    }
    write("619-318", spec)


def clone_lane_sheet(n, title, op, draft_key_map, gaps, two_w20, notes_printed, extra_note,
                     sign_defs, pages=2, rot=0, duration="Short Term"):
    """Build a 318-like lane+3band sibling."""
    d = load_draft(n)
    taper_id = draft_key_map["taper"]
    roll_id = draft_key_map["roll"]
    pv_id = draft_key_map["pv"]
    sign_id = draft_key_map["sign"]

    size_rows = d["tables"][sign_id]["rows"]
    # Harden sizes + sync with sign_defs
    codes_needed = [s[0] for s in sign_defs]
    by_code = {r["signCode"]: r for r in size_rows}
    size_rows = []
    for code, fw, *rest in [(s[0], s[4] if len(s) > 4 else "48x48") for s in sign_defs]:
        pass
    # rebuild size rows from sign_defs
    size_rows = []
    for sd in sign_defs:
        code, legend, key, zone = sd[0], sd[1], sd[2], sd[3]
        fw = sd[4] if len(sd) > 4 else ("48x24" if code in ("G20-2", "NYW8-33") else
                                         "18x18" if code == "WARNING FLAG" else "48x48")
        nf = "18x18" if code == "WARNING FLAG" else ("36x18" if code == "G20-2" else None)
        if code == "NYW8-33":
            nf = "48x24"
        size_rows.append(size_row(code, fw, nf))

    taper_rows = d["tables"][taper_id]["rows"]
    roll_rows = d["tables"][roll_id]["rows"]
    pv_rows = d["tables"][pv_id]["rows"]

    bands = SH7 if d.get("taperShape") == "sevenBand" else SH3
    two = two_w20 and "D" in gaps

    zones = corridor_merging_fixed(gaps, taper_id, roll_id, two_w20=two,
                                   shoulder_overlay=(bands == SH3 or True))
    # For seven-band sheets still use shoulder overlay

    signs = []
    for sd in sign_defs:
        code, legend, key, zone = sd[0], sd[1], sd[2], sd[3]
        fw = sd[4] if len(sd) > 4 else "48x48"
        post = code not in ("NYW8-33", "WARNING FLAG")
        mounted = "protectiveVehicle" if code == "NYW8-33" else (
            "W20-1, W4-2R" if code == "WARNING FLAG" else None)
        flags = code in ("W20-1", "W4-2R")
        item = sign_item(code, legend, key, zone, size=fw, post=post, mounted=mounted, flags=flags)
        if key is None:
            item["signLibraryKey"] = None
            if code == "NYW8-33":
                item["signLibraryNote"] = "Vehicle-mounted."
        signs.append(item)

    pv_used = pv_id
    inp = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
         "usedBy": [taper_id, roll_id]},
        {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "usedBy": [taper_id]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": bands,
         "default": ">= 8 ft" if ">= 8 ft" in bands else "12 ft", "usedBy": [taper_id]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                     "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": [pv_used]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": [pv_used]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "FREEWAY", "usedBy": [sign_id]},
    ]

    spec = {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(n, title, op, pages, rot, extra_note),
        "applicability": {
            "roadType": "Freeway",
            "roadway": "Freeway ramp-adjacent",
            "closure": title.split("CONTROL ")[-1] if "CONTROL " in title else "Lane/shoulder",
            "duration": duration,
            "durationDefinition": notes_printed[0] if notes_printed else "",
            "speedRangeMph": {"allowed": SPEEDS},
            "laneWidthFt": [10, 11, 12],
            "shoulderWidthBands": bands,
            "areaTypes": None,
            "areaTypeNote": "Fixed plan gaps; no AW spacing table.",
        },
        "inputs": inp,
        "tableRoles": {
            "note": "Roles by CONTENT.",
            "taperAndBuffer": taper_id,
            "rollAheadDistance": roll_id,
            "protectiveVehicle": pv_id,
            "signSizes": sign_id,
        },
        "tables": {
            taper_id: {
                "title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph", "laneWidthFt", "shoulderWidthBand"],
                "rows": taper_rows,
                "note": "seven-band" if bands == SH7 else "lane+3 shoulder bands == 302-02 overlap",
            },
            roll_id: {
                "title": "ROLL AHEAD DISTANCE",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph"],
                "rows": roll_rows,
                "usageNote": "MIN/MAX range.",
            },
            pv_id: {
                "title": "PROTECTIVE VEHICLE REQUIREMENTS",
                "confidence": "verbatim",
                "keyedBy": ["closureType", "exposureCondition"],
                "rows": pv_rows,
                "legend": {"PVH": "PROTECTIVE VEHICLE HEAVY", "TMIA": "TMIA REQUIRED"},
            },
            sign_id: {
                "title": "REQUIRED SIGN SIZES",
                "confidence": "verbatim",
                "keyedBy": ["signCode", "signSizeClass"],
                "rows": size_rows,
            },
        },
        "corridor": {
            "confidence": "drawing",
            "description": f"Fixed gaps {gaps}; MERGING+DOWNSTREAM corridor.",
            "zones": zones,
        },
        "orderTable": {
            "confidence": "drawing",
            "description": "Upstream merging corridor walk.",
            "alignments": order_merging(two_w20=two),
        },
        "signs": {"confidence": "verbatim", "items": signs},
        "symbols": {
            "confidence": "drawing",
            "items": [
                {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True,
                 "stationAnchor": {"zone": "laneTaper", "end": "upstream"}},
                {"id": "protectiveVehicle", "sheetLabel": "VEH", "required": True,
                 "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"}},
                {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES",
                 "required": True, "runs": [
                     {"id": "laneTaperRun", "zone": "laneTaper", "deviceCountSource": None},
                     {"id": "downstreamRun", "zone": "downstreamTaper", "deviceCountSource": None},
                 ]},
                {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
            ],
        },
        "annotations": {
            "confidence": "drawing",
            "dimensions": dims_merging(two, True),
            "lateralDimensions": [],
            "leaderCallouts": [],
            "notLabeled": [],
        },
        "details": {},
        "notes": {"confidence": "verbatim", "printed": notes_printed, "planCallouts": []},
        "rules": base_rules(str(n)),
        "knownCodeDeviations": [],
    }
    write(f"619-{n}", spec)


def build_minimal_mobile(n, title, draft, sign_defs, notes, rot=0):
    """113 / 211 style — roll + one/two signs, no taper table."""
    d = load_draft(n)
    sign_id = f"{n}-02" if n != 113 else "113-02"
    pv_id = f"{n}-01"
    if n == 113:
        sign_id, pv_id = "113-02", "113-01"
    size_rows = []
    signs = []
    for sd in sign_defs:
        code, legend, key, zone, fw = sd[0], sd[1], sd[2], sd[3], sd[4]
        nf = "18x18" if code == "WARNING FLAG" else None
        size_rows.append(size_row(code, fw, nf))
        post = code not in ("WARNING FLAG", "NYW8-33")
        mounted = "W21-5AL" if code == "WARNING FLAG" and n == 113 else (
            "W20-1, W21-5aL" if code == "WARNING FLAG" else None)
        item = sign_item(code, legend, key, zone, size=fw, post=post, mounted=mounted)
        if key is None:
            item["signLibraryKey"] = None
        signs.append(item)

    # Fixed roll-ahead from notes
    roll_table = {
        "title": "ROLL AHEAD DISTANCE (from notes)",
        "confidence": "drawing",
        "keyedBy": ["preconstructionPostedSpeedMph"],
        "note": "No numbered roll-ahead table — Note gives 80 ft (<=45 mph ramp) / 160 ft (>45).",
        "rows": [
            {"speedBand": "> 45 MPH", "minMph": 46, "maxMph": None,
             "min": {"ft": 160, "skipLines": 4}, "max": {"ft": 160, "skipLines": 4}},
            {"speedBand": "<= 45 MPH", "minMph": None, "maxMph": 45,
             "min": {"ft": 80, "skipLines": 2}, "max": {"ft": 80, "skipLines": 2}},
        ],
    }

    # Primary advance sign
    primary = sign_defs[0]
    zones = [
        {"id": "signA", "order": 1, "kind": "sign", "signCode": primary[0],
         "sheetLegend": primary[1]},
        {"id": "gapA", "order": 2, "kind": "gap", "sheetLabel": "A",
         "lengthSource": {"fixedFt": d["planGapsFt"]["A"]}, "dimensioned": True},
        {"id": "protectiveVehicle", "order": 3, "kind": "symbol", "sheetLabel": "PVH",
         "lengthSource": None},
        {"id": "rollAheadDistance", "order": 4, "kind": "clearance",
         "sheetLabel": "ROLL AHEAD DISTANCE",
         "lengthSource": {"table": f"{n}-roll", "column": "range",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True},
        {"id": "workArea", "order": 5, "kind": "workArea", "sheetLabel": "WORK AREA",
         "lengthSource": None, "hatched": True, "dimensioned": False},
    ]
    # If second corridor sign (211 has W20-1 further upstream)
    has_w20 = any(s[0] == "W20-1" for s in sign_defs)
    order_rows = [
        {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
    ]
    if n == 211 and has_w20:
        # W21-5aL then W20-1
        zones = [
            {"id": "signB", "order": 1, "kind": "sign", "signCode": "W20-1",
             "sheetLegend": "ROAD WORK AHEAD"},
            {"id": "gapB", "order": 2, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"fixedFt": 1000}, "dimensioned": True},
            {"id": "signA", "order": 3, "kind": "sign", "signCode": "W21-5aL",
             "sheetLegend": "LEFT SHOULDER CLOSED"},
            {"id": "gapA", "order": 4, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 1000}, "dimensioned": False,
             "note": "Undimensioned gap from W21-5aL to PV — plan shows 1000' to W20-1."},
            {"id": "protectiveVehicle", "order": 5, "kind": "symbol", "sheetLabel": "PVH",
             "lengthSource": None},
            {"id": "rollAheadDistance", "order": 6, "kind": "clearance",
             "sheetLabel": "ROLL AHEAD DISTANCE",
             "lengthSource": {"table": "211-roll", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 7, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
        ]
        order_rows = [
            {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
            {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W21-5aL", "spacingZone": "gapA"},
            {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W20-1", "spacingZone": "gapB"},
        ]
        dim_zones = [
            {"zone": "gapB", "label": "1000'", "reference": None},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": None},
        ]
    else:
        order_rows.append(
            {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": primary[0], "spacingZone": "gapA"}
        )
        dim_zones = [
            {"zone": "gapA", "label": "1000'", "reference": None},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": None},
        ]

    roll_key = f"{n}-roll"
    spec = {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(n, title,
                            "MOBILE OPERATION" if n == 113 else "SHORT DURATION OPERATION",
                            1, rot, d["findings"][0]),
        "applicability": {
            "roadType": "Freeway",
            "roadway": "Freeway exit ramp",
            "closure": "Left shoulder on exit ramp",
            "duration": "Mobile" if n == 113 else "Short Duration",
            "durationDefinition": notes[0] if notes else "",
            "speedRangeMph": {"allowed": [45, 50, 55, 65],
                              "note": "Roll-ahead note uses 45 mph ramp-speed threshold."},
            "laneWidthFt": None,
            "shoulderWidthBands": SH3,
            "areaTypes": None,
        },
        "inputs": [
            {"id": "preconstructionPostedSpeedMph", "type": "integer",
             "allowed": [45, 50, 55, 65], "usedBy": [roll_key]},
            {"id": "shoulderWidthBand", "type": "enum", "allowed": SH3, "default": ">= 8 ft",
             "usedBy": []},
            {"id": "exposureCondition", "type": "enum",
             "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                         "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": [pv_id]},
            {"id": "closureType", "type": "enum",
             "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
             "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": [pv_id]},
            {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
             "default": "FREEWAY", "usedBy": [sign_id]},
        ],
        "tableRoles": {
            "note": "No taperAndBuffer / advanceWarningSpacing.",
            "protectiveVehicle": pv_id,
            "rollAheadDistance": roll_key,
            "signSizes": sign_id,
        },
        "tables": {
            pv_id: {
                "title": "PROTECTIVE VEHICLE REQUIREMENTS",
                "confidence": "verbatim",
                "keyedBy": ["closureType", "exposureCondition"],
                "rows": d["tables"][pv_id]["rows"],
            },
            roll_key: roll_table,
            sign_id: {
                "title": "REQUIRED SIGN SIZES",
                "confidence": "verbatim",
                "keyedBy": ["signCode", "signSizeClass"],
                "rows": size_rows,
            },
        },
        "corridor": {"confidence": "drawing", "description": "Minimal mobile/short-duration ramp shoulder.",
                     "zones": zones},
        "orderTable": {
            "confidence": "drawing",
            "description": "Roll ahead + advance sign(s).",
            "alignments": [{
                "alignIdx": 1, "name": "Upstream",
                "station0": "Protective vehicle",
                "walkDirection": "Upstream, against traffic",
                "rows": order_rows,
                "excludedRows": [
                    {"label": "BUFFER SPACE", "reason": "No buffer on this sheet."},
                    {"label": "MERGING TAPER", "reason": "No taper on this sheet."},
                    {"label": "SHOULDER TAPER", "reason": "No taper on this sheet."},
                ],
            }],
        },
        "signs": {"confidence": "verbatim", "items": signs},
        "symbols": {
            "confidence": "drawing",
            "items": [
                {"id": "protectiveVehicle", "sheetLabel": "PVH+TMIA", "required": True,
                 "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"}},
                {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
            ],
        },
        "annotations": {
            "confidence": "drawing",
            "dimensions": dim_zones,
            "lateralDimensions": [],
            "leaderCallouts": [],
            "notLabeled": [],
        },
        "details": {},
        "notes": {"confidence": "verbatim", "printed": notes, "planCallouts": []},
        "rules": [
            {"id": "roll-ahead-empty", "severity": "error", "source": "Notes",
             "assert": "No workers/equipment in roll ahead distance.",
             "commonFailure": "Occupying roll ahead."},
            {"id": "fixed-gap", "severity": "error", "source": "Plan",
             "assert": "Advance gap is fixed plan callout (1000').",
             "commonFailure": "Using freeway AW table."},
        ],
        "knownCodeDeviations": [],
    }
    write(f"619-{n}", spec)


def main():
    build_318()

    short_notes = [
        "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD.",
        "2. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN THE ACTIVE WORK SPACE.",
        "3. NO WORK ACTIVITY, EQUIPMENT, OR STORAGE OF VEHICLES, OR MATERIAL SHALL OCCUR WITHIN THE BUFFER SPACE AT ANY TIME.",
    ]
    inter_notes = [
        "1. INTERMEDIATE-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION MORE THAN ONE DAYLIGHT PERIOD UP TO 3 DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR.",
        "2. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE ACTIVE WORK SPACE.",
        "3. NO WORK ACTIVITY, EQUIPMENT, OR STORAGE OF VEHICLES, OR MATERIAL SHALL OCCUR WITHIN THE BUFFER SPACE AT ANY TIME.",
    ]
    long_notes = [
        "1. LONG-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION FOR MORE THAN 3 DAYS.",
        "2. CHANNELIZING DEVICE SPACING PER CHANNELIZING APPLICATION TABLE.",
        "3. NO WORK ACTIVITY, EQUIPMENT, OR STORAGE OF VEHICLES, OR MATERIAL SHALL OCCUR WITHIN THE BUFFER SPACE AT ANY TIME.",
    ]

    # 319
    clone_lane_sheet(
        319,
        "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR EXIT RAMP",
        "SHORT TERM OPERATION",
        {"taper": "319-01", "roll": "319-02", "pv": "319-03", "sign": "319-sign"},
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, True,
        short_notes + ["4. RAMP AREA CHANNELIZING — see plan.", "5. SEE NOTES ON SHEET."],
        "Family 5 exit-ramp sibling of 318. Same taper as 302-02/318-01. Adds E5-1/E5-2 ramp signs.",
        [
            ("W20-1", "ROAD WORK 1 MILE / ½ MILE", "W20-01RM", "signD"),
            ("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
            ("W4-2R", "(merge symbol)", "W04-02R", "signA"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("NYW8-33", "LANE CLOSED", None, None, "48x24"),
            ("W5-4", "RAMP NARROWS", "W05-04", None),
            ("E5-1", "EXIT", None, None, "72x60"),
            ("E5-2", "EXIT", "E05-02", None, "48x36"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
    )

    # 418 intermediate exit
    clone_lane_sheet(
        418,
        "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR EXIT RAMP",
        "INTERMEDIATE TERM OPERATION",
        {"taper": "418-01", "roll": "418-02", "pv": "418-03", "sign": "418-sign"},
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, True,
        inter_notes,
        "Family 5 intermediate exit-ramp. Lane+3band like 319; 20' device spacing typical of intermediate.",
        [
            ("W20-1", "ROAD WORK 1 MILE / ½ MILE", "W20-01RM", "signD"),
            ("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
            ("W4-2R", "(merge symbol)", "W04-02R", "signA"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("NYW8-33", "LANE CLOSED", None, None, "48x24"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        duration="Intermediate",
    )

    # 518 long-term exit — 3 gaps (1000/1500/2640), one W20-1
    clone_lane_sheet(
        518,
        "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR EXIT RAMP",
        "LONG TERM OPERATION",
        {"taper": "518-01", "roll": "518-02", "pv": "518-03", "sign": "518-sign"},
        {"A": 1000, "B": 1500, "C": 2640}, False,
        long_notes,
        "Family 5 long-term exit-ramp. Lane+3band; gaps 1000/1500/2640 (standard freeway C).",
        [
            ("W20-1", "ROAD WORK 1 MILE", "W20-01RM", "signC"),
            ("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
            ("W4-2R", "(merge symbol)", "W04-02R", "signA"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("NYW8-33", "LANE CLOSED", None, None, "48x24"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        duration="Long Term",
    )

    # 417 intermediate entrance — 7-band
    clone_lane_sheet(
        417,
        "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR ENTRANCE RAMP",
        "INTERMEDIATE TERM OPERATION",
        {"taper": "417-01", "roll": "417-02", "pv": "417-pv", "sign": "417-sign"},
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, True,
        inter_notes,
        "Family 5 intermediate entrance. 7-band shoulder grid; laneTaper aliases 10/11/12 ft cols "
        "(plan MERGING L refs same table). Gaps like 318.",
        [
            ("W20-1", "ROAD WORK 1 MILE / ½ MILE", "W20-01RM", "signD"),
            ("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
            ("W4-2R", "(merge symbol)", "W04-02R", "signA"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("NYW8-33", "LANE CLOSED", None, None, "48x24"),
            ("W4-1R", "(ramp merge)", "W04-01R", None),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        duration="Intermediate",
    )

    # 517 long-term entrance — 7-band, 3 gaps
    clone_lane_sheet(
        517,
        "WORK ZONE TRAFFIC CONTROL FREEWAY SINGLE LANE CLOSURE NEAR ENTRANCE RAMP",
        "LONG TERM OPERATION",
        {"taper": "517-01", "roll": "517-02", "pv": "517-pv", "sign": "517-sign"},
        {"A": 1000, "B": 1500, "C": 2640}, False,
        long_notes,
        "Family 5 long-term entrance. 7-band; gaps 1000/1500/2640.",
        [
            ("W20-1", "ROAD WORK 1 MILE", "W20-01RM", "signC"),
            ("W20-5", "RIGHT LANE CLOSED ½ MILE", "W20-05RM", "signB"),
            ("W4-2R", "(merge symbol)", "W04-02R", "signA"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("NYW8-33", "LANE CLOSED", None, None, "48x24"),
            ("W4-1R", "(ramp merge)", "W04-01R", None),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        duration="Long Term",
    )

    # 416 partial exit intermediate — 7-band, shoulder-style signs W21-5
    # Still uses laneTaper alias for order MERGING; may be better as SHOULDER — use merging for gate consistency with draft
    clone_lane_sheet(
        416,
        "WORK ZONE TRAFFIC CONTROL FREEWAY PARTIAL EXIT RAMP CLOSURE",
        "INTERMEDIATE TERM OPERATION",
        {"taper": "416-01", "roll": "416-02", "pv": "416-pv", "sign": "416-sign"},
        {"A": 1000, "B": 500}, False,
        inter_notes,
        "Family 5 intermediate partial exit ramp. 7-band shoulder; gaps 1000/500; W21-5aR (not W20-5).",
        [
            ("W20-1", "ROAD WORK AHEAD", "W20-01RA", "signC"),
            ("W21-5aR", "RIGHT SHOULDER CLOSED", "W21-05aR", "signB"),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        duration="Intermediate",
    )

    # 316 partial exit short-term — like 416 but short-term + W21-5
    clone_lane_sheet(
        316,
        "WORK ZONE TRAFFIC CONTROL FREEWAY PARTIAL EXIT RAMP CLOSURE",
        "SHORT TERM OPERATION",
        {"taper": "316-01", "roll": "316-02", "pv": "316-03", "sign": "316-05"},
        {"A": 1000, "B": 1500}, False,
        short_notes,
        "Family 5 short-term partial exit ramp. Rotation=270. W21-5/W20-1/G20-2; gaps 1000/1500. "
        "Taper cells from 302-02 overlap (verified where text layer readable).",
        [
            ("W20-1", "ROAD WORK AHEAD", "W20-01RA", "signC"),
            ("W21-5", "SHOULDER WORK", "W21-05", "signB"),
            ("W5-4", "RAMP NARROWS", "W05-04", None),
            ("G20-2", "END ROAD WORK", "G20-02", "signEndRoadWork", "48x24"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        rot=270,
    )

    # Fix 416/316 order — they don't have W4-2R. Need custom order.
    # Rebuild 416 and 316 with shoulder-style corridor after.

    build_minimal_mobile(
        113,
        "WORK ZONE TRAFFIC CONTROL FREEWAY LEFT SHOULDER CLOSURE ON EXIT RAMP",
        load_draft(113),
        [
            ("W21-5AL", "LEFT SHOULDER CLOSED", "W21-05aL", "signA", "48x48"),
            ("W5-4", "RAMP NARROWS", "W05-04", None, "48x48"),
        ],
        [
            "1. MOBILE IS WORK THAT MOVES INTERMITTENTLY OR CONTINUOUSLY.",
            "2. IF DURATION AT A LOCATION EXCEEDS 15 MINUTES, RECONFIGURE TO 619-211.",
            "3. THERE SHALL BE NO WORKERS, EQUIPMENT, OR OTHER VEHICLES IN THE ROLL AHEAD DISTANCE.",
            "4. THE 80 FT ROLL AHEAD DISTANCE IS BASED ON RAMP SPEEDS OF 45 MPH OR LESS. IF SPEEDS ARE GREATER THAN 45 MPH INCREASE DISTANCE TO 160 FT.",
            "5. VEHICLE SHALL STAY AS FAR RIGHT ON THE SHOULDER AS POSSIBLE.",
        ],
        rot=270,
    )

    build_minimal_mobile(
        211,
        "WORK ZONE TRAFFIC CONTROL FREEWAY LEFT SHOULDER CLOSURE ON EXIT RAMP",
        load_draft(211),
        [
            ("W21-5aL", "LEFT SHOULDER CLOSED", "W21-05aL", "signA", "48x48"),
            ("W20-1", "ROAD WORK AHEAD", "W20-01RA", "signB", "48x48"),
            ("W5-4", "RAMP NARROWS", "W05-04", None, "48x48"),
            ("W13-4P", "NEXT RAMP XX MILES", "W13-04P", None, "36x36"),
            ("WARNING FLAG", None, None, None, "18x18"),
        ],
        [
            "1. SHORT DURATION IS WORK THAT OCCUPIES A LOCATION FOR UP TO 1 HOUR.",
            "2. THE OPERATOR(S) SHALL REMAIN IN THE PROTECTIVE VEHICLE(S).",
            "3. THERE SHALL BE NO WORKERS, EQUIPMENT, OR OTHER VEHICLES IN THE ROLL AHEAD DISTANCE.",
            "4. THE 80 FT ROLL AHEAD DISTANCE IS BASED ON RAMP SPEEDS OF 45 MPH OR LESS. IF SPEEDS ARE GREATER THAN 45 MPH INCREASE DISTANCE TO 160 FT.",
            "5. VEH #1 SHALL STAY AS FAR RIGHT ON THE SHOULDER AS POSSIBLE.",
            "6. TRUCK OFF-TRACKING SHOULD BE CONSIDERED WHEN DETERMINING WHETHER THE MINIMAL LANE WIDTH OF 10' IS ADEQUATE.",
        ],
    )

    # Patch 416/316 corridors for W21-based (no W4-2R) — rewrite order zones
    patch_partial_exit(416, "416-01", "416-02", {"A": 1000, "B": 500}, "W21-5aR", "W21-05aR")
    patch_partial_exit(316, "316-01", "316-02", {"A": 1000, "B": 1500}, "W21-5", "W21-05")

    print("ALL Family 5 specs built")


def patch_partial_exit(n, taper_id, roll_id, gaps, sign_b, key_b):
    """Replace corridor/order for partial-exit sheets that use W21 not W4-2R."""
    path = SPEC / f"619-{n}.json"
    spec = json.loads(path.read_text(encoding="utf-8"))
    zones = [
        {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1",
         "sheetLegend": "ROAD WORK AHEAD"},
        {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
         "lengthSource": {"fixedFt": gaps.get("B", 1500)}, "dimensioned": True,
         "spans": "W20-1 to W21"},
        {"id": "signB", "order": 3, "kind": "sign", "signCode": sign_b,
         "sheetLegend": "SHOULDER CLOSED / SHOULDER WORK"},
        {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
         "lengthSource": {"fixedFt": gaps["A"]}, "dimensioned": True,
         "spans": f"{sign_b} to shoulder taper"},
        {"id": "shoulderTaper", "order": 5, "kind": "taper", "sheetLabel": "SHOULDER TAPER",
         "sheetReference": f"(SEE TABLE {taper_id})",
         "lengthSource": {"table": taper_id, "column": "shoulderTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
         "dimensioned": True},
        {"id": "bufferSpace", "order": 6, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
         "sheetReference": f"(SEE TABLE {taper_id})",
         "lengthSource": {"table": taper_id, "column": "longitudinalBufferSpace",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True},
        {"id": "protectiveVehicle", "order": 7, "kind": "symbol", "sheetLabel": "PVH",
         "lengthSource": None},
        {"id": "rollAheadDistance", "order": 8, "kind": "clearance",
         "sheetLabel": "ROLL AHEAD DISTANCE",
         "lengthSource": {"table": roll_id, "column": "range",
                          "lookupBy": ["preconstructionPostedSpeedMph"]},
         "dimensioned": True, "mustBeEmpty": True},
        {"id": "workArea", "order": 9, "kind": "workArea", "sheetLabel": "WORK AREA",
         "lengthSource": None, "hatched": True, "dimensioned": False},
        {"id": "downstreamTaper", "order": 10, "kind": "taper", "sheetLabel": "DOWNSTREAM TAPER",
         "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}},
         "sheetText": "50'-100'", "dimensioned": True},
        {"id": "gapEndRoadWork", "order": 11, "kind": "gap", "sheetLabel": None,
         "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False},
        {"id": "signEndRoadWork", "order": 12, "kind": "sign", "signCode": "G20-2",
         "sheetLegend": "END ROAD WORK"},
    ]
    spec["corridor"]["zones"] = zones
    spec["orderTable"]["alignments"] = [
        {
            "alignIdx": 1, "name": "Upstream",
            "station0": "Upstream edge of the WORK AREA",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance",
                 "label": "ROLL AHEAD DISTANCE"},
                {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                {"rowNum": 3, "type": "Non-Sign", "zone": "shoulderTaper", "label": "SHOULDER TAPER"},
                {"rowNum": 4, "type": "Sign", "zone": "signB", "signCode": sign_b,
                 "spacingZone": "gapB"},
                {"rowNum": 5, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                 "spacingZone": "gapC"},
            ],
            "excludedRows": [
                {"label": "MERGING TAPER", "reason": "Partial exit ramp — shoulder taper only."},
                {"label": "Vehicle Space", "reason": "Not on this sheet."},
            ],
        },
        {
            "alignIdx": 2, "name": "Downstream",
            "station0": "Downstream edge of the WORK AREA",
            "walkDirection": "Downstream, with traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper",
                 "label": "DOWNSTREAM TAPER"},
                {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                 "spacingZone": "gapEndRoadWork"},
            ],
        },
    ]
    # Fix sign corridor zones
    for s in spec["signs"]["items"]:
        if s["signCode"] == "W20-1":
            s["corridorZone"] = "signC"
            s["signLibraryKey"] = "W20-01RA"
        if s["signCode"] == sign_b:
            s["corridorZone"] = "signB"
            s["signLibraryKey"] = key_b
        if s["signCode"] == "G20-2":
            s["corridorZone"] = "signEndRoadWork"
    # Remove W4-2R from signs/sizes if present
    spec["signs"]["items"] = [s for s in spec["signs"]["items"] if s["signCode"] != "W4-2R"]
    sid = spec["tableRoles"]["signSizes"]
    spec["tables"][sid]["rows"] = [
        r for r in spec["tables"][sid]["rows"] if r["signCode"] != "W4-2R"
    ]
    spec["annotations"]["dimensions"] = [
        {"zone": "gapC", "label": "C", "reference": None},
        {"zone": "gapB", "label": "B", "reference": None},
        {"zone": "shoulderTaper", "label": "SHOULDER TAPER", "reference": None},
        {"zone": "bufferSpace", "label": "BUFFER SPACE", "reference": None},
        {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": None},
        {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER", "reference": None},
    ]
    # symbols: no laneTaper anchor
    for sym in spec["symbols"]["items"]:
        if sym["id"] == "arrowPanel":
            sym["stationAnchor"] = {"zone": "shoulderTaper", "end": "upstream"}
        if sym["id"] == "channelizingDevices":
            sym["runs"] = [
                {"id": "shoulderTaperRun", "zone": "shoulderTaper", "deviceCountSource": None},
                {"id": "downstreamRun", "zone": "downstreamTaper", "deviceCountSource": None},
            ]
    write(f"619-{n}", spec)


if __name__ == "__main__":
    main()
