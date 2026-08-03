"""Family 1 builders part 2: 312, 325, 412, 423, 523 + main()."""
from __future__ import annotations

import copy
import json
import pathlib
import sys

# Reuse part1 definitions by importing after exec, or duplicate minimal helpers.
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "Bridge"))

# Import by running part1 module functions
import importlib.util
spec = importlib.util.spec_from_file_location("f1p1", ROOT / "Bridge/_build_family1_part1.py")
f1 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(f1)

SPEC_DIR = f1.SPEC_DIR
SRC = f1.SRC
SH_BANDS = f1.SH_BANDS
SPEEDS = f1.SPEEDS
write = f1.write
clone_table = f1.clone_table
base_from_311 = f1.base_from_311
size_row = f1.size_row
sign_item = f1.sign_item
channelizing_stub = f1.channelizing_stub
excluded_default = f1.excluded_default
T_PV_INT = f1.T_PV_INT
T_PV_LONG = f1.T_PV_LONG
T_AW_AB = f1.T_AW_AB
T_TAPER_NOSH = f1.taper_no_shoulder()
build_317_like = f1.build_317_like
remap_table_refs = f1.remap_table_refs
apply_merging_label = f1.apply_merging_label
ref311 = f1.ref311


def build_312():
    """TWLT short-term: A/B only, no shoulder taper cols, L + L/2, extra signs."""
    s = base_from_311(
        number="619-312",
        title="MULTI-LANE UNDIVIDED ROADWAY TWO-WAY LEFT TURN LANE CLOSURE",
        operation="SHORT TERM OPERATION",
        sourceUrl=f"{SRC}/619-312.pdf",
        localPdf="Bridge/captures/619-312.pdf",
        pdfPages=2,
        pageRotation=0,
        transcribedBy="Cursor (Family 1 TWLT sibling of 619-311)",
        provenanceNote=(
            "Two-way left turn lane closure. Table roles by content: 01=AW(A/B), 02=roll, "
            "03=PV, 04=taper(buffer+lane only, no shoulder), 05=sizes. Plan dimensions L and "
            "L/2 (shifting) from laneTaper; no shoulder-taper columns. Extra signs W9-3, R4-7."
        ),
    )
    s["applicability"].update({
        "closure": "Two-way left turn lane closure",
        "roadway": "Multilane undivided with TWLT",
        "shoulderWidthBands": None,
        "shoulderWidthBandNote": "No shoulder-taper columns on Table 312-04.",
        "areaTypes": ["URBAN", "RURAL"],
        "areaTypeNote": "Table 312-01 prints A/B only (no C).",
    })
    roles = {
        "note": "Roles by CONTENT: 01=AW, 02=roll, 03=PV, 04=taper, 05=sizes. AW is A/B only.",
        "advanceWarningSpacing": "312-01",
        "rollAheadDistance": "312-02",
        "protectiveVehicle": "312-03",
        "taperAndBuffer": "312-04",
        "signSizes": "312-05",
    }
    s["tableRoles"] = roles
    aw = copy.deepcopy(T_AW_AB)
    taper = copy.deepcopy(T_TAPER_NOSH)
    s["tables"] = {
        "312-01": aw,
        "312-02": clone_table("311-04", "312-02"),
        "312-03": clone_table("311-01", "312-03"),
        "312-04": taper,
        "312-05": {
            "title": "REQUIRED SIGN SIZES",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                size_row("G20-2", "36x18", "48x24"),
                size_row("NYW8-33", "48x24", "48x24"),
                size_row("R4-7", "24x30", "36x48"),
                size_row("W4-2L", "36x36", "48x48"),
                size_row("W9-3", "36x36", "48x48"),
                size_row("W20-1", "36x36", "48x48"),
                size_row("W20-5", "36x36", "48x48"),
                size_row("WARNING FLAG", "18x18", "18x18"),
            ],
        },
    }
    s["inputs"] = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": SPEEDS,
         "usedBy": ["312-01", "312-02", "312-03", "312-04"]},
        {"id": "laneWidthFt", "type": "integer", "allowed": [10, 11, 12], "usedBy": ["312-04"]},
        {"id": "shoulderWidthBand", "type": "enum", "allowed": SH_BANDS, "default": ">= 8 ft",
         "usedBy": [], "note": "Not used for taper lookup on this sheet."},
        {"id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL"], "usedBy": ["312-01"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                     "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": ["312-03"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
         "usedBy": ["312-03"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["312-05"]},
    ]
    # Corridor: W20-1 —B— W20-5/W9-3 cluster —A— L/2 overlay — L — buffer — roll — work
    # Simplified walk matching order-table engine: roll, buffer, L, W4-2L@A, W20-5@B, W20-1
    # L/2 is overlay inside gap A (like shoulder taper on 311) OR sequential — plan shows
    # L/2 and L as separate dimensions; L/2 shares character of shifting taper upstream of L.
    s["corridor"] = {
        "confidence": "drawing",
        "description": (
            "TWLT: advance W20-1/W20-5/W9-3/W4-2L with A/B gaps; L/2 shifting taper then L "
            "merging taper from Table 312-04 lane columns; buffer; roll ahead; dual-direction plan."
        ),
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1",
             "sheetLegend": "ROAD WORK XX"},
            {"id": "gapB", "order": 2, "kind": "gap", "sheetLabel": "B",
             "sheetReference": "(SEE TABLE 312-01)",
             "lengthSource": {"table": "312-01", "column": "B",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-5",
             "sheetLegend": "LANE CLOSED YY"},
            {"id": "gapA", "order": 4, "kind": "gap", "sheetLabel": "A",
             "sheetReference": "(SEE TABLE 312-01)",
             "lengthSource": {"table": "312-01", "column": "A",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True, "containsOverlay": "shiftingTaper",
             "note": "L/2 shifting taper is drawn at the head of L; treated as overlay inside A."},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W4-2L",
             "sheetLegend": "(merge symbol left)"},
            {"id": "shiftingTaper", "order": 6, "kind": "taper", "sheetLabel": "L/2",
             "sheetReference": "(SEE TABLE 312-04)",
             "lengthSource": {"table": "312-04", "column": "laneTaper", "scale": 0.5,
                              "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
             "dimensioned": True, "consumesStation": False, "containedIn": "gapA",
             "stationAnchor": {"zone": "laneTaper", "end": "upstream"},
             "note": "Plan labels L/2 — half of lane taper L from Table 312-04."},
            {"id": "laneTaper", "order": 7, "kind": "taper", "sheetLabel": "L",
             "sheetReference": "(SEE TABLE 312-04)",
             "lengthSource": {"table": "312-04", "column": "laneTaper",
                              "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
             "dimensioned": True},
            {"id": "bufferSpace", "order": 8, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "sheetReference": "(SEE TABLE 312-04)",
             "lengthSource": {"table": "312-04", "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True},
            {"id": "protectiveVehicle1", "order": 9, "kind": "symbol",
             "sheetLabel": "VEH #2", "lengthSource": None},
            {"id": "rollAheadDistance", "order": 10, "kind": "clearance",
             "sheetLabel": "ROLL AHEAD DISTANCE", "sheetReference": "(SEE TABLE 312-02)",
             "lengthSource": {"table": "312-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True},
            {"id": "workArea", "order": 11, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
            {"id": "downstreamTaper", "order": 12, "kind": "taper",
             "sheetLabel": "DOWNSTREAM TAPER",
             "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}},
             "sheetText": "50'-100'", "dimensioned": True},
            {"id": "gapEndRoadWork", "order": 13, "kind": "gap", "sheetLabel": None,
             "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False},
            {"id": "signEndRoadWork", "order": 14, "kind": "sign", "signCode": "G20-2",
             "sheetLegend": "END ROAD WORK"},
            {"id": "signW9_3", "order": 15, "kind": "sign", "signCode": "W9-3",
             "sheetLegend": "CENTER LANE CLOSED YY",
             "note": "Plan shows W9-3 on TWLT approach; not a separate station row in primary walk."},
            {"id": "signR4_7", "order": 16, "kind": "sign", "signCode": "R4-7",
             "sheetLegend": "KEEP RIGHT",
             "note": "Plan shows R4-7; not a primary upstream station row."},
        ],
    }
    s["orderTable"] = {
        "confidence": "drawing",
        "alignments": [
            {
                "alignIdx": 1, "name": "Upstream",
                "station0": "Upstream edge of the WORK AREA",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance",
                     "label": "ROLL AHEAD DISTANCE"},
                    {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                    {"rowNum": 3, "type": "Non-Sign", "zone": "laneTaper", "label": "L"},
                    {"rowNum": 4, "type": "Sign", "zone": "signA", "signCode": "W4-2L",
                     "spacingZone": "gapA"},
                    {"rowNum": 5, "type": "Sign", "zone": "signB", "signCode": "W20-5",
                     "spacingZone": "gapB"},
                    {"rowNum": 6, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                     "spacingZone": "gapB"},
                ],
                "overlayZones": [{
                    "zone": "shiftingTaper",
                    "anchor": {"zone": "laneTaper", "end": "upstream"},
                    "direction": "upstream",
                    "note": "L/2 overlay — does not consume station.",
                }],
                "excludedRows": excluded_default() + [
                    {"label": "SHOULDER TAPER", "reason": "No shoulder-taper column on 312-04."},
                    {"label": "LANE TAPER", "reason": "Sheet labels this L, not LANE TAPER."},
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
        ],
    }
    # Fix signC spacing — need distinct gapC or reuse; sheet has only A/B.
    # Use gapB for furthest sign spacing (B) and invent gapC as copy of B for W20-1.
    # Better: change row 6 spacingZone — sheet only has A/B; W20-1 uses B from AW table.
    # Keep as gapB for both mid and far is wrong. Add gapC with same lengthSource as B.
    s["corridor"]["zones"].insert(1, {
        "id": "gapC", "order": 1.5, "kind": "gap", "sheetLabel": "B",
        "sheetReference": "(SEE TABLE 312-01)",
        "lengthSource": {"table": "312-01", "column": "B",
                         "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
        "dimensioned": True,
        "note": "Sheet has no C column; furthest gap uses B value (A==B on this sheet).",
    })
    # Fix orders to be unique ascending integers
    for i, z in enumerate(s["corridor"]["zones"], 1):
        z["order"] = i
    s["orderTable"]["alignments"][0]["rows"][5]["spacingZone"] = "gapC"

    s["signs"] = {"confidence": "verbatim", "items": [
        sign_item("W20-1", sheetLegend="ROAD WORK XX",
                  legendSubstitution={"placeholder": "XX", "table": "312-01", "column": "XX"},
                  warningFlags=True, corridorZone="signC",
                  sizeNonFreeway="36x36", sizeFreeway="48x48", signLibraryBase="W20-01R"),
        sign_item("W20-5", sheetLegend="LANE CLOSED YY",
                  legendSubstitution={"placeholder": "YY", "table": "312-01", "column": "YY"},
                  corridorZone="signB",
                  sizeNonFreeway="36x36", sizeFreeway="48x48", signLibraryBase="W20-05L",
                  note="TWLT/left-side approach — SignLibrary W20-05L + legend suffix."),
        sign_item("W4-2L", sheetLegend="(merge symbol)", warningFlags=True,
                  corridorZone="signA", sizeNonFreeway="36x36", sizeFreeway="48x48",
                  signLibraryKey="W04-02L"),
        sign_item("W9-3", sheetLegend="CENTER LANE CLOSED YY", corridorZone="signW9_3",
                  sizeNonFreeway="36x36", sizeFreeway="48x48", signLibraryKey="W09-03",
                  legendSubstitution={"placeholder": "YY", "table": "312-01", "column": "YY"},
                  signLibraryBase="W09-03"),
        sign_item("R4-7", sheetLegend="KEEP RIGHT", shape="rectangle", corridorZone="signR4_7",
                  sizeNonFreeway="24x30", sizeFreeway="36x48", signLibraryKey="R04-07"),
        sign_item("G20-2", sheetLegend="END ROAD WORK", shape="rectangle",
                  corridorZone="signEndRoadWork",
                  sizeNonFreeway="36x18", sizeFreeway="48x24", signLibraryKey="G20-02"),
        sign_item("NYW8-33", sheetLegend="LANE CLOSED", shape="rectangle",
                  postMounted=False, mountedOn="protectiveVehicle1",
                  sizeNonFreeway="48x24", sizeFreeway="48x24"),
        sign_item("WARNING FLAG", shape="flag", postMounted=False,
                  mountedOn="W20-1, W4-2L", sizeNonFreeway="18x18", sizeFreeway="18x18"),
    ]}
    # Fix W9-3 — can't have both signLibraryKey and base; prefer base+legend
    for si in s["signs"]["items"]:
        if si["signCode"] == "W9-3":
            si.pop("signLibraryKey", None)
    s["symbols"] = {"confidence": "drawing", "items": [
        {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True,
         "stationAnchor": {"zone": "laneTaper", "end": "upstream"}},
        {"id": "protectiveVehicle1", "sheetLabel": "VEH #2", "required": "per Table 312-03",
         "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"},
         "carriesSign": "NYW8-33"},
        {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
         "longitudinalSpacing": {"maxFt": 40,
                                 "sheetText": "CHANNELIZING DEVICE SPACING SHALL NOT EXCEED 40'"}},
        {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
    ]}
    s["annotations"] = {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapC", "label": "B", "reference": "(SEE TABLE 312-01)"},
            {"zone": "gapB", "label": "B", "reference": "(SEE TABLE 312-01)"},
            {"zone": "gapA", "label": "A", "reference": "(SEE TABLE 312-01)"},
            {"zone": "shiftingTaper", "label": "L/2", "reference": "(SEE TABLE 312-04)"},
            {"zone": "laneTaper", "label": "L", "reference": "(SEE TABLE 312-04)"},
            {"zone": "bufferSpace", "label": "BUFFER SPACE", "reference": "(SEE TABLE 312-04)"},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE",
             "reference": "(SEE TABLE 312-02)"},
            {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER", "reference": None},
        ],
        "lateralDimensions": [],
        "leaderCallouts": [],
        "notLabeled": [{"item": "Vehicle Space", "reason": "Not on this sheet."}],
    }
    s["details"] = {}
    s["notes"] = {
        "confidence": "verbatim",
        "printed": [
            "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD. THIS SETUP MAY ALSO BE USED FOR A SHORT DURATION TWO WAY LEFT TURN LANE CLOSURE.",
            "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACING MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS. IF THERE IS A CONFLICT, MOVE THE SIGN UPSTREAM.",
            "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
            "4. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
            "5. ADJACENT LANE CLOSURES ARE RECOMMENDED WHEN THE PRECONSTRUCTION SPEED LIMIT IS 45 MPH OR HIGHER.",
            "6. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN THE ACTIVE WORK SPACE.",
        ],
        "planCallouts": [],
        "tableNotes": [],
    }
    s["rules"] = [
        {"id": "l2-is-overlay", "severity": "error", "source": "Plan L/2 dimension",
         "assert": "L/2 shifting taper is an overlay (consumesStation=false), not a sequential row.",
         "commonFailure": "Adding L/2 as a station that pushes advance signs upstream."},
        {"id": "no-shoulder-taper-row", "severity": "error", "source": "Table 312-04",
         "assert": "No SHOULDER TAPER sequential row — table has no shoulder columns.",
         "commonFailure": "Copying 311 shoulder overlay."},
        {"id": "aw-ab-only", "severity": "error", "source": "Table 312-01",
         "assert": "Advance spacing has A/B only (no C).",
         "commonFailure": "Looking up column C."},
    ]
    s["knownCodeDeviations"] = []
    s["knownExcerpts"] = {
        "from619-311": ["312-02==311-04", "312-03==311-01", "312-04 buffer+lane == 311-02"],
        "differsFrom311": ["no shoulder cols", "A/B only", "L/2", "W9-3/R4-7/W4-2L", "W20-5"],
    }
    write("312", s)


def build_325_family(num: str, duration: str, operation: str, pv, chan_title: str,
                     spacing: int, sizes: list, notes: list, provenance: str,
                     w20_5_variant: str = "W20-5"):
    """Double interior lane closure family: 325 / 423 / 523."""
    roles = {
        "note": (
            f"Roles by CONTENT for {num}: check titles — numbering differs across "
            "325/423/523. Assigned by content not suffix."
        ),
        "protectiveVehicle": f"{num}-01" if num == "325" else f"{num}-04",
        "rollAheadDistance": f"{num}-02" if num == "325" else f"{num}-03",
        "taperAndBuffer": f"{num}-03" if num == "325" else f"{num}-02",
        "advanceWarningSpacing": f"{num}-04" if num == "325" else f"{num}-01",
        "channelizingApplication": f"{num}-05",
        "signSizes": f"{num}-06",
    }
    # Normalize table ids for 423/523 where AW=01, taper=02, roll=03, PV=04
    if num != "325":
        tables = {
            f"{num}-01": clone_table("311-03", f"{num}-01"),
            f"{num}-02": clone_table("311-02", f"{num}-02"),
            f"{num}-03": clone_table("311-04", f"{num}-03"),
            f"{num}-04": copy.deepcopy(pv),
            f"{num}-05": channelizing_stub(chan_title, f"{spacing} ft"),
            f"{num}-06": {},
        }
        roles = {
            "note": "Roles by CONTENT: 01=AW, 02=taper, 03=roll, 04=PV, 05=channelizing, 06=sizes.",
            "advanceWarningSpacing": f"{num}-01",
            "taperAndBuffer": f"{num}-02",
            "rollAheadDistance": f"{num}-03",
            "protectiveVehicle": f"{num}-04",
            "channelizingApplication": f"{num}-05",
            "signSizes": f"{num}-06",
        }
    else:
        tables = {
            "325-01": clone_table("311-01", "325-01"),
            "325-02": clone_table("311-04", "325-02"),
            "325-03": clone_table("311-02", "325-03"),
            "325-04": clone_table("311-03", "325-04"),
            "325-05": channelizing_stub(chan_title, f"{spacing} ft"),
            "325-06": {},
        }
        roles = {
            "note": "Roles by CONTENT: 01=PV, 02=roll, 03=taper, 04=AW, 05=channelizing, 06=sizes.",
            "protectiveVehicle": "325-01",
            "rollAheadDistance": "325-02",
            "taperAndBuffer": "325-03",
            "advanceWarningSpacing": "325-04",
            "channelizingApplication": "325-05",
            "signSizes": "325-06",
        }

    s = build_317_like(
        num, duration=duration,
        title="MULTI-LANE TWO-WAY ROADWAY DOUBLE INTERIOR LANE CLOSURE",
        operation=operation,
        pv_table=pv, roles=roles, tables=tables, sizes=sizes,
        device_spacing=spacing, notes=notes, provenance=provenance,
        channelizing_title=chan_title, taper_label_merging=True,
    )
    s["applicability"]["closure"] = "Double interior lane closure"
    s["applicability"]["roadway"] = "Multilane two-way"
    # Dual-direction note
    s["corridor"]["description"] = (
        "Double interior lane closure — advance warning on both approaches. "
        "Primary upstream walk uses MERGING TAPER + buffer + roll (same skeleton as 311)."
    )
    # Sign code for mid advance
    mid = w20_5_variant
    for z in s["corridor"]["zones"]:
        if z["id"] == "signB":
            z["signCode"] = mid
    for al in s["orderTable"]["alignments"]:
        for r in al["rows"]:
            if r.get("zone") == "signB":
                r["signCode"] = mid
    # Ensure mid sign exists in signs.items
    codes = {i["signCode"] for i in s["signs"]["items"]}
    if mid not in codes:
        # find W20-5* in sizes
        for r in sizes:
            if r["signCode"].startswith("W20-5"):
                s["signs"]["items"].append(sign_item(
                    r["signCode"], sheetLegend="LANE CLOSED YY",
                    legendSubstitution={"placeholder": "YY",
                                        "table": roles["advanceWarningSpacing"],
                                        "column": "YY"},
                    corridorZone="signB",
                    sizeNonFreeway=r["NON-FREEWAY"], sizeFreeway=r["FREEWAY"],
                    signLibraryBase="W20-05L" if "L" in r["signCode"] else "W20-05R",
                ))
                break
    # Drop W20-5R if mid is different and not in size table
    size_codes = {r["signCode"] for r in sizes}
    s["signs"]["items"] = [i for i in s["signs"]["items"] if i["signCode"] in size_codes]
    # Fix W4-2 direction for left-oriented double interior sheets
    if any(r["signCode"] == "W4-2L" for r in sizes):
        for z in s["corridor"]["zones"]:
            if z.get("signCode") == "W4-2R":
                z["signCode"] = "W4-2L"
        for al in s["orderTable"]["alignments"]:
            for r in al["rows"]:
                if r.get("signCode") == "W4-2R":
                    r["signCode"] = "W4-2L"
        for i in s["signs"]["items"]:
            if i["signCode"] == "W4-2L":
                i["signLibraryKey"] = "W04-02L"
                i["corridorZone"] = "signA"
    return s


def build_325():
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
        "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "4. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC.",
        "5. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 40' IN THE ACTIVE WORK SPACE.",
    ]
    s = build_325_family(
        "325", "Short Term", "SHORT TERM OPERATION",
        clone_table("311-01", "325-01"),
        "CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES",
        40, sizes, notes,
        "Double interior short-term. Taper/AW/roll/PV == 311. Adds channelizing matrix. "
        "Dual-approach plan; primary walk MERGING TAPER. Roles: 01=PV,02=roll,03=taper,04=AW.",
        w20_5_variant="W20-5",
    )
    s["knownExcerpts"] = {
        "from619-311": ["325-01==311-01", "325-02==311-04", "325-03==311-02", "325-04==311-03"],
        "differsFrom311": ["double interior", "channelizing matrix", "W20-5", "2 pages"],
    }
    write("325", s)


def build_423():
    sizes = [
        size_row("G20-2", "36x18", "48x24"),
        size_row("NYW8-33", "48x24", "48x24"),
        size_row("W4-2L", "36x36", "48x48"),
        size_row("W20-1", "36x36", "48x48"),
        size_row("W20-5", "36x36", "48x48"),
        size_row("WARNING FLAG", "18x18", "18x18"),
    ]
    # NYR9-11 may be on plan but check size table - recon showed no NYR9-11 in size for 423
    notes = [
        "1. INTERMEDIATE-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION MORE THAN ONE DAYLIGHT PERIOD UP TO 3 CONSECUTIVE DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS.",
        "3. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE ACTIVE WORK SPACE.",
        "4. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "5. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE.",
        "6. THE NYR9-11 SIGN IS RECOMMENDED.",
    ]
    s = build_325_family(
        "423", "Intermediate Term", "INTERMEDIATE TERM OPERATION",
        T_PV_INT,
        "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES",
        20, sizes, notes,
        "Intermediate double interior. Taper/AW/roll == 311. PV = 011 INTERMEDIATE. "
        "20' spacing. W4-2L. Roles: 01=AW,02=taper,03=roll,04=PV.",
        w20_5_variant="W20-5",
    )
    s["knownExcerpts"] = {
        "from619-311": ["423-01==311-03", "423-02==311-02", "423-03==311-04"],
        "from619-011": ["423-04 == 011-01 INTERMEDIATE_TERM"],
        "differsFrom311": ["20' spacing", "W4-2L", "intermediate PV", "double interior"],
    }
    write("423", s)


def build_523():
    sizes = [
        size_row("G20-2", "36x18", "48x24"),
        size_row("NYW8-33", "48x24", "48x24"),
        size_row("NYR9-11", "24x42", "48x84"),
        size_row("W4-2L", "36x36", "48x48"),
        size_row("W20-1", "36x36", "48x48"),
        size_row("W20-5", "36x36", "48x48"),
        size_row("WARNING FLAG", "18x18", "18x18"),
    ]
    notes = [
        "1. LONG-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION FOR MORE THAN 3 CONSECUTIVE DAYS.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACINGS MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS.",
        "3. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE ACTIVE WORK SPACE.",
        "4. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "5. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE.",
        "6. THE NYR9-11 SIGN IS RECOMMENDED.",
    ]
    s = build_325_family(
        "523", "Long Term", "LONG TERM OPERATION",
        T_PV_LONG,
        "CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WORK ZONES",
        20, sizes, notes,
        "Long-term double interior. Taper/AW/roll == 311. PV = 011 LONG_TERM. "
        "NYR9-11 in size table. 20' spacing. Plan shows W20-5LA callout — size table prints W20-5.",
        w20_5_variant="W20-5",
    )
    s["knownExcerpts"] = {
        "from619-311": ["523-01==311-03", "523-02==311-02", "523-03==311-04"],
        "from619-011": ["523-04 == 011-01 LONG_TERM"],
        "differsFrom311": ["long-term PV", "NYR9-11", "20' spacing", "W4-2L", "double interior"],
    }
    write("523", s)


def build_412():
    """Intermediate TWLT — like 312 + intermediate extras. rotation=270."""
    # Start from 312 if exists, else build 312 first
    if not (SPEC_DIR / "619-312.json").is_file():
        build_312()
    s = json.loads((SPEC_DIR / "619-312.json").read_text(encoding="utf-8"))
    raw = json.dumps(s)
    raw = raw.replace("312-", "412-").replace("619-312", "619-412")
    s = json.loads(raw)
    s["sheet"].update({
        "number": "619-412",
        "operation": "INTERMEDIATE TERM OPERATION",
        "sourceUrl": f"{SRC}/619-412.pdf",
        "localPdf": "Bridge/captures/619-412.pdf",
        "pdfPages": 2,
        "pageRotation": 270,
        "transcribedBy": "Cursor (Family 1 intermediate TWLT sibling)",
        "provenanceNote": (
            "Intermediate TWLT. Diff vs 312: PV=011 INTERMEDIATE; adds NYR9-11; "
            "channelizing matrix 412-05; 20' spacing; pages rotation=270. "
            "Tables 412-01 AW A/B, 412-02 roll, 412-03 PV, 412-04 taper(no shoulder), "
            "412-06 sizes — roles by content. (Sheet also prints 412-06 / 011-03 refs.)"
        ),
    })
    s["applicability"]["duration"] = "Intermediate Term"
    s["applicability"]["durationDefinition"] = (
        "Stationary work occupying a location more than one daylight period up to "
        "3 consecutive days, or nighttime work lasting more than 1 hour."
    )
    # Insert channelizing + remap roles to match PDF titles on page 2
    # From recon: 412 has 01..06; use 312 mapping + channelizing as 05, sizes as 06
    # Current after replace: 412-01 AW, 02 roll, 03 PV, 04 taper, 05 sizes
    # Need: keep structure, rename sizes to 06, add 05 channelizing
    if "412-05" in s["tables"] and s["tables"]["412-05"].get("title", "").startswith("REQUIRED"):
        s["tables"]["412-06"] = s["tables"].pop("412-05")
    s["tables"]["412-03"] = copy.deepcopy(T_PV_INT)
    s["tables"]["412-05"] = channelizing_stub(
        "CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIONARY WORK ZONES", "20 ft")
    # Add NYR9-11 to sizes
    rows = s["tables"]["412-06"]["rows"]
    if not any(r["signCode"] == "NYR9-11" for r in rows):
        rows.insert(1, size_row("NYR9-11", "24x42", "48x84"))
    s["tableRoles"] = {
        "note": "Roles by CONTENT (rotation=270 PDF): 01=AW, 02=roll, 03=PV, 04=taper, 05=channelizing, 06=sizes.",
        "advanceWarningSpacing": "412-01",
        "rollAheadDistance": "412-02",
        "protectiveVehicle": "412-03",
        "taperAndBuffer": "412-04",
        "channelizingApplication": "412-05",
        "signSizes": "412-06",
    }
    for inp in s["inputs"]:
        if inp["id"] == "signSizeClass":
            inp["usedBy"] = ["412-06"]
        if inp["id"] == "exposureCondition" or inp["id"] == "closureType":
            inp["usedBy"] = ["412-03"]
    # Add NYR9-11 sign item
    codes = {i["signCode"] for i in s["signs"]["items"]}
    if "NYR9-11" not in codes:
        s["signs"]["items"].append(sign_item(
            "NYR9-11", sheetLegend="WORK ZONE SPEEDING", shape="rectangle",
            sizeNonFreeway="24x42", sizeFreeway="48x84", signLibraryKey="NYR9-11",
        ))
    for sym in s["symbols"]["items"]:
        if sym.get("id") == "channelizingDevices":
            sym["longitudinalSpacing"] = {"maxFt": 20, "sheetText": "not to exceed 20'"}
    s["notes"]["printed"] = [
        "1. INTERMEDIATE-TERM STATIONARY IS WORK THAT OCCUPIES A LOCATION MORE THAN ONE DAYLIGHT PERIOD UP TO 3 CONSECUTIVE DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR.",
        "2. IN URBAN CONDITIONS, ADVANCE WARNING SIGN SPACING MAY BE ADJUSTED IN ORDER TO ACCOMMODATE SIDE STREETS AND DRIVEWAYS.",
        "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR THE ROLL AHEAD DISTANCE.",
        "4. CHANNELIZING DEVICE SPACING (CENTER TO CENTER) SHALL NOT EXCEED 20' IN THE ACTIVE WORK SPACE.",
        "5. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE.",
        "6. THE NYR9-11 SIGN IS RECOMMENDED.",
    ]
    s["knownExcerpts"] = {
        "from619-312": ["corridor L/L2 TWLT skeleton", "AW A/B", "taper no shoulder"],
        "from619-011": ["412-03 == 011-01 INTERMEDIATE_TERM"],
        "differsFrom312": ["intermediate PV", "20' spacing", "NYR9-11", "rotation=270"],
    }
    write("412", s)


def main():
    print("=== Building Family 1 specs ===")
    f1.build_203()
    f1.build_202()
    f1.build_317()
    f1.build_414()
    build_312()
    build_412()
    build_325()
    build_423()
    build_523()
    print("=== done ===")


if __name__ == "__main__":
    main()
