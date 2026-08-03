"""Build Family 3 sibling specs (205/315/401/415/501) as diffs vs 619-301."""
from __future__ import annotations

import copy
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
ref = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))


def load_draft(num: str) -> dict:
    return json.loads((ROOT / f"Data/sheet-specs/_draft_619{num}_tables.json").read_text(encoding="utf-8"))


def sort_taper(tables: dict, tid: str) -> None:
    if tid in tables and "rows" in tables[tid] and tables[tid]["rows"] and "speedMph" in tables[tid]["rows"][0]:
        tables[tid]["rows"] = sorted(tables[tid]["rows"], key=lambda r: r["speedMph"])


def base_from_301(**sheet_fields) -> dict:
    s = copy.deepcopy(ref)
    s["sheet"].update(sheet_fields)
    return s


def write(num: str, s: dict) -> None:
    out = ROOT / f"Data/sheet-specs/619-{num}.json"
    out.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("wrote", out.name, "tables", list(s["tables"]), "zones", len(s["corridor"]["zones"]))


# ---------------------------------------------------------------------------
# 619-401 Intermediate shoulder (closest to 301 + intermediate extras)
# ---------------------------------------------------------------------------
def build_401():
    d = load_draft("401")
    sort_taper(d["tables"], "401-03")
    s = base_from_301(
        number="619-401",
        title="WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT SHOULDER CLOSURE",
        operation="INTERMEDIATE TERM OPERATION",
        approved="2026-04-29",
        sourceUrl="https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-401_E3.pdf",
        localPdf="Bridge/captures/619-401.pdf",
        pdfPages=2,
        pageRotation=0,
        transcribedBy="Cursor (Family 3 intermediate sibling of 619-301)",
        provenanceNote="Intermediate shoulder closure. 401-03 buffer+shoulder match 301-03; adds laneTaper + channelizing matrix; 20' device spacing; roll-ahead speed-keyed (not GVW); PVH+TMIA.",
    )
    s["applicability"]["duration"] = "Intermediate Term"
    s["applicability"]["durationDefinition"] = (
        "Stationary work occupying a location more than one daylight period up to 3 consecutive days, "
        "or nighttime work lasting more than 1 hour."
    )
    s["applicability"]["laneWidthFt"] = [10, 11, 12]
    s["applicability"]["laneWidthNote"] = "401-03 includes laneTaper columns (lateral shift) in addition to shoulder taper."
    s["tableRoles"] = d["tableRoles"]
    s["tables"] = d["tables"]
    # Remap corridor table refs 301-* -> 401-*
    for z in s["corridor"]["zones"]:
        ls = z.get("lengthSource") or {}
        if isinstance(ls, dict) and ls.get("table"):
            ls["table"] = ls["table"].replace("301-03", "401-03").replace("301-02", "401-02").replace("301-01", "401-01")
        if z.get("sheetReference"):
            z["sheetReference"] = z["sheetReference"].replace("301-03", "401-03").replace("301-02", "401-02")
    for d0 in s["annotations"]["dimensions"]:
        if d0.get("reference"):
            d0["reference"] = d0["reference"].replace("301-03", "401-03").replace("301-02", "401-02")
    for inp in s["inputs"]:
        inp["usedBy"] = [u.replace("301-", "401-") for u in inp.get("usedBy", [])]
    # Roll ahead is speed-keyed on 401 — drop GVW input emphasis
    s["inputs"] = [i for i in s["inputs"] if i["id"] != "protectiveVehicleGvwLbs"]
    # Channelizing 20'
    for sym in s["symbols"]["items"]:
        if sym.get("id") == "channelizingDevices":
            sym["longitudinalSpacing"] = {"maxFt": 20, "sheetText": "Note — intermediate 20' max (not 301's 40')."}
            for run in sym.get("runs", []):
                dcs = run.get("deviceCountSource")
                if dcs and dcs.get("table"):
                    dcs["table"] = dcs["table"].replace("301-03", "401-03")
    printed = [n for n in d["notes"]["printed"] if not str(n).startswith("N")]
    s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}
    s["rules"] = [
        {"id": "device-spacing-20ft", "severity": "error", "source": "Intermediate notes",
         "assert": "Channelizing spacing <= 20 ft in active work space.",
         "commonFailure": "Copying 40 ft from short-term 301."},
        {"id": "shoulder-taper-consumes-station", "severity": "error", "source": "Plan layout",
         "assert": "SHOULDER TAPER is sequential (Family 3 pattern).",
         "commonFailure": "Treating as Family 2 overlay."},
        {"id": "sign-order-shoulder", "severity": "error", "source": "Plan layout",
         "assert": "Advance signs are shoulder-closed family (W21-*), not W20-5R/W4-2R.",
         "commonFailure": "Using lane-closure signs."},
        {"id": "roll-ahead-by-speed", "severity": "warning", "source": "Table 401-02",
         "assert": "Roll ahead is speed-banded (unlike 301 GVW).",
         "commonFailure": "Passing protective_vehicle_gvw as if this were 301."},
    ]
    s["knownCodeDeviations"] = [
        {"id": "sign-codes-may-differ", "severity": "warning",
         "assert": "Draft lists W21-5bU/W21-5c variants — verify SignLibrary keys before live placement."},
    ]
    s["knownExcerpts"] = {"from619-301": ["401-03 buffer+shoulder == 301-03 on 45-65"],
                          "differsFrom301": ["20' spacing", "channelizing matrix", "speed roll-ahead", "laneTaper columns"]}
    # Keep 301 sign set if draft signs incomplete — patch codes from draft size table
    size_codes = {r["signCode"] for r in s["tables"]["401-05"]["rows"]}
    # Prefer W21-5aR/W21-5bR if present; else keep 301 items that exist in size table
    write("401", s)


# ---------------------------------------------------------------------------
# 619-501 Long-term shoulder + barrier
# ---------------------------------------------------------------------------
def build_501():
    d = load_draft("501")
    sort_taper(d["tables"], "501-01")
    s = base_from_301(
        number="619-501",
        title="WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT SHOULDER CLOSURE",
        operation="LONG TERM OPERATION",
        approved="2026-05-06",
        sourceUrl="https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-501_E3.pdf",
        localPdf="Bridge/captures/619-501.pdf",
        pdfPages=2,
        pageRotation=0,
        transcribedBy="Cursor (Family 3 long-term barrier sibling of 619-301)",
        provenanceNote="Long-term shoulder closure with temporary positive barrier. No PV/roll-ahead tables. 501-01 matches 301-03 buffer+shoulder. Flare rates 501-03.",
    )
    s["applicability"]["duration"] = "Long Term"
    s["applicability"]["durationDefinition"] = "Stationary work occupying a location more than 3 consecutive days."
    s["tableRoles"] = d["tableRoles"]
    s["tables"] = d["tables"]
    s["inputs"] = [i for i in s["inputs"] if i["id"] not in (
        "protectiveVehicleGvwLbs", "exposureCondition", "closureType")]
    for inp in s["inputs"]:
        inp["usedBy"] = [u.replace("301-03", "501-01").replace("301-02", "501-01").replace("301-01", "501-01")
                         for u in inp.get("usedBy", [])]
    # Corridor: drop PV and roll ahead; add barrier; remap taper table
    zones = []
    for z in s["corridor"]["zones"]:
        if z["id"] in ("protectiveVehicle", "rollAheadDistance"):
            continue
        ls = z.get("lengthSource") or {}
        if isinstance(ls, dict) and ls.get("table"):
            ls["table"] = "501-01" if "301-03" in ls["table"] or "301-02" in ls["table"] else ls["table"]
        if z.get("sheetReference"):
            z["sheetReference"] = z["sheetReference"].replace("301-03", "501-01").replace("301-02", "501-01")
        zones.append(z)
    # Insert barrier before work area
    wa_idx = next(i for i, z in enumerate(zones) if z["id"] == "workArea")
    zones.insert(wa_idx, {
        "id": "positiveBarrier", "order": zones[wa_idx]["order"],
        "kind": "barrier", "sheetLabel": "TEMPORARY POSITIVE BARRIER",
        "sheetReference": "(SEE TABLE 501-03)", "lengthSource": None, "dimensioned": False,
        "note": "Barrier not on shoulder taper (sheet note).",
    })
    for i, z in enumerate(zones):
        z["order"] = i + 1
    s["corridor"]["zones"] = zones
    s["corridor"]["description"] = (
        "Long-term shoulder closure: SHOULDER TAPER + advance signs; temporary positive barrier "
        "along work area (no PV/roll-ahead). Flare rates from Table 501-03."
    )
    s["orderTable"] = {
        "confidence": "drawing",
        "description": "No ROLL AHEAD / BUFFER rows — long-term barrier shoulder sheet.",
        "alignments": [
            {
                "alignIdx": 1, "name": "Upstream",
                "station0": "Upstream end of SHOULDER TAPER / barrier start",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "shoulderTaper", "label": "SHOULDER TAPER"},
                    {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W21-5aR", "spacingZone": "gapA"},
                    {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W21-5bR", "spacingZone": "gapB"},
                    {"rowNum": 4, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
                ],
                "excludedRows": [
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No roll-ahead on long-term barrier sheet."},
                    {"label": "BUFFER SPACE", "reason": "Buffer exists in 501-01 but barrier replaces PV/roll-ahead pattern; not a sequential row."},
                    {"label": "MERGING TAPER", "reason": "Shoulder closure."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            },
            {
                "alignIdx": 2, "name": "Downstream",
                "station0": "Downstream edge of WORK AREA",
                "walkDirection": "Downstream, with traffic",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper", "label": "DOWNSTREAM TAPER"},
                    {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                     "spacingZone": "gapEndRoadWork"},
                ],
            },
        ],
    }
    # Fix annotations — drop roll ahead dim; remap refs
    s["annotations"]["dimensions"] = [
        d0 for d0 in s["annotations"]["dimensions"]
        if d0["zone"] not in ("rollAheadDistance",)
    ]
    for d0 in s["annotations"]["dimensions"]:
        if d0.get("reference"):
            d0["reference"] = d0["reference"].replace("301-03", "501-01").replace("301-02", "501-01")
    printed = [n for n in d["notes"]["printed"] if not str(n).startswith("N")]
    s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}
    s["symbols"]["items"] = [x for x in s["symbols"]["items"] if x["id"] != "protectiveVehicle"]
    s["symbols"]["items"].append({
        "id": "positiveBarrier", "sheetLabel": "TEMPORARY POSITIVE BARRIER", "required": True,
        "stationAnchor": {"zone": "positiveBarrier", "end": "both"},
        "note": "Flare rates Table 501-03.",
    })
    s["rules"] = [
        {"id": "no-roll-ahead-no-pv", "severity": "error", "source": "Sheet structure",
         "assert": "Do not emit ROLL AHEAD or PV rows.",
         "commonFailure": "Cloning 301 upstream walk."},
        {"id": "barrier-not-on-taper", "severity": "error", "source": "Sheet notes",
         "assert": "Temporary positive barrier not placed along the shoulder taper.",
         "commonFailure": "Running barrier through the taper."},
        {"id": "shoulder-taper-consumes-station", "severity": "error", "source": "Plan",
         "assert": "SHOULDER TAPER is sequential.",
         "commonFailure": "Family 2 overlay pattern."},
    ]
    s["knownCodeDeviations"] = [
        {"id": "barrier-placement-unimplemented", "severity": "error",
         "assert": "No temporary positive barrier placer yet."},
    ]
    s["knownExcerpts"] = {"from619-301": ["501-01 buffer+shoulder == 301-03"],
                          "differsFrom301": ["No PV/roll-ahead", "barrier + flare rates", "long-term notes"]}
    write("501", s)


# ---------------------------------------------------------------------------
# 619-315 Short-term shoulder at ramp
# ---------------------------------------------------------------------------
def build_315():
    d = load_draft("315")
    sort_taper(d["tables"], "315-03")
    s = base_from_301(
        number="619-315",
        title="WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT SHOULDER CLOSURE AT RAMP APPROACH",
        operation="SHORT TERM OPERATION",
        approved="2023-05-05",
        sourceUrl="https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-315_E1.pdf",
        localPdf="Bridge/captures/619-315.pdf",
        pdfPages=2,
        pageRotation=270,
        transcribedBy="Cursor (Family 3 ramp sibling of 619-301)",
        provenanceNote="Short-term shoulder at ramp. 315-01/02 match 301; 315-03 adds lateralShiftTaper; plan spacing 2640'/1500'/1000' (FREEWAY-like C); W3-7a ramp plaque.",
    )
    s["applicability"]["closure"] = "Right shoulder closure at ramp approach"
    s["applicability"]["laneWidthFt"] = [10, 11, 12]
    s["tableRoles"] = d["tableRoles"]
    s["tables"] = d["tables"]
    for z in s["corridor"]["zones"]:
        ls = z.get("lengthSource") or {}
        if isinstance(ls, dict):
            if ls.get("table"):
                ls["table"] = (ls["table"].replace("301-03", "315-03")
                               .replace("301-02", "315-02").replace("301-01", "315-01"))
            if z["id"] == "gapC" and "fixedFt" in ls:
                ls["fixedFt"] = 2640  # plan uses FREEWAY-like C
                z["note"] = "Plan callout 2640' (FREEWAY C) — ramp sheet uses standard freeway C unlike 301's 1320'."
            if z["id"] == "gapB" and "fixedFt" in ls:
                ls["fixedFt"] = 1500
            if z["id"] == "gapA" and "fixedFt" in ls:
                ls["fixedFt"] = 1000
        if z.get("sheetReference"):
            z["sheetReference"] = (z["sheetReference"].replace("301-03", "315-03")
                                   .replace("301-02", "315-02"))
    for d0 in s["annotations"]["dimensions"]:
        if d0["zone"] == "gapC":
            d0["label"] = "2640'"
        if d0.get("reference"):
            d0["reference"] = d0["reference"].replace("301-03", "315-03").replace("301-02", "315-02")
    # Downstream taper may be 500' on plan — keep 50-100 unless draft says otherwise; note in knownCodeDeviations
    for inp in s["inputs"]:
        inp["usedBy"] = [u.replace("301-", "315-") for u in inp.get("usedBy", [])]
    printed = [n for n in d["notes"]["printed"] if not str(n).startswith("N")]
    s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}
    s["signs"]["items"].append({
        "signCode": "W3-7a", "sheetLegend": "RAMP distance plaque", "shape": "rectangle",
        "postMounted": True, "required": False, "signLibraryKey": None,
        "note": "Note 8 — ramp distance plaque; may be plan-only (not in size table).",
        "sizeNonFreeway": None, "sizeFreeway": None,
    })
    # Remove from size cross-check by not requiring size row — put in symbols if validator complains
    s["rules"] = [
        {"id": "ramp-spacing-freeway-C", "severity": "error", "source": "Plan callouts",
         "assert": "Gap C is 2640' (FREEWAY), not 301's 1320'.",
         "commonFailure": "Copying 301's fixed 1320' gap C."},
        {"id": "shoulder-taper-consumes-station", "severity": "error", "source": "Plan",
         "assert": "SHOULDER TAPER sequential.", "commonFailure": "Family 2 overlay."},
        {"id": "w3-7a-ramp", "severity": "warning", "source": "Note 8",
         "assert": "W3-7a ramp plaque when applicable.",
         "commonFailure": "Omitting ramp-specific plaque."},
    ]
    s["knownCodeDeviations"] = [
        {"id": "w3-7a-not-in-size-table", "severity": "warning",
         "assert": "W3-7a may not appear in 315-04 — tracked as optional sign."},
        {"id": "downstream-500", "severity": "warning",
         "assert": "Plan may show 500' downstream taper — confirm before relying on 50-100 default."},
    ]
    s["knownExcerpts"] = {"from619-301": ["315-01/02 == 301-01/02", "315-03 shoulder+buffer == 301-03"],
                          "differsFrom301": ["lateralShiftTaper", "2640' C", "W3-7a", "ramp approach"]}
    # Drop W3-7a from signs.items if it breaks size cross-check — move to symbols
    s["signs"]["items"] = [i for i in s["signs"]["items"] if i["signCode"] != "W3-7a"]
    s["symbols"]["items"].append({
        "id": "w37aRampPlaque", "sheetLabel": "W3-7a", "required": False,
        "note": "Ramp distance plaque (Note 8).",
    })
    write("315", s)


# ---------------------------------------------------------------------------
# 619-415 Intermediate shoulder at ramp
# ---------------------------------------------------------------------------
def build_415():
    d = load_draft("415")
    sort_taper(d["tables"], "415-01")
    s = base_from_301(
        number="619-415",
        title="WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT SHOULDER CLOSURE AT RAMP APPROACH",
        operation="INTERMEDIATE TERM OPERATION",
        approved="2026-04-30",
        sourceUrl="https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-415_E3.pdf",
        localPdf="Bridge/captures/619-415.pdf",
        pdfPages=2,
        pageRotation=0,
        transcribedBy="Cursor (Family 3 intermediate ramp sibling of 619-301)",
        provenanceNote="Intermediate shoulder at ramp. 415-01 buffer+shoulder match 301-03; speed-keyed roll-ahead; 20' spacing; NYR9-11; plan spacing includes 1320'.",
    )
    s["applicability"]["duration"] = "Intermediate Term"
    s["applicability"]["closure"] = "Right shoulder closure at ramp approach"
    s["applicability"]["laneWidthFt"] = [10, 11, 12]
    s["tableRoles"] = d["tableRoles"]
    s["tables"] = d["tables"]
    s["inputs"] = [i for i in s["inputs"] if i["id"] != "protectiveVehicleGvwLbs"]
    for inp in s["inputs"]:
        inp["usedBy"] = [u.replace("301-03", "415-01").replace("301-02", "415-02")
                         .replace("301-01", "415-03").replace("301-04", "415-05")
                         for u in inp.get("usedBy", [])]
    for z in s["corridor"]["zones"]:
        ls = z.get("lengthSource") or {}
        if isinstance(ls, dict) and ls.get("table"):
            ls["table"] = (ls["table"].replace("301-03", "415-01")
                           .replace("301-02", "415-02").replace("301-01", "415-03"))
        if z.get("sheetReference"):
            z["sheetReference"] = (z["sheetReference"].replace("301-03", "415-01")
                                   .replace("301-02", "415-02"))
        # Keep 301's 1320/1500/1000 — plan shows 1320'
    for d0 in s["annotations"]["dimensions"]:
        if d0.get("reference"):
            d0["reference"] = d0["reference"].replace("301-03", "415-01").replace("301-02", "415-02")
    for sym in s["symbols"]["items"]:
        if sym.get("id") == "channelizingDevices":
            sym["longitudinalSpacing"] = {"maxFt": 20, "sheetText": "Intermediate Note — 20' max."}
            for run in sym.get("runs", []):
                dcs = run.get("deviceCountSource")
                if dcs and dcs.get("table"):
                    dcs["table"] = "415-01"
    printed = [n for n in d["notes"]["printed"] if not str(n).startswith("N")]
    s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}
    s["rules"] = [
        {"id": "device-spacing-20ft", "severity": "error", "source": "Note 3",
         "assert": "Channelizing <= 20 ft.", "commonFailure": "Using 301's 40 ft."},
        {"id": "shoulder-taper-consumes-station", "severity": "error", "source": "Plan",
         "assert": "SHOULDER TAPER sequential.", "commonFailure": "Family 2 overlay."},
        {"id": "nyr9-11-recommended", "severity": "warning", "source": "Note 6",
         "assert": "NYR9-11 recommended 1000' before first advance warning.",
         "commonFailure": "Omitting NYR9-11 on intermediate ramp sheets."},
    ]
    s["knownExcerpts"] = {"from619-301": ["415-01 buffer+shoulder == 301-03"],
                          "differsFrom301": ["20' spacing", "speed roll-ahead", "ramp", "NYR9-11"]}
    write("415", s)


# ---------------------------------------------------------------------------
# 619-205 Short duration — mobile-ish, no taper table
# ---------------------------------------------------------------------------
def build_205():
    d = load_draft("205")
    s = base_from_301(
        number="619-205",
        title="WORK ZONE TRAFFIC CONTROL MULTILANE DIVIDED ROADWAY AND FREEWAY RIGHT SHOULDER CLOSURE",
        operation="SHORT DURATION OPERATION",
        approved="2021-12-06",
        sourceUrl="https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-205.pdf",
        localPdf="Bridge/captures/619-205.pdf",
        pdfPages=1,
        pageRotation=0,
        transcribedBy="Cursor (Family 3 short-duration sibling of 619-301)",
        provenanceNote="Short-duration shoulder. NO taper/buffer table. Operator remains in PV. Roll-ahead by speed. Plan spacing ~1000'. Signs W21-5 / W20-1. P+TMIA not PVH+TMIA.",
    )
    s["applicability"]["duration"] = "Short Duration"
    s["applicability"]["durationDefinition"] = (
        "Work that occupies a location up to 1 hour (short duration) — operator stays in the protective vehicle."
    )
    s["applicability"]["speedRangeMph"] = {
        "allowed": [45, 50, 55, 65],
        "note": "Roll-ahead table covers >=55 and 45-50 bands; 65 uses >=55 row.",
    }
    s["applicability"]["laneWidthFt"] = None
    s["tableRoles"] = d["tableRoles"]
    s["tables"] = d["tables"]
    # Ensure speedBands absent is OK; FREEWAY-only PV
    s["inputs"] = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": [45, 50, 55, 65],
         "usedBy": ["205-02"]},
        {"id": "exposureCondition", "type": "enum",
         "allowed": [
             "WORKERS ON FOOT OR WORK VEHICLE EXPOSED TO TRAFFIC",
             "OTHER HAZARDS EXPOSED (IE EQUIPMENT, MATERIALS, EXCAVATION)",
         ], "usedBy": ["205-01"]},
        {"id": "closureType", "type": "enum",
         "allowed": ["SHOULDER CLOSURE OR ENCROACHMENT"],
         "default": "SHOULDER CLOSURE OR ENCROACHMENT", "usedBy": ["205-01"]},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"], "default": "FREEWAY"},
    ]
    s["corridor"] = {
        "confidence": "drawing",
        "description": "Short-duration: ROLL AHEAD + occupied PV; advance W21-5 then W20-1 at ~1000'; no shoulder taper table/dimension.",
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK"},
            {"id": "gapB", "order": 2, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"fixedFt": 1000}, "dimensioned": True,
             "note": "Plan shows 1000' spacing callout."},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W21-5",
             "sheetLegend": "SHOULDER CLOSED / WORK"},
            {"id": "gapA", "order": 4, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 1000}, "dimensioned": True},
            {"id": "protectiveVehicle", "order": 5, "kind": "symbol", "sheetLabel": "WORK VEHICLE / PV",
             "lengthSource": None, "note": "Operator remains in vehicle (short duration)."},
            {"id": "rollAheadDistance", "order": 6, "kind": "clearance", "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": "(SEE TABLE 205-02)",
             "lengthSource": {"table": "205-02", "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 7, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
            {"id": "spotter", "order": 8, "kind": "symbol", "sheetLabel": "SPOTTER RECOMMENDED",
             "lengthSource": None, "required": False},
        ],
    }
    s["orderTable"] = {
        "confidence": "drawing",
        "description": "Short-duration: Roll Ahead + W21-5 + W20-1. No taper/buffer rows.",
        "alignments": [{
            "alignIdx": 1, "name": "Upstream",
            "station0": "Work vehicle / work area",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE"},
                {"rowNum": 2, "type": "Sign", "zone": "signB", "signCode": "W21-5", "spacingZone": "gapA"},
                {"rowNum": 3, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapB"},
            ],
            "excludedRows": [
                {"label": "BUFFER SPACE", "reason": "No buffer table on short-duration sheet."},
                {"label": "SHOULDER TAPER", "reason": "No taper table on short-duration sheet."},
                {"label": "MERGING TAPER", "reason": "Shoulder / short-duration."},
                {"label": "Vehicle Space", "reason": "Not on this sheet."},
            ],
        }],
    }
    s["signs"] = {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "shape": "diamond", "postMounted": True, "corridorZone": "signC",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RA"},
            {"signCode": "W21-5", "shape": "diamond", "postMounted": True, "corridorZone": "signB",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W21-05",
             "note": "Sheet prints W21-5 (generic); SignLibrary W21-05."},
            {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False, "mountedOn": "W20-1",
             "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
        ],
    }
    # Align size table codes
    for row in s["tables"]["205-03"]["rows"]:
        pass
    s["symbols"] = {
        "confidence": "drawing",
        "items": [
            {"id": "protectiveVehicle", "sheetLabel": "WORK VEHICLE", "required": True,
             "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"},
             "note": "Operator stays in vehicle; two-way radio."},
            {"id": "spotter", "sheetLabel": "SPOTTER RECOMMENDED", "required": False,
             "stationAnchor": {"zone": "workArea", "end": "downstream"}},
            {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
             "longitudinalSpacing": {"maxFt": 40, "sheetText": "Plan — 40' cone spacing."},
             "runs": [{"id": "workRun", "zone": "workArea", "deviceCountSource": None}]},
        ],
    }
    s["annotations"] = {
        "confidence": "drawing",
        "dimensions": [
            {"zone": "gapA", "label": "1000'", "reference": None},
            {"zone": "gapB", "label": "1000'", "reference": None},
            {"zone": "rollAheadDistance", "label": "ROLL AHEAD DISTANCE", "reference": "(SEE TABLE 205-02)"},
        ],
        "labels": [{"text": "SPOTTER RECOMMENDED", "zone": "spotter"}],
    }
    printed = [n for n in d["notes"]["printed"] if not str(n).startswith("N")]
    s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}
    s["rules"] = [
        {"id": "no-taper-table", "severity": "error", "source": "Sheet structure",
         "assert": "No SHOULDER TAPER / BUFFER sequential rows from a taper table.",
         "commonFailure": "Cloning 301's taper+buffer walk."},
        {"id": "operator-in-vehicle", "severity": "error", "source": "Short duration",
         "assert": "Protective/work vehicle remains occupied.",
         "commonFailure": "Treating as unoccupied PVH like 301."},
        {"id": "roll-ahead-by-speed", "severity": "error", "source": "Table 205-02",
         "assert": "Roll ahead from speed bands, not GVW.",
         "commonFailure": "Using 301 GVW lookup."},
    ]
    s["knownCodeDeviations"] = [
        {"id": "w21-5-generic", "severity": "warning",
         "assert": "Sheet uses W21-5; SignLibrary key W21-05 — confirm cell exists."},
    ]
    s["knownExcerpts"] = {"differsFrom301": ["No taper table", "P+TMIA", "speed roll-ahead", "short duration occupied PV"]}
    write("205", s)


if __name__ == "__main__":
    build_401()
    build_501()
    build_315()
    build_415()
    build_205()
