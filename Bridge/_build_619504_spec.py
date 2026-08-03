"""Build 619-504.json — long-term one-lane Family 2 with positive barrier.

Structural break from 302/402: NO protective vehicle, NO roll-ahead table;
positive barrier + flare rates (504-03) instead. Order-table walk is Merging
Taper + advance signs (buffer not dimensioned on plan despite 504-02 column).
"""
from __future__ import annotations

import copy
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
ref = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
draft = json.loads((ROOT / "Data/sheet-specs/_draft_619504_tables.json").read_text(encoding="utf-8"))
s = copy.deepcopy(ref)

s["sheet"] = {
    "number": "619-504",
    "title": "WORK ZONE TRAFFIC CONTROL MULTI-LANE DIVIDED ROADWAY AND FREEWAY RIGHT LANE CLOSURE",
    "series": "WORK ZONE TRAFFIC CONTROL",
    "operation": "LONG TERM OPERATION",
    "units": "U.S. CUSTOMARY",
    "scale": "NOT TO SCALE",
    "approved": "2026-05-06",
    "issuedUnder": "E3 revision",
    "signedBy": None,
    "sourceUrl": "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/619-504_E3.pdf",
    "localPdf": "Bridge/captures/619-504.pdf",
    "localRender": None,
    "pdfPages": 2,
    "transcribedBy": "Cursor (long-term barrier sibling; tables from _draft_619504_tables.json; corridor from extract_plan_geometry — no roll ahead / no PV)",
    "transcribedOn": "2026-08-03",
    "provenanceNote": (
        "Long-term Family 2 sibling. 504-01/02 identical to 402-04/03. No PV or roll-ahead "
        "tables — temporary positive barrier (504-03 flare rates, Notes 4-7). Plan dimensions "
        "A/B/C + shoulder taper L/3 + merging taper L; buffer column exists in 504-02 but is "
        "NOT labeled on the plan. Sign mid-advance is W20-5 (table), Note 2 uses W20-5L for left."
    ),
}

s["applicability"]["duration"] = "Long Term"
s["applicability"]["durationDefinition"] = (
    "Stationary work that occupies a location more than 3 consecutive days (Note 1)."
)
s["applicability"]["closure"] = "Right lane closure"
s["applicability"]["closureNote"] = (
    "Note 2: left lane closures — substitute W20-5L, W4-2L, and OM3-L."
)

s["tableRoles"] = draft["tableRoles"]
s["tables"] = draft["tables"]

# Drop PV/exposure inputs; keep speed/lane/shoulder/roadType
s["inputs"] = [i for i in s["inputs"] if i["id"] not in (
    "exposureCondition", "closureType", "roadTypeForProtectiveVehicle"
)]
for inp in s["inputs"]:
    inp["usedBy"] = [
        u.replace("302-02", "504-02").replace("302-03", "504-01")
         .replace("302-04", "504-05").replace("302-05", "504-02")
         .replace("302-01", "504-02")
        for u in inp.get("usedBy", [])
    ]

# Corridor without roll ahead / PV cluster
s["corridor"] = {
    "confidence": "drawing",
    "description": (
        "Long-term right lane closure with temporary positive barrier. "
        "Upstream: advance signs C/B/A, shoulder taper overlay on A, merging taper L "
        "(channelizing — barrier must NOT sit on the merging taper per Note 4), then "
        "barrier along the closed lane with tapered end / impact attenuator (flare from "
        "Table 504-03). No roll-ahead or protective-vehicle zones on this sheet."
    ),
    "zones": [
        {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK XX"},
        {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C", "sheetReference": "(SEE TABLE 504-01)",
         "lengthSource": {"table": "504-01", "column": "C", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True, "spans": "W20-1 to W20-5"},
        {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-5", "sheetLegend": "RIGHT LANE CLOSED YY"},
        {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B", "sheetReference": "(SEE TABLE 504-01)",
         "lengthSource": {"table": "504-01", "column": "B", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True, "spans": "W20-5 to W4-2R"},
        {"id": "signA", "order": 5, "kind": "sign", "signCode": "W4-2R", "sheetLegend": "(merge symbol)"},
        {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A", "sheetReference": "(SEE TABLE 504-01)",
         "lengthSource": {"table": "504-01", "column": "A", "lookupBy": ["roadTypeForSignSpacing"]},
         "dimensioned": True, "spans": "W4-2R to upstream end of MERGING TAPER",
         "containsOverlay": "shoulderTaper"},
        {"id": "shoulderTaper", "order": 7, "kind": "taper", "sheetLabel": "SHOULDER TAPER",
         "sheetReference": "(SEE TABLE 504-02)",
         "lengthSource": {"table": "504-02", "column": "shoulderTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
         "dimensioned": True, "consumesStation": False, "containedIn": "gapA",
         "stationAnchor": {"zone": "laneTaper", "end": "upstream"}},
        {"id": "laneTaper", "order": 8, "kind": "taper", "sheetLabel": "MERGING TAPER",
         "sheetReference": "(SEE TABLE 504-02)",
         "lengthSource": {"table": "504-02", "column": "laneTaper",
                          "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"]},
         "dimensioned": True,
         "note": "Note 4: temporary positive barrier shall NOT be placed along the merging taper."},
        {"id": "positiveBarrier", "order": 9, "kind": "barrier", "sheetLabel": "TEMPORARY POSITIVE BARRIER",
         "sheetReference": "(SEE TABLE 504-03)",
         "lengthSource": None,
         "lengthNote": "Project-specific work-area length; flare rates from Table 504-03.",
         "dimensioned": False,
         "note": "Notes 4-7. Movable barrier option (Note 5) can reopen the lane during peaks."},
        {"id": "workArea", "order": 10, "kind": "workArea", "sheetLabel": "WORK AREA",
         "lengthSource": None, "hatched": True, "dimensioned": False},
        {"id": "downstreamTaper", "order": 11, "kind": "taper", "sheetLabel": "DOWNSTREAM TAPER",
         "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}}, "sheetText": "50'-100'",
         "dimensioned": True},
        {"id": "gapEndRoadWork", "order": 12, "kind": "gap", "sheetLabel": None,
         "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}, "dimensioned": False,
         "spans": "End of downstream taper to G20-2"},
        {"id": "signEndRoadWork", "order": 13, "kind": "sign", "signCode": "G20-2",
         "sheetLegend": "END ROAD WORK"},
    ],
}

s["orderTable"] = {
    "confidence": "drawing",
    "description": "No ROLL AHEAD / BUFFER / Vehicle Space rows — long-term barrier sheet.",
    "alignments": [
        {
            "alignIdx": 1,
            "name": "Upstream",
            "station0": "Upstream end of the MERGING TAPER / barrier start",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "laneTaper", "label": "MERGING TAPER"},
                {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W4-2R", "spacingZone": "gapA"},
                {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W20-5", "spacingZone": "gapB"},
                {"rowNum": 4, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
            ],
            "overlayZones": [
                {"zone": "shoulderTaper", "anchor": {"zone": "laneTaper", "end": "upstream"},
                 "direction": "upstream",
                 "note": "L/3 overlay inside gap A."}
            ],
            "excludedRows": [
                {"label": "ROLL AHEAD DISTANCE", "reason": "No roll-ahead table/zone on long-term barrier sheet."},
                {"label": "BUFFER SPACE", "reason": "Buffer column exists in 504-02 but is not dimensioned on the plan."},
                {"label": "Vehicle Space", "reason": "Not on this sheet."},
                {"label": "Upstream Taper Temp Barrier", "reason": "Barrier is longitudinal along work area, not an upstream taper row."},
                {"label": "Upstream Taper Box/Corr Beam", "reason": "Flare rates in 504-03; not a sequential order-table row."},
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

# Signs — W20-5 not W20-5R; add OM3-L / NYR9-11 / R2 from size table
s["signs"] = {
    "confidence": "verbatim",
    "note": "Mid advance is W20-5 (table 504-05). Note 2 left-closure uses W20-5L + W4-2L + OM3-L.",
    "items": [],
}
for row in s["tables"]["504-05"]["rows"]:
    code = row["signCode"]
    base = {
        "signCode": code,
        "sheetLegend": None,
        "legendSubstitution": None,
        "shape": "diamond" if code.startswith("W") else ("flag" if "FLAG" in code else "rectangle"),
        "postMounted": "FLAG" not in code and code != "NYW8-33",
        "corridorZone": None,
        "sizeNonFreeway": row.get("NON-FREEWAY"),
        "sizeFreeway": row.get("FREEWAY"),
        "signLibraryKey": None,
    }
    if code == "W20-1":
        base.update({
            "sheetLegend": "ROAD WORK XX",
            "legendSubstitution": {"placeholder": "XX", "table": "504-01", "column": "XX"},
            "corridorZone": "signC", "signLibraryBase": "W20-01R", "warningFlags": True,
        })
    elif code in ("W20-5", "W20-5R"):
        base["signCode"] = "W20-5"
        base.update({
            "sheetLegend": "RIGHT LANE CLOSED YY",
            "legendSubstitution": {"placeholder": "YY", "table": "504-01", "column": "YY"},
            "corridorZone": "signB", "signLibraryBase": "W20-05R", "warningFlags": False,
        })
    elif code == "W4-2R":
        base.update({"corridorZone": "signA", "signLibraryKey": "W04-02R", "warningFlags": True})
    elif code == "G20-2":
        base.update({"corridorZone": "signEndRoadWork", "signLibraryKey": "G20-02"})
    elif code == "WARNING FLAG":
        base.update({"postMounted": False, "mountedOn": "W20-1, W4-2R"})
    elif code.startswith("R2") or "NYR2" in code:
        base.update({"required": True, "signLibraryKey": "R2-1",
                     "note": "Listed on 504-05; 402-style mid-A/B placement not restated in 504 notes."})
    s["signs"]["items"].append(base)

# Deduplicate W20-5 if both forms appeared
seen = set()
uniq = []
for i in s["signs"]["items"]:
    if i["signCode"] in seen:
        continue
    seen.add(i["signCode"])
    uniq.append(i)
s["signs"]["items"] = uniq

s["symbols"] = {
    "confidence": "drawing",
    "items": [
        {"id": "arrowPanel", "sheetLabel": "ARROW PANEL", "required": True,
         "stationAnchor": {"zone": "laneTaper", "end": "upstream"},
         "note": "Note 5: when movable barrier reopens the lane, arrow panel at shoulder taper end in CAUTION mode."},
        {"id": "positiveBarrier", "sheetLabel": "TEMPORARY POSITIVE BARRIER", "required": True,
         "stationAnchor": {"zone": "positiveBarrier", "end": "both"},
         "note": "Flare rates Table 504-03. Not on merging taper (Note 4)."},
        {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
         "longitudinalSpacing": {"maxFt": 40, "sheetText": "Table 504-04 spacing rows include 20 FT and 40 FT — confirm provision."},
         "runs": [
             {"id": "shoulderTaperRun", "zone": "shoulderTaper",
              "deviceCountSource": {"table": "504-02", "column": "shoulderTaper.devices"}},
             {"id": "mergingTaperRun", "zone": "laneTaper",
              "deviceCountSource": {"table": "504-02", "column": "laneTaper.devices"}},
             {"id": "downstreamRun", "zone": "downstreamTaper", "deviceCountSource": None},
         ],
        },
        {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True,
         "stationAnchor": {"zone": "workArea", "end": "both"}, "hatched": True},
    ],
}

s["annotations"] = {
    "confidence": "drawing",
    "dimensions": [
        {"zone": "gapC", "label": "C", "reference": "(SEE TABLE 504-01)"},
        {"zone": "gapB", "label": "B", "reference": "(SEE TABLE 504-01)"},
        {"zone": "gapA", "label": "A", "reference": "(SEE TABLE 504-01)"},
        {"zone": "shoulderTaper", "label": "SHOULDER TAPER", "reference": "(SEE TABLE 504-02)"},
        {"zone": "laneTaper", "label": "MERGING TAPER", "reference": "(SEE TABLE 504-02)"},
        {"zone": "downstreamTaper", "label": "50'-100' DOWNSTREAM TAPER", "reference": None},
    ],
    "labels": [
        {"text": "TEMPORARY POSITIVE BARRIER", "zone": "positiveBarrier"},
        {"text": "ARROW PANEL", "zone": "laneTaper"},
    ],
}

s["details"] = {"504A": {"title": "LONG TERM SHOULDER CLOSURE DETAIL", "note": "Referenced by Note 5 when movable barrier reopens the travel lane."}}

printed = [n for n in draft["notes"]["printed"] if not n.startswith("N")]
s["notes"] = {"confidence": "verbatim", "printed": printed, "planCallouts": [], "tableNotes": []}

s["rules"] = [
    {"id": "no-roll-ahead-no-pv", "severity": "error", "source": "Sheet structure",
     "assert": "Do not emit ROLL AHEAD DISTANCE or protective-vehicle rows for this sheet.",
     "commonFailure": "Cloning the 302/402 upstream walk including roll ahead."},
    {"id": "barrier-not-on-merging-taper", "severity": "error", "source": "Note 4",
     "assert": "Temporary positive barrier must not be placed along the merging taper.",
     "commonFailure": "Running barrier through the taper."},
    {"id": "shoulder-taper-is-an-overlay", "severity": "error", "source": "Plan dimensions",
     "assert": "Shoulder taper overlays gap A.",
     "commonFailure": "Sequential shoulder-taper station."},
    {"id": "sign-order", "severity": "error", "source": "Plan layout",
     "assert": "Upstream signs W4-2R, then W20-5, then W20-1.",
     "commonFailure": "Reversing order or using W20-5a from the two-lane sibling."},
    {"id": "flare-rates", "severity": "warning", "source": "Table 504-03",
     "assert": "Barrier end flare rates come from Table 504-03 by barrier type and speed.",
     "commonFailure": "Using a fixed flare for every speed."},
]

s["knownCodeDeviations"] = [
    {"id": "barrier-placement-unimplemented", "severity": "error",
     "assert": "Current placement path has no temporary positive barrier / flare-rate placer."},
    {"id": "buffer-table-but-not-on-plan", "severity": "warning",
     "assert": "504-02 still has a buffer column identical to 402, but the plan does not dimension BUFFER SPACE — do not invent a buffer station."},
]

s["knownExcerpts"] = {
    "from619-402": ["504-01 == 402-04", "504-02 == 402-03 including 800/20/21"],
    "differsFrom402": ["No PV / roll-ahead", "504-03 flare rates new", "Long-term notes", "Barrier on plan"],
}

out = ROOT / "Data/sheet-specs/619-504.json"
out.write_text(json.dumps(s, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print("wrote", out, "signs", [i["signCode"] for i in s["signs"]["items"]])
