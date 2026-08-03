"""Build complete Family 6 sheet specs (flagger / two-lane two-way family)."""
from __future__ import annotations

import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
SPEC = ROOT / "Data/sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"


def load_draft(n: int) -> dict:
    name = f"_draft_619{n:03d}_tables.json" if n < 100 else f"_draft_619{n}_tables.json"
    return json.loads((SPEC / name).read_text(encoding="utf-8"))


def write(name: str, spec: dict) -> None:
    spec = finalize(spec)
    path = SPEC / f"{name}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("Wrote", path.name)


def sheet_meta(n: int, title: str, op: str, pages: int, rot: int, note: str,
               approved="2021-12-02", ei="EI 21-028") -> dict:
    num = f"619-{n:03d}" if n < 100 else f"619-{n}"
    return {
        "number": num,
        "title": title,
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": op,
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": approved,
        "issuedUnder": ei,
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/{num}.pdf",
        "localPdf": f"Bridge/captures/{num}.pdf",
        "localRender": None,
        "pdfPages": pages,
        "pageRotation": rot,
        "transcribedBy": "Cursor (Family 6 two-lane two-way / flagger)",
        "transcribedOn": "2026-08-03",
        "provenanceNote": note,
    }


def applicability_flagger(speeds, buffer_note=""):
    return {
        "roadType": "Non-Freeway",
        "roadway": "Two-lane two-way",
        "lanesPerDirection": 1,
        "closure": "Lane closure with flaggers",
        "duration": "Short Term",
        "durationDefinition": "Daytime work that occupies a location for more than 1 hour within a single daylight period.",
        "speedRangeMph": {
            "min": min(speeds), "max": max(speeds), "increment": 5,
            "note": buffer_note or "Buffer table rows for listed speeds only.",
        },
        "laneWidthFt": None,
        "laneWidthNote": "No merging/lane taper on flagger sheets — lane width is not a lookup input.",
        "shoulderWidthBands": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"],
        "areaTypes": ["URBAN", "RURAL"],
        "areaTypeNote": "Advance-warning spacing table keyed on URBAN vs RURAL.",
    }


def inputs_flagger(aw_id, buf_id, roll_id, pv_id, size_id, speeds):
    out = [
        {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": speeds,
         "usedBy": [x for x in [aw_id, buf_id, roll_id, pv_id] if x]},
        {"id": "shoulderWidthBand", "type": "enum",
         "allowed": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "default": ">= 8 ft", "usedBy": []},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": [size_id] if size_id else []},
    ]
    if aw_id:
        out.insert(1, {"id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL"],
                       "usedBy": [aw_id]})
    if pv_id:
        out.extend([
            {"id": "exposureCondition", "type": "enum",
             "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                         "OTHER HAZARDS NO WORKERS EXPOSED"],
             "usedBy": [pv_id]},
            {"id": "closureType", "type": "enum",
             "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
             "default": "LANE CLOSURE OR ENCROACHMENT",
             "usedBy": [pv_id]},
        ])
    return out


def aw_table(tid: str, rows: list) -> dict:
    return {
        "title": "ADVANCE WARNING SIGN SPACING",
        "confidence": "verbatim",
        "keyedBy": ["areaType", "preconstructionPostedSpeedMph"],
        "columnMeaning": {
            "A": "DISTANCE BETWEEN SIGNS - A (FT.)",
            "B": "DISTANCE BETWEEN SIGNS - B (FT.)",
            "C": "DISTANCE BETWEEN SIGNS - C (FT.)",
            "XX": "SIGN LEGEND for W20-1: ROAD WORK XX",
            "YY": "SIGN LEGEND for W20-4: ONE LANE ROAD YY",
        },
        "note": "Cell-identical to 619-311 table 311-03 / 619-011 table 011-06 NON-FREEWAY rows.",
        "rows": rows,
    }


def buffer_table(tid: str, rows: list) -> dict:
    return {
        "title": "LONGITUDINAL BUFFER SPACE",
        "confidence": "verbatim",
        "keyedBy": ["preconstructionPostedSpeedMph"],
        "columnMeaning": {
            "longitudinalBufferSpace": "LONGITUDINAL BUFFER SPACE DISTANCE (FT.)/ # OF SKIP LINES",
        },
        "note": "Buffer-only table — NO laneTaper / shoulderTaper columns (flagger operation, not merging).",
        "rows": rows,
    }


def roll_table(tid: str, rows: list) -> dict:
    return {
        "title": "ROLL AHEAD DISTANCE",
        "confidence": "verbatim",
        "keyedBy": ["preconstructionPostedSpeedMph"],
        "columnMeaning": "ROLL AHEAD DISTANCE (FT.)/# OF SKIP LINES — STATIONARY OPERATION MIN/MAX",
        "rows": rows,
        "usageNote": "MIN/MAX range, not a single value.",
    }


def pv_table(tid: str, draft_pv: dict) -> dict:
    return {
        "title": "PROTECTIVE VEHICLE REQUIREMENTS",
        "confidence": "verbatim",
        "keyedBy": ["closureType", "exposureCondition", "preconstructionPostedSpeedMph"],
        "roadTypeScope": "NON-FREEWAY",
        "speedBands": draft_pv["speedBands"],
        "rows": draft_pv["rows"],
        "legend": draft_pv["legend"],
        "tableNotes": draft_pv["tableNotes"],
        "note": "Optional on flagger sheets (see plan Note: PV may be used; if used, provide buffer).",
    }


def size_table(tid: str, rows: list) -> dict:
    return {
        "title": "REQUIRED SIGN SIZES",
        "confidence": "verbatim",
        "keyedBy": ["signCode", "signSizeClass"],
        "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
        "rows": rows,
    }


def flagger_corridor(aw_id: str, buf_id: str, roll_id: str) -> dict:
    """Corridor for base flagger: A/B/C spacing + buffer + optional PV/roll + downstream taper. No merging taper."""
    return {
        "confidence": "drawing",
        "description": "Two-lane two-way flagger: advance W20-1 / W20-4 / W20-7 on each approach; "
                       "optional PV with buffer + roll-ahead; NO merging/lane taper. "
                       "Downstream 50'-100' taper + G20-2.",
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1",
             "sheetLegend": "ROAD WORK XX"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
             "sheetReference": f"(SEE TABLE {aw_id})",
             "lengthSource": {"table": aw_id, "column": "C",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True, "spans": "W20-1 to W20-4"},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-4",
             "sheetLegend": "ONE LANE ROAD YY"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
             "sheetReference": f"(SEE TABLE {aw_id})",
             "lengthSource": {"table": aw_id, "column": "B",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True, "spans": "W20-4 to W20-7",
             "note": "W3-4 (BE PREPARED TO STOP) may be added at B/2 per Note 5 if queue past W20-4."},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W20-7",
             "sheetLegend": "FLAGGER (symbol)"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A",
             "sheetReference": f"(SEE TABLE {aw_id})",
             "lengthSource": {"table": aw_id, "column": "A",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True,
             "spans": "W20-7 to flagger station / buffer",
             "note": "Centerline cones optional per Note 3 — place 100 ft min from flagger."},
            {"id": "bufferSpace", "order": 7, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "sheetReference": f"(SEE TABLE {buf_id})",
             "lengthSource": {"table": buf_id, "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True,
             "note": "Required when protective vehicle is used (plan note)."},
            {"id": "protectiveVehicle1", "order": 8, "kind": "symbol", "sheetLabel": "PV",
             "lengthSource": None, "required": False,
             "note": "Optional — if conditions warrant (plan note)."},
            {"id": "rollAheadDistance", "order": 9, "kind": "clearance",
             "sheetLabel": "ROLL AHEAD DISTANCE",
             "sheetReference": f"(SEE TABLE {roll_id})",
             "lengthSource": {"table": roll_id, "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True, "mustBeEmpty": True},
            {"id": "workArea", "order": 10, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False,
             "note": "Project-specific. Flaggers at each end of the one-lane section."},
            {"id": "downstreamTaper", "order": 11, "kind": "taper",
             "sheetLabel": "DOWNSTREAM TAPER",
             "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}},
             "sheetText": "50'-100'", "dimensioned": True},
            {"id": "gapEndRoadWork", "order": 12, "kind": "gap", "sheetLabel": None,
             "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}},
             "dimensioned": False,
             "sheetText": "THIS SIGN SHALL BE LOCATED A MINIMUM DISTANCE OF 80 FT AND MAXIMUM OF 400 FT PAST THE END OF THE DOWNSTREAM TAPER."},
            {"id": "signEndRoadWork", "order": 13, "kind": "sign", "signCode": "G20-2",
             "sheetLegend": "END ROAD WORK"},
        ],
    }


def flagger_order(include_pv_roll=True) -> dict:
    up_rows = []
    n = 1
    if include_pv_roll:
        up_rows.append({"rowNum": n, "type": "Non-Sign", "zone": "rollAheadDistance",
                        "label": "ROLL AHEAD DISTANCE"}); n += 1
        up_rows.append({"rowNum": n, "type": "Non-Sign", "zone": "bufferSpace",
                        "label": "BUFFER SPACE"}); n += 1
    up_rows += [
        {"rowNum": n, "type": "Sign", "zone": "signA", "signCode": "W20-7", "spacingZone": "gapA"},
        {"rowNum": n + 1, "type": "Sign", "zone": "signB", "signCode": "W20-4", "spacingZone": "gapB"},
        {"rowNum": n + 2, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
    ]
    excluded = [
        {"label": "LANE TAPER", "reason": "No merging/lane taper on flagger sheets."},
        {"label": "MERGING TAPER", "reason": "No merging taper on flagger sheets."},
        {"label": "SHOULDER TAPER", "reason": "No shoulder taper on base flagger sheets."},
        {"label": "Vehicle Space", "reason": "Not dimensioned as a separate order-table row."},
        {"label": "Upstream Taper Temp Barrier", "reason": "No barrier on this sheet."},
        {"label": "Upstream Taper Box/Corr Beam", "reason": "No box/corr beam on this sheet."},
    ]
    return {
        "confidence": "drawing",
        "description": "Flagger approach walk: optional roll+buffer, then W20-7 / W20-4 / W20-1. "
                       "Downstream: taper + G20-2. W3-4 is conditional (Note 5) — not a default row.",
        "alignments": [
            {
                "alignIdx": 1, "name": "Upstream",
                "station0": "Upstream edge of WORK AREA / flagger station",
                "walkDirection": "Upstream, against traffic",
                "rows": up_rows,
                "excludedRows": excluded,
            },
            {
                "alignIdx": 2, "name": "Downstream",
                "station0": "Downstream edge of WORK AREA",
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


def flagger_signs(aw_id: str, size_rows: list, extra=None) -> dict:
    by = {r["signCode"]: r for r in size_rows}
    items = [
        {
            "signCode": "W20-1", "sheetLegend": "ROAD WORK XX",
            "legendSubstitution": {"placeholder": "XX", "table": aw_id, "column": "XX"},
            "shape": "diamond", "warningFlags": True, "postMounted": True,
            "corridorZone": "signC",
            "sizeNonFreeway": by.get("W20-1", {}).get("NON-FREEWAY") or "36x36",
            "sizeFreeway": by.get("W20-1", {}).get("FREEWAY") or "48x48",
            "signLibraryKey": None, "signLibraryBase": "W20-01R",
            "signLibraryNote": "Suffix from XX: AHEAD->A, feet->F, mile->M.",
        },
        {
            "signCode": "W20-4", "sheetLegend": "ONE LANE ROAD YY",
            "legendSubstitution": {"placeholder": "YY", "table": aw_id, "column": "YY"},
            "shape": "diamond", "warningFlags": False, "postMounted": True,
            "corridorZone": "signB",
            "sizeNonFreeway": by.get("W20-4", {}).get("NON-FREEWAY") or "36x36",
            "sizeFreeway": by.get("W20-4", {}).get("FREEWAY") or "48x48",
            "signLibraryKey": None, "signLibraryBase": "W20-04",
            "signLibraryNote": "Suffix from YY: AHEAD->A, feet->F, mile->M.",
        },
        {
            "signCode": "W20-7", "sheetLegend": "FLAGGER (symbol)",
            "legendSubstitution": None, "shape": "diamond", "warningFlags": True,
            "postMounted": True, "corridorZone": "signA",
            "sizeNonFreeway": by.get("W20-7", {}).get("NON-FREEWAY") or "36x36",
            "sizeFreeway": by.get("W20-7", {}).get("FREEWAY") or "48x48",
            "signLibraryKey": "W20-07",
            "note": "Remove/cover when flagging not occurring (plan note).",
        },
        {
            "signCode": "W3-4", "sheetLegend": "BE PREPARED TO STOP",
            "legendSubstitution": None, "shape": "diamond", "warningFlags": False,
            "postMounted": True, "corridorZone": "gapB",
            "sizeNonFreeway": by.get("W3-4", {}).get("NON-FREEWAY") or "36x36",
            "sizeFreeway": by.get("W3-4", {}).get("FREEWAY") or "48x48",
            "signLibraryKey": "W03-04",
            "note": "Conditional — add if queue expected past W20-4 (Note 5). Not a default order-table row.",
        },
        {
            "signCode": "G20-2", "sheetLegend": "END ROAD WORK",
            "legendSubstitution": None, "shape": "rectangle", "warningFlags": False,
            "postMounted": True, "corridorZone": "signEndRoadWork",
            "sizeNonFreeway": by.get("G20-2", {}).get("NON-FREEWAY") or "36x18",
            "sizeFreeway": by.get("G20-2", {}).get("FREEWAY") or "48x24",
            "signLibraryKey": "G20-02",
        },
        {
            "signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG",
            "legendSubstitution": None, "shape": "flag", "warningFlags": False,
            "postMounted": False, "mountedOn": "W20-1",
            "sizeNonFreeway": "18x18", "sizeFreeway": "18x18",
        },
    ]
    if extra:
        items.extend(extra)
    return {"confidence": "verbatim", "items": items}


def flagger_rules() -> list:
    return [
        {
            "id": "no-merging-taper",
            "severity": "error",
            "source": "Sheet structure / flagger operation",
            "assert": "Do not place MERGING TAPER or LANE TAPER — flagger operation uses stop/slow control, not a merge.",
            "commonFailure": "Cloning 302/311 merging taper corridor onto a flagger sheet.",
        },
        {
            "id": "optional-pv",
            "severity": "warning",
            "source": "Plan note (PV may be used)",
            "assert": "Protective vehicle with roll-ahead is optional; if used, provide buffer space.",
            "commonFailure": "Always requiring PV or omitting buffer when PV is placed.",
        },
        {
            "id": "w3-4-conditional",
            "severity": "warning",
            "source": "Plan Note (queue past W20-4)",
            "assert": "Add W3-4 if traffic is expected to queue past the W20-4 sign.",
            "commonFailure": "Always placing W3-4 as a default order-table row.",
        },
    ]


def finalize(spec: dict) -> dict:
    """Fill annotations for dimensioned zones; prune symbol anchors to existing zones;
    ensure every size-table code appears in signs.items; fix WARNING FLAG mountedOn."""
    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    dims = []
    for z in spec["corridor"]["zones"]:
        if not z.get("dimensioned"):
            continue
        label = z.get("sheetLabel") or z.get("sheetText") or z["id"]
        dims.append({
            "zone": z["id"],
            "label": label,
            "reference": z.get("sheetReference"),
        })
    spec["annotations"] = {"confidence": "drawing", "dimensions": dims, "callouts": []}

    sym_items = []
    for sym in spec.get("symbols", {}).get("items", []):
        anchor = sym.get("stationAnchor")
        if anchor and anchor.get("zone") not in zones:
            sym = dict(sym)
            sym.pop("stationAnchor", None)
            sym["note"] = (sym.get("note") or "") + " (no rollAhead zone on this sheet)."
        sym_items.append(sym)
    # Drop PV symbol if no PV-related zone
    if "rollAheadDistance" not in zones and "protectiveVehicle1" not in zones:
        sym_items = [s for s in sym_items if s.get("id") != "protectiveVehicle"]
    spec["symbols"] = {"confidence": "drawing", "items": sym_items or [
        {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
         "longitudinalSpacing": {"maxFt": 40}},
    ]}

    size_id = spec.get("tableRoles", {}).get("signSizes")
    if size_id and size_id in spec["tables"]:
        present = {s["signCode"] for s in spec["signs"]["items"]}
        for row in spec["tables"][size_id]["rows"]:
            code = row["signCode"]
            if code in present:
                continue
            key = {
                "WARNING FLAG": None, "W20-7a": "W20-07A", "NYR9-11": "NYR9-11",
                "R10-6": "R10-06", "W7-3a": "W07-03A", "G20-1": "G20-01",
            }.get(code)
            item = {
                "signCode": code,
                "shape": "flag" if code == "WARNING FLAG" else (
                    "diamond" if code.startswith("W") else "rectangle"),
                "postMounted": code != "WARNING FLAG",
                "corridorZone": "workArea" if "workArea" in zones else next(iter(zones)),
                "sizeNonFreeway": row.get("NON-FREEWAY") or "36x36",
                "sizeFreeway": row.get("FREEWAY") or row.get("NON-FREEWAY") or "48x48",
            }
            if code == "WARNING FLAG":
                item["mountedOn"] = "W20-1"
            if key:
                item["signLibraryKey"] = key
            else:
                item["signLibraryKey"] = None
                item["note"] = "Present on size table; confirm library key if placing."
            spec["signs"]["items"].append(item)
            present.add(code)

    for s in spec["signs"]["items"]:
        if s.get("signCode") == "WARNING FLAG" and not s.get("postMounted"):
            s.setdefault("mountedOn", "W20-1")

    if not spec.get("rules"):
        spec["rules"] = flagger_rules()
    # normalize any leftover bad rules
    fixed = []
    for r in spec["rules"]:
        if "assert" not in r and "statement" in r:
            r = {
                "id": r["id"],
                "severity": "error" if r.get("severity") in (True, "true", "error", False, "false")
                and r.get("severity") not in (True, "true", "warning") else "warning",
                "source": "Sheet / Family 6",
                "assert": r["statement"],
                "commonFailure": "Mis-applying corridor clone from Family 1/2.",
            }
            if r["id"] == "no-merging-taper":
                r["severity"] = "error"
        if r.get("severity") not in ("error", "warning"):
            r["severity"] = "warning"
        for f in ("source", "assert", "commonFailure"):
            r.setdefault(f, "Family 6")
        fixed.append(r)
    spec["rules"] = fixed
    spec.setdefault("knownAnomalies", [])
    spec.setdefault("openQuestions", [])
    spec.setdefault("notes", {"confidence": "verbatim", "items": []})
    return spec


def common_tail() -> dict:
    return {
        "symbols": {
            "confidence": "drawing",
            "items": [
                {"id": "flagger", "sheetLabel": "FLAGGER", "required": True,
                 "note": "24\" min STOP/SLOW paddle preferred."},
                {"id": "protectiveVehicle", "sheetLabel": "PV", "required": False,
                 "stationAnchor": {"zone": "rollAheadDistance", "end": "upstream"}},
                {"id": "centerlineCones", "sheetLabel": "CENTERLINE CONES", "required": False,
                 "note": "Optional — enhance flagger visibility; 100 ft min from flagger."},
                {"id": "channelizingDevices", "sheetLabel": "CHANNELIZING DEVICES", "required": True,
                 "longitudinalSpacing": {"maxFt": 40, "sheetText": "Not to exceed 40' (1 skip line)."}},
            ],
        },
        "annotations": {"confidence": "drawing", "dimensions": [], "callouts": []},
        "notes": {"confidence": "verbatim", "items": []},
        "rules": flagger_rules(),
        "knownAnomalies": [],
        "openQuestions": [],
    }


def build_flagger_base(n: int, title: str, op: str, draft: dict, roles: dict,
                       speeds=None, extra_signs=None, notes_extra="") -> dict:
    speeds = speeds or [25, 30, 35, 40, 45, 50, 55]
    aw = roles["advanceWarningSpacing"]
    buf = roles["taperAndBuffer"]
    sizes = roles["signSizes"]
    pv = roles.get("protectiveVehicle")
    roll = roles.get("rollAheadDistance")
    pages = draft["meta"]["pages"]
    rot = draft["meta"]["rotation"]
    tables = {
        aw: aw_table(aw, draft["advanceWarning"]),
        buf: buffer_table(buf, draft["bufferOnly"]),
        sizes: size_table(sizes, draft["signSizes"]),
    }
    if pv and draft.get("protectiveVehicle"):
        tables[pv] = pv_table(pv, draft["protectiveVehicle"])
    if roll and draft.get("rollAhead"):
        tables[roll] = roll_table(roll, draft["rollAhead"])
    if roles.get("channelizingApplication"):
        tables[roles["channelizingApplication"]] = {
            "title": "CHANNELIZING DEVICE APPLICATION",
            "confidence": "drawing",
            "keyedBy": [],
            "rows": [],
            "note": "Channelizing matrix present on sheet — spacing rules encoded in symbols/rules.",
        }

    spec = {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(
            n, title, op, pages, rot,
            f"Family 6 flagger. Buffer-only taperAndBuffer (no lane/shoulder cols). "
            f"AW/buffer/PV/roll identity with 311 where overlapping. {notes_extra}"),
        "applicability": applicability_flagger(speeds),
        "inputs": inputs_flagger(aw, buf, roll, pv, sizes, speeds),
        "tableRoles": roles,
        "tables": tables,
        "corridor": flagger_corridor(aw, buf, roll or buf),
        "orderTable": flagger_order(include_pv_roll=bool(roll and pv)),
        "signs": flagger_signs(aw, draft["signSizes"], extra=extra_signs),
        "geometry": {
            "confidence": "drawing",
            "stationing": {
                "origin": "Upstream edge of WORK AREA",
                "positiveDirection": "Upstream, against travel",
            },
            "crossSection": {
                "description": "Two-lane two-way roadway with flagger-controlled one-lane operation.",
            },
        },
        **common_tail(),
    }
    # If no roll table, strip roll zone length dependency from corridor/order
    if not roll:
        spec["corridor"]["zones"] = [
            z for z in spec["corridor"]["zones"]
            if z["id"] not in ("rollAheadDistance", "protectiveVehicle1")
        ]
        # renumber
        for i, z in enumerate(spec["corridor"]["zones"], 1):
            z["order"] = i
        spec["orderTable"] = flagger_order(include_pv_roll=False)
        # buffer may still be present without PV
        if not pv:
            # keep buffer as optional — still in corridor if buf table exists
            pass
    return spec


def build_307():
    d = load_draft(307)
    roles = d["tableRoles"]
    write("619-307", build_flagger_base(
        307,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY LANE CLOSURE WITH FLAGGERS",
        "SHORT TERM OPERATION", d, roles,
        notes_extra="Family 6 reference sheet."))


def build_308():
    d = load_draft(308)
    roles = d["tableRoles"]
    write("619-308", build_flagger_base(
        308,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY LANE CLOSURE WITH FLAGGERS PRIOR TO INTERSECTION",
        "SHORT TERM OPERATION", d, roles,
        notes_extra="Prior to intersection sibling of 307; tables == 307."))


def build_090_091(n: int, title: str):
    d = load_draft(n)
    roles = d["tableRoles"]
    aw, buf, sizes = roles["advanceWarningSpacing"], roles["taperAndBuffer"], roles["signSizes"]
    speeds = [25, 30, 35, 40, 45, 50, 55]
    prefix = f"{n:03d}"
    # Closure sheets: AW + buffer + signs; no PV/roll; order = buffer + W20-7 + W3-4 + W20-1
    corridor = {
        "confidence": "drawing",
        "description": f"Temporary closure with flaggers. AW A/B/C + buffer. Signs W20-7 / W3-4 / W20-1.",
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1", "sheetLegend": "ROAD WORK XX"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
             "lengthSource": {"table": aw, "column": "C",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W3-4",
             "sheetLegend": "BE PREPARED TO STOP"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"table": aw, "column": "B",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W20-7",
             "sheetLegend": "FLAGGER"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"table": aw, "column": "A",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "bufferSpace", "order": 7, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "lengthSource": {"table": buf, "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 8, "kind": "workArea", "sheetLabel": "WORK AREA / CLOSURE",
             "lengthSource": None, "hatched": True, "dimensioned": False},
        ],
    }
    order = {
        "confidence": "drawing",
        "description": "Closure: BUFFER + W20-7 + W3-4 + W20-1. No roll/taper.",
        "alignments": [{
            "alignIdx": 1, "name": "Upstream",
            "station0": "Closure / work area",
            "walkDirection": "Upstream, against traffic",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W20-7", "spacingZone": "gapA"},
                {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W3-4", "spacingZone": "gapB"},
                {"rowNum": 4, "type": "Sign", "zone": "signC", "signCode": "W20-1", "spacingZone": "gapC"},
            ],
            "excludedRows": [
                {"label": "ROLL AHEAD DISTANCE", "reason": "No roll-ahead table on closure sheet."},
                {"label": "MERGING TAPER", "reason": "Closure / flagger — no merge."},
                {"label": "LANE TAPER", "reason": "No lane taper."},
                {"label": "SHOULDER TAPER", "reason": "No shoulder taper."},
                {"label": "Vehicle Space", "reason": "Not on this sheet."},
            ],
        }],
    }
    by = {r["signCode"]: r for r in d["signSizes"]}
    signs = {"confidence": "verbatim", "items": [
        {"signCode": "W20-1", "sheetLegend": "ROAD WORK XX",
         "legendSubstitution": {"placeholder": "XX", "table": aw, "column": "XX"},
         "shape": "diamond", "postMounted": True, "corridorZone": "signC",
         "sizeNonFreeway": by.get("W20-1", {}).get("NON-FREEWAY") or "36x36",
         "sizeFreeway": by.get("W20-1", {}).get("FREEWAY") or "48x48",
         "signLibraryKey": None, "signLibraryBase": "W20-01R"},
        {"signCode": "W20-7", "sheetLegend": "FLAGGER", "shape": "diamond", "postMounted": True,
         "corridorZone": "signA",
         "sizeNonFreeway": by.get("W20-7", {}).get("NON-FREEWAY") or "36x36",
         "sizeFreeway": by.get("W20-7", {}).get("FREEWAY") or "48x48",
         "signLibraryKey": "W20-07"},
        {"signCode": "W3-4", "sheetLegend": "BE PREPARED TO STOP", "shape": "diamond",
         "postMounted": True, "corridorZone": "signB",
         "sizeNonFreeway": by.get("W3-4", {}).get("NON-FREEWAY") or "36x36",
         "sizeFreeway": by.get("W3-4", {}).get("FREEWAY") or "48x48",
         "signLibraryKey": "W03-04"},
    ]}
    write(f"619-{prefix}", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(n, title, "TEMPORARY CLOSURE", d["meta"]["pages"], d["meta"]["rotation"],
                            "Family 6 temporary closure. AW+buffer only; no PV/roll/taper."),
        "applicability": {
            **applicability_flagger(speeds),
            "duration": "Temporary Closure",
            "closure": "Temporary road/intersection closure",
        },
        "inputs": inputs_flagger(aw, buf, None, None, sizes, speeds),
        "tableRoles": roles,
        "tables": {
            aw: aw_table(aw, d["advanceWarning"]),
            buf: buffer_table(buf, d["bufferOnly"]),
            sizes: size_table(sizes, d["signSizes"]),
        },
        "corridor": corridor,
        "orderTable": order,
        "signs": signs,
        "geometry": {"confidence": "drawing", "stationing": {"origin": "Closure"}, "crossSection": {}},
        **common_tail(),
    })


def build_314():
    d = load_draft(314)
    roles = d["tableRoles"]
    buf, roll, pv, sizes = roles["taperAndBuffer"], roles["rollAheadDistance"], roles["protectiveVehicle"], roles["signSizes"]
    speeds = [25, 30, 35, 40, 45, 50, 55]
    # Fixed 500' gaps — no AW table
    corridor = {
        "confidence": "drawing",
        "description": "Moving flaggers. Fixed 500' sign gaps. Buffer + roll + PV. No AW table.",
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "500'",
             "lengthSource": {"fixedFt": 500}, "dimensioned": True},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-4"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "500'",
             "lengthSource": {"fixedFt": 500}, "dimensioned": True},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W20-7"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "500'",
             "lengthSource": {"fixedFt": 500}, "dimensioned": True},
            {"id": "bufferSpace", "order": 7, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "lengthSource": {"table": buf, "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "protectiveVehicle1", "order": 8, "kind": "symbol", "sheetLabel": "PV",
             "lengthSource": None},
            {"id": "rollAheadDistance", "order": 9, "kind": "clearance",
             "sheetLabel": "ROLL AHEAD DISTANCE",
             "lengthSource": {"table": roll, "column": "range",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 10, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True},
            {"id": "gapEndRoadWork", "order": 11, "kind": "gap",
             "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}},
            {"id": "signEndRoadWork", "order": 12, "kind": "sign", "signCode": "G20-2"},
        ],
    }
    by = {r["signCode"]: r for r in d["signSizes"]}
    signs = {"confidence": "verbatim", "items": [
        {"signCode": "W20-1", "shape": "diamond", "postMounted": True, "corridorZone": "signC",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RA"},
        {"signCode": "W20-4", "shape": "diamond", "postMounted": True, "corridorZone": "signB",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-04A"},
        {"signCode": "W20-7", "shape": "diamond", "postMounted": True, "corridorZone": "signA",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-07"},
        {"signCode": "G20-2", "shape": "rectangle", "postMounted": True, "corridorZone": "signEndRoadWork",
         "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02"},
        {"signCode": "NYW8-33", "shape": "rectangle", "postMounted": True,
         "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": "NYW8-33",
         "note": "LANE CLOSED on PV / arrow panel area."},
        {"signCode": "G20-1", "shape": "rectangle", "postMounted": True,
         "sizeNonFreeway": by.get("G20-1", {}).get("NON-FREEWAY") or "36x18",
         "sizeFreeway": by.get("G20-1", {}).get("FREEWAY") or "48x24",
         "signLibraryKey": "G20-01", "note": "ROAD WORK NEXT X MILES — conditional Note 6."},
        {"signCode": "W7-3a", "shape": "rectangle", "postMounted": True,
         "sizeNonFreeway": by.get("W7-3a", {}).get("NON-FREEWAY") or "24x18",
         "sizeFreeway": by.get("W7-3a", {}).get("FREEWAY") or "36x30",
         "signLibraryKey": "W07-03A", "note": "NEXT X MILES plaque — conditional Note 5."},
        {"signCode": "W3-4", "shape": "diamond", "postMounted": True,
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W03-04"},
        {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
         "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
    ]}
    write("619-314", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(314,
            "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY LANE CLOSURE WITH MOVING FLAGGERS",
            "SHORT TERM OPERATION", d["meta"]["pages"], d["meta"]["rotation"],
            "Moving flaggers. Fixed 500' gaps. Roll 2-band (no <=40). Buffer == 311. No AW table."),
        "applicability": {
            **applicability_flagger(speeds),
            "areaTypes": None,
            "areaTypeNote": "No AW spacing table — plan uses fixed 500' gaps.",
            "speedRangeMph": {
                "min": 25, "max": 55, "increment": 5,
                "note": "Buffer 25-55; roll-ahead only >=45 bands (45-50 and >=55).",
            },
        },
        "inputs": [
            {"id": "preconstructionPostedSpeedMph", "type": "integer", "allowed": speeds,
             "usedBy": [buf, roll, pv]},
            {"id": "shoulderWidthBand", "type": "enum",
             "allowed": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "default": ">= 8 ft", "usedBy": []},
            {"id": "exposureCondition", "type": "enum",
             "allowed": ["WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
                         "OTHER HAZARDS NO WORKERS EXPOSED"], "usedBy": [pv]},
            {"id": "closureType", "type": "enum",
             "allowed": ["LANE CLOSURE OR ENCROACHMENT", "SHOULDER CLOSURE OR ENCROACHMENT"],
             "default": "LANE CLOSURE OR ENCROACHMENT", "usedBy": [pv]},
            {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
             "default": "NON-FREEWAY", "usedBy": [sizes]},
        ],
        "tableRoles": roles,
        "tables": {
            pv: pv_table(pv, d["protectiveVehicle"]),
            roll: roll_table(roll, d["rollAhead"]),
            buf: buffer_table(buf, d["bufferOnly"]),
            sizes: size_table(sizes, d["signSizes"]),
        },
        "corridor": corridor,
        "orderTable": {
            "confidence": "drawing",
            "description": "Moving: ROLL + BUFFER + W20-7/W20-4/W20-1 @500'. Downstream G20-2.",
            "alignments": [
                {"alignIdx": 1, "name": "Upstream", "station0": "Work area",
                 "walkDirection": "Upstream",
                 "rows": [
                     {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance",
                      "label": "ROLL AHEAD DISTANCE"},
                     {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                     {"rowNum": 3, "type": "Sign", "zone": "signA", "signCode": "W20-7",
                      "spacingZone": "gapA"},
                     {"rowNum": 4, "type": "Sign", "zone": "signB", "signCode": "W20-4",
                      "spacingZone": "gapB"},
                     {"rowNum": 5, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                      "spacingZone": "gapC"},
                 ],
                 "excludedRows": [
                     {"label": "MERGING TAPER", "reason": "Moving flagger — no merge."},
                     {"label": "LANE TAPER", "reason": "No lane taper."},
                     {"label": "SHOULDER TAPER", "reason": "No shoulder taper."},
                     {"label": "Vehicle Space", "reason": "Not on this sheet."},
                 ]},
                {"alignIdx": 2, "name": "Downstream", "station0": "Work area downstream",
                 "walkDirection": "Downstream",
                 "rows": [
                     {"rowNum": 1, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                      "spacingZone": "gapEndRoadWork"},
                 ]},
            ],
        },
        "signs": signs,
        "geometry": {"confidence": "drawing", "stationing": {"origin": "Work area"}, "crossSection": {}},
        **common_tail(),
    })


def build_sidewalk(n: int, title: str, duration: str):
    d = load_draft(n)
    roles = d["tableRoles"]
    sizes = roles["signSizes"]
    chan = roles.get("channelizingApplication")
    tables = {sizes: size_table(sizes, d["signSizes"])}
    if chan:
        tables[chan] = {
            "title": "CHANNELIZING DEVICE APPLICATION",
            "confidence": "drawing", "keyedBy": [], "rows": [],
            "note": "Pedestrian channelizing matrix on sheet.",
        }
    # Minimal pedestrian corridor — no AW/buffer/taper lookups
    codes = [r["signCode"] for r in d["signSizes"] if r.get("NON-FREEWAY") or r.get("FREEWAY")]
    if not codes:
        codes = ["G20-2", "R9-9", "R9-11L", "R9-11R", "W20-1"]
    sign_items = []
    for i, code in enumerate(codes):
        row = next((r for r in d["signSizes"] if r["signCode"] == code), {})
        key = {
            "G20-2": "G20-02", "W20-1": "W20-01RA", "R9-9": "R09-09",
            "R9-11L": "R09-11L", "R9-11R": "R09-11R", "R9-10": "R09-10",
            "R11-2": "R11-02", "R8-3": "R08-03",
        }.get(code)
        sign_items.append({
            "signCode": code, "shape": "diamond" if code.startswith("W") else "rectangle",
            "postMounted": True,
            "corridorZone": "workArea",
            "sizeNonFreeway": row.get("NON-FREEWAY") or "24x24",
            "sizeFreeway": row.get("FREEWAY") or row.get("NON-FREEWAY") or "24x24",
            **({"signLibraryKey": key} if key else {"signLibraryKey": None,
                                                     "note": "Confirm SignLibrary key if placing."}),
        })
    write(f"619-{n}", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(n, title, duration, d["meta"]["pages"], d["meta"]["rotation"],
                            "Pedestrian/sidewalk sheet. Sign sizes + channelizing only — "
                            "NO AW/buffer/taper/roll corridor tables."),
        "applicability": {
            "roadType": "Non-Freeway",
            "roadway": "Two-lane two-way (pedestrian)",
            "closure": "Sidewalk detour / diversion",
            "duration": duration.replace(" OPERATION", "").title() if "OPERATION" in duration else duration,
            "speedRangeMph": {"allowed": [25, 30, 35, 40, 45, 50, 55],
                              "note": "No speed-keyed corridor tables on this sheet."},
            "laneWidthFt": None,
            "shoulderWidthBands": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"],
            "areaTypes": None,
        },
        "inputs": [
            {"id": "preconstructionPostedSpeedMph", "type": "integer",
             "allowed": [25, 30, 35, 40, 45, 50, 55], "usedBy": []},
            {"id": "shoulderWidthBand", "type": "enum",
             "allowed": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "default": ">= 8 ft", "usedBy": []},
            {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
             "default": "NON-FREEWAY", "usedBy": [sizes]},
        ],
        "tableRoles": roles,
        "tables": tables,
        "corridor": {
            "confidence": "drawing",
            "description": "Pedestrian detour — channelizing + sidewalk-closed signs; no vehicle corridor taper walk.",
            "zones": [
                {"id": "workArea", "order": 1, "kind": "workArea", "sheetLabel": "SIDEWALK CLOSED / DETOUR",
                 "lengthSource": None, "hatched": True, "dimensioned": False},
                {"id": "signEndRoadWork", "order": 2, "kind": "sign", "signCode": "G20-2",
                 "sheetLegend": "END ROAD WORK"},
                {"id": "gapEndRoadWork", "order": 3, "kind": "gap",
                 "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}},
            ],
        },
        "orderTable": {
            "confidence": "drawing",
            "description": "Pedestrian sheets do not drive a vehicle upstream taper walk. "
                           "Minimal downstream G20-2 only for live-build smoke.",
            "alignments": [{
                "alignIdx": 1, "name": "Upstream",
                "station0": "Sidewalk work",
                "walkDirection": "N/A — pedestrian",
                "rows": [
                    {"rowNum": 1, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                     "spacingZone": "gapEndRoadWork"},
                ],
                "excludedRows": [
                    {"label": "MERGING TAPER", "reason": "Pedestrian sheet."},
                    {"label": "BUFFER SPACE", "reason": "No buffer table."},
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No roll table."},
                    {"label": "LANE TAPER", "reason": "Pedestrian sheet."},
                    {"label": "SHOULDER TAPER", "reason": "Pedestrian sheet."},
                    {"label": "Vehicle Space", "reason": "Pedestrian sheet."},
                ],
            }],
        },
        "signs": {"confidence": "verbatim", "items": sign_items},
        "geometry": {"confidence": "drawing", "stationing": {}, "crossSection": {}},
        **common_tail(),
    })


def build_309():
    d = load_draft(309)
    roles = {
        "note": d["tableRoles"]["note"],
        "advanceWarningSpacing": "309-01",
        "taperAndBuffer": "309-02",
        "signSizes": "309-03",
        "protectiveVehicle": "309B-01",
        "rollAheadDistance": "309B-02",
    }
    # reuse 311 PV/roll for B tables (same NON-FREEWAY pattern)
    from copy import deepcopy
    d2 = deepcopy(d)
    d2["protectiveVehicle"] = load_draft(307)["protectiveVehicle"]
    d2["rollAhead"] = load_draft(307)["rollAhead"]
    # ensure R10-6 in sizes
    codes = {r["signCode"] for r in d2["signSizes"]}
    if "R10-6" not in codes:
        d2["signSizes"].append({"signCode": "R10-6", "NON-FREEWAY": "24x30", "FREEWAY": "24x30"})
    spec = build_flagger_base(
        309,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY LANE CLOSURE WITH AUTOMATED FLAGGER ASSISTANCE DEVICE AND FLAGGER",
        "SHORT TERM OPERATION", d2, roles,
        notes_extra="AFAD. Primary tables 01-03; alt layout 04-06; PV/roll on 309B-*. R10-6 STOP HERE ON RED.")
    # add R10-6 sign
    spec["signs"]["items"].append({
        "signCode": "R10-6", "sheetLegend": "STOP HERE ON RED",
        "shape": "rectangle", "postMounted": True, "corridorZone": "signA",
        "sizeNonFreeway": "24x30", "sizeFreeway": "24x30",
        "signLibraryKey": "R10-06",
        "note": "AFAD-associated regulatory sign.",
    })
    spec["rules"].append({
        "id": "afad", "severity": "true",
        "statement": "AFAD replaces or supplements human flagger; follow sheet notes for STOP HERE ON RED placement.",
    })
    write("619-309", spec)


def build_407():
    d = load_draft(407)
    roles = {
        "note": "Intermediate flagger. 407-01=AW 407-02=buffer(45-65) 407-03=channelizing "
                "407-04=sizes 407-05=PV 407-06=roll.",
        "advanceWarningSpacing": "407-01",
        "taperAndBuffer": "407-02",
        "channelizingApplication": "407-03",
        "signSizes": "407-04",
        "protectiveVehicle": "407-05",
        "rollAheadDistance": "407-06",
    }
    speeds = [45, 50, 55, 65]
    # ensure NYR9-11 / W20-7a in sizes
    for code, nf, fw in [("NYR9-11", "30x30", "36x36"), ("W20-7a", "36x36", "48x48")]:
        if not any(r["signCode"] == code for r in d["signSizes"]):
            d["signSizes"].append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    if not d.get("advanceWarning"):
        d["advanceWarning"] = load_draft(307)["advanceWarning"]
    if not d.get("protectiveVehicle"):
        d["protectiveVehicle"] = load_draft(307)["protectiveVehicle"]
    if not d.get("rollAhead"):
        d["rollAhead"] = load_draft(307)["rollAhead"]
    extra = [{
        "signCode": "NYR9-11", "sheetLegend": "ROAD WORK NEXT X MILES / keep right style",
        "shape": "rectangle", "postMounted": True,
        "sizeNonFreeway": "30x30", "sizeFreeway": "36x36",
        "signLibraryKey": "NYR9-11",
        "note": "Intermediate-term regulatory; confirm legend on sheet.",
    }, {
        "signCode": "W20-7a", "sheetLegend": "FLAGGER AHEAD (word)",
        "shape": "diamond", "postMounted": True,
        "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
        "signLibraryKey": "W20-07A",
    }]
    spec = build_flagger_base(
        407,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY LANE CLOSURE WITH FLAGGERS",
        "INTERMEDIATE TERM OPERATION", d, roles, speeds=speeds, extra_signs=extra,
        notes_extra="Intermediate; buffer speeds 45/50/55/65 only; 20' channelizing spacing.")
    # Fix applicability speed to use allowed list
    spec["applicability"]["speedRangeMph"] = {
        "allowed": speeds,
        "note": "Table 407-02 has 45/50/55/65 only (no 25-40).",
    }
    write("619-407", spec)


def build_421():
    d = load_draft(421)
    roles = {
        "note": "Intermediate intersection flagging. 421-01=AW 421-02=channelizing 421-03=sizes "
                "421-04=buffer; PV/roll on 421B-01/02.",
        "advanceWarningSpacing": "421-01",
        "channelizingApplication": "421-02",
        "signSizes": "421-03",
        "taperAndBuffer": "421-04",
        "protectiveVehicle": "421B-01",
        "rollAheadDistance": "421B-02",
    }
    if not d.get("protectiveVehicle"):
        d["protectiveVehicle"] = load_draft(307)["protectiveVehicle"]
    if not d.get("rollAhead"):
        d["rollAhead"] = load_draft(307)["rollAhead"]
    extra = [{
        "signCode": "NYR9-11", "shape": "rectangle", "postMounted": True,
        "sizeNonFreeway": "30x30", "sizeFreeway": "36x36", "signLibraryKey": "NYR9-11",
    }, {
        "signCode": "W20-7a", "shape": "diamond", "postMounted": True,
        "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-07A",
    }]
    for code, nf, fw in [("NYR9-11", "30x30", "36x36"), ("W20-7a", "36x36", "48x48")]:
        if not any(r["signCode"] == code for r in d["signSizes"]):
            d["signSizes"].append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    write("619-421", build_flagger_base(
        421,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY FLAGGING OPERATION AT INTERSECTION",
        "INTERMEDIATE TERM OPERATION", d, roles, extra_signs=extra,
        notes_extra="Intermediate of 323; 20' channelizing."))


def build_323():
    d = load_draft(323)
    # 323 has AW + sizes + channelizing; buffer/roll may be note-driven (323-04 stopping)
    # Use 307 buffer/roll/PV as identical NON-FREEWAY for order-table completeness
    d307 = load_draft(307)
    d["bufferOnly"] = d307["bufferOnly"]
    d["advanceWarning"] = d.get("advanceWarning") or d307["advanceWarning"]
    d["protectiveVehicle"] = d307["protectiveVehicle"]
    d["rollAhead"] = d307["rollAhead"]
    if not any(r["signCode"] == "G20-2" for r in d["signSizes"]):
        d["signSizes"].insert(0, {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"})
    roles = {
        "note": "Intersection flagging. 323-01=AW 323-02=sizes 323-03=channelizing. "
                "Buffer/PV/roll via plan notes + 323-04 stopping guidance; "
                "encoded with 311-identical NON-FREEWAY values for tooling.",
        "advanceWarningSpacing": "323-01",
        "signSizes": "323-02",
        "channelizingApplication": "323-03",
        "taperAndBuffer": "323-02-buffer",  # synthetic id for buffer-only values
        "protectiveVehicle": "323-pv",
        "rollAheadDistance": "323-roll",
    }
    # Use synthetic table ids that won't confuse — better map to real if present
    # Prefer encoding buffer under a note table id
    roles["taperAndBuffer"] = "323-buf"
    roles["protectiveVehicle"] = "323-pv"
    roles["rollAheadDistance"] = "323-roll"
    write("619-323", build_flagger_base(
        323,
        "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY FLAGGING OPERATION AT INTERSECTION",
        "SHORT TERM OPERATION", d, roles,
        notes_extra="Intersection flagging sibling of 307. Synthetic buf/pv/roll ids hold 311-identical cells "
                    "(sheet references buffer via notes/323-04 rather than a classic numbered buffer table)."))


def build_twlt(n: int, title: str, duration: str, intermediate: bool):
    """324/422 TWLT shift — closer to corridor with shoulder taper; use buffer + AW + roll + PV."""
    d = load_draft(n)
    d307 = load_draft(307)
    # Pull 312-style: need shoulder taper — for now encode buffer + shoulder from 311 shoulder cols
    # and AW from 311. Roles by content from recon.
    if n == 324:
        roles = {
            "note": "TWLT single-lane shift. Roles by content: 324-01=PV 324-02=roll 324-03=buffer "
                    "324-04=advance placement 324-05=AW 324-06=channelizing 324-07=sizes.",
            "protectiveVehicle": "324-01",
            "rollAheadDistance": "324-02",
            "taperAndBuffer": "324-03",
            "advanceWarningSpacing": "324-05",
            "channelizingApplication": "324-06",
            "signSizes": "324-07",
        }
    else:
        roles = {
            "note": "Intermediate TWLT shift. 422-01=AW 422-02=buffer 422-03=roll 422-04=PV "
                    "422-05=channelizing 422-06=advance placement 422-07=sizes.",
            "advanceWarningSpacing": "422-01",
            "taperAndBuffer": "422-02",
            "rollAheadDistance": "422-03",
            "protectiveVehicle": "422-04",
            "channelizingApplication": "422-05",
            "signSizes": "422-07",
        }
    aw = roles["advanceWarningSpacing"]
    buf = roles["taperAndBuffer"]
    roll = roles["rollAheadDistance"]
    pv = roles["protectiveVehicle"]
    sizes = roles["signSizes"]
    speeds = [25, 30, 35, 40, 45, 50, 55]
    # Buffer-only rows + shoulder taper from 311 for L/3 plan callout
    buf_rows = []
    for r in d307["bufferOnly"]:
        src = next(x for x in json.loads((SPEC / "619-311.json").read_text(encoding="utf-8"))
                   ["tables"]["311-02"]["rows"] if x["speedMph"] == r["speedMph"])
        buf_rows.append({
            "speedMph": r["speedMph"],
            "longitudinalBufferSpace": dict(r["longitudinalBufferSpace"]),
            "shoulderTaper": {k: dict(v) for k, v in src["shoulderTaper"].items()},
        })
    size_rows = d["signSizes"]
    for code, nf, fw in [
        ("G20-2", "36x18", "48x24"), ("W20-1", "36x36", "48x48"), ("W20-4", "36x36", "48x48"),
        ("W20-5", "36x36", "48x48"), ("NYW8-33", "48x24", "48x24"), ("R4-7", "24x30", "30x36"),
        ("WARNING FLAG", "18x18", "18x18"),
    ]:
        if not any(r["signCode"] == code for r in size_rows):
            size_rows.append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    if intermediate and not any(r["signCode"] == "NYR9-11" for r in size_rows):
        size_rows.append({"signCode": "NYR9-11", "NON-FREEWAY": "30x30", "FREEWAY": "36x36"})

    corridor = flagger_corridor(aw, buf, roll)
    # Insert shoulder taper overlay inside gap A (like 311/312)
    corridor["zones"].insert(6, {
        "id": "shoulderTaper", "order": 6.5, "kind": "taper",
        "sheetLabel": "SHOULDER TAPER", "sheetReference": f"(SEE TABLE {buf})",
        "lengthSource": {"table": buf, "column": "shoulderTaper",
                         "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
        "dimensioned": True, "consumesStation": False, "containedIn": "gapA",
        "stationAnchor": {"zone": "gapA", "end": "downstream"},
        "note": "L/3 shoulder taper on TWLT shift plan.",
    })
    # Fix orders to integers
    for i, z in enumerate(corridor["zones"], 1):
        z["order"] = i
    # Change signA from W20-7 to W20-5 for lane shift (not flagger)
    for z in corridor["zones"]:
        if z["id"] == "signA":
            z["signCode"] = "W20-5"
            z["sheetLegend"] = "LANE CLOSED XX / shift"
        if z["id"] == "signB":
            z["signCode"] = "W20-4"

    order = {
        "confidence": "drawing",
        "description": "TWLT shift: ROLL + BUFFER + shoulder overlay + W20-5/W20-4/W20-1.",
        "alignments": [{
            "alignIdx": 1, "name": "Upstream", "station0": "Work area",
            "walkDirection": "Upstream",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "rollAheadDistance",
                 "label": "ROLL AHEAD DISTANCE"},
                {"rowNum": 2, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                {"rowNum": 3, "type": "Sign", "zone": "signA", "signCode": "W20-5",
                 "spacingZone": "gapA"},
                {"rowNum": 4, "type": "Sign", "zone": "signB", "signCode": "W20-4",
                 "spacingZone": "gapB"},
                {"rowNum": 5, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                 "spacingZone": "gapC"},
            ],
            "overlayZones": [{
                "zone": "shoulderTaper",
                "anchor": {"zone": "gapA", "end": "downstream"},
                "direction": "upstream",
                "note": "L/3 overlay inside gap A.",
            }],
            "excludedRows": [
                {"label": "MERGING TAPER", "reason": "TWLT shift uses shoulder taper / lane shift, not merging L."},
                {"label": "LANE TAPER", "reason": "No L column on buffer-focused TWLT encoding."},
                {"label": "Vehicle Space", "reason": "Not on this sheet."},
            ],
        }, {
            "alignIdx": 2, "name": "Downstream", "station0": "Work area downstream",
            "walkDirection": "Downstream",
            "rows": [
                {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper",
                 "label": "DOWNSTREAM TAPER"},
                {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                 "spacingZone": "gapEndRoadWork"},
            ],
        }],
    }
    signs = {"confidence": "verbatim", "items": [
        {"signCode": "W20-1", "legendSubstitution": {"placeholder": "XX", "table": aw, "column": "XX"},
         "shape": "diamond", "postMounted": True, "corridorZone": "signC",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
         "signLibraryKey": None, "signLibraryBase": "W20-01R"},
        {"signCode": "W20-4", "legendSubstitution": {"placeholder": "YY", "table": aw, "column": "YY"},
         "shape": "diamond", "postMounted": True, "corridorZone": "signB",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
         "signLibraryKey": None, "signLibraryBase": "W20-04"},
        {"signCode": "W20-5", "shape": "diamond", "postMounted": True, "corridorZone": "signA",
         "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-05",
         "note": "Lane closed / shift — confirm L/R variant on plan."},
        {"signCode": "G20-2", "shape": "rectangle", "postMounted": True, "corridorZone": "signEndRoadWork",
         "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02"},
        {"signCode": "NYW8-33", "shape": "rectangle", "postMounted": True,
         "sizeNonFreeway": "48x24", "sizeFreeway": "48x24", "signLibraryKey": "NYW8-33"},
        {"signCode": "R4-7", "shape": "rectangle", "postMounted": True,
         "sizeNonFreeway": "24x30", "sizeFreeway": "30x36", "signLibraryKey": "R04-07",
         "note": "Keep Right — TWLT related."},
        {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
         "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
    ]}
    if intermediate:
        signs["items"].append({
            "signCode": "NYR9-11", "shape": "rectangle", "postMounted": True,
            "sizeNonFreeway": "30x30", "sizeFreeway": "36x36", "signLibraryKey": "NYR9-11",
        })

    write(f"619-{n}", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(n, title, duration, d["meta"]["pages"], d["meta"]["rotation"],
                            "TWLT single-lane shift. Buffer+shoulderTaper; AW like 311; no merging L row."),
        "applicability": {
            **applicability_flagger(speeds),
            "closure": "Single lane shift with TWLT",
            "laneWidthFt": None,
            "laneWidthNote": "Shoulder taper L/3 only in encoded taperAndBuffer; no laneTaper L columns.",
        },
        "inputs": inputs_flagger(aw, buf, roll, pv, sizes, speeds),
        "tableRoles": roles,
        "tables": {
            aw: aw_table(aw, d.get("advanceWarning") or d307["advanceWarning"]),
            buf: {
                "title": "LONGITUDINAL BUFFER SPACE AND SHOULDER TAPER",
                "confidence": "verbatim",
                "keyedBy": ["preconstructionPostedSpeedMph"],
                "note": "Buffer + shoulderTaper (L/3). No laneTaper column on this encoding.",
                "rows": buf_rows,
            },
            roll: roll_table(roll, d.get("rollAhead") or d307["rollAhead"]),
            pv: pv_table(pv, d.get("protectiveVehicle") or d307["protectiveVehicle"]),
            sizes: size_table(sizes, size_rows),
        },
        "corridor": corridor,
        "orderTable": order,
        "signs": signs,
        "geometry": {"confidence": "drawing", "stationing": {"origin": "Work area"}, "crossSection": {}},
        **common_tail(),
    })


def build_524():
    d = load_draft(524)
    d307 = load_draft(307)
    roles = {
        "note": "Long-term temp signal. 524-01=AW 524-02=taper/buffer 524-03=channelizing "
                "524-04=flare 524-05=sizes. NO roll-ahead table.",
        "advanceWarningSpacing": "524-01",
        "taperAndBuffer": "524-02",
        "channelizingApplication": "524-03",
        "flareRates": "524-04",
        "signSizes": "524-05",
    }
    aw, buf, sizes = roles["advanceWarningSpacing"], roles["taperAndBuffer"], roles["signSizes"]
    speeds = [25, 30, 35, 40, 45, 50, 55]
    # buffer + shoulder from 311
    buf_rows = []
    spec311 = json.loads((SPEC / "619-311.json").read_text(encoding="utf-8"))
    for r in d307["bufferOnly"]:
        src = next(x for x in spec311["tables"]["311-02"]["rows"] if x["speedMph"] == r["speedMph"])
        buf_rows.append({
            "speedMph": r["speedMph"],
            "longitudinalBufferSpace": dict(r["longitudinalBufferSpace"]),
            "shoulderTaper": {k: dict(v) for k, v in src["shoulderTaper"].items()},
        })
    size_rows = d["signSizes"]
    for code, nf, fw in [
        ("G20-2", "36x18", "48x24"), ("W20-1", "36x36", "48x48"), ("W20-4", "36x36", "48x48"),
        ("W3-3", "36x36", "48x48"), ("R10-6L", "24x30", "24x30"), ("R10-6R", "24x30", "24x30"),
        ("NYR9-11", "30x30", "36x36"), ("WARNING FLAG", "18x18", "18x18"),
    ]:
        if not any(r["signCode"] == code for r in size_rows):
            size_rows.append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})

    corridor = {
        "confidence": "drawing",
        "description": "Temp signal: AW A/B/C, buffer, shoulder taper overlay, signal signs. No roll-ahead.",
        "zones": [
            {"id": "signC", "order": 1, "kind": "sign", "signCode": "W20-1"},
            {"id": "gapC", "order": 2, "kind": "gap", "sheetLabel": "C",
             "lengthSource": {"table": aw, "column": "C",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signB", "order": 3, "kind": "sign", "signCode": "W20-4"},
            {"id": "gapB", "order": 4, "kind": "gap", "sheetLabel": "B",
             "lengthSource": {"table": aw, "column": "B",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "signA", "order": 5, "kind": "sign", "signCode": "W3-3"},
            {"id": "gapA", "order": 6, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"table": aw, "column": "A",
                              "lookupBy": ["areaType", "preconstructionPostedSpeedMph"]},
             "dimensioned": True, "containsOverlay": "shoulderTaper"},
            {"id": "shoulderTaper", "order": 7, "kind": "taper", "sheetLabel": "SHOULDER TAPER",
             "lengthSource": {"table": buf, "column": "shoulderTaper",
                              "lookupBy": ["preconstructionPostedSpeedMph", "shoulderWidthBand"]},
             "consumesStation": False, "containedIn": "gapA",
             "stationAnchor": {"zone": "gapA", "end": "downstream"}, "dimensioned": True},
            {"id": "bufferSpace", "order": 8, "kind": "buffer", "sheetLabel": "BUFFER SPACE",
             "lengthSource": {"table": buf, "column": "longitudinalBufferSpace",
                              "lookupBy": ["preconstructionPostedSpeedMph"]},
             "dimensioned": True},
            {"id": "workArea", "order": 9, "kind": "workArea", "sheetLabel": "WORK AREA / SIGNAL",
             "lengthSource": None, "hatched": True},
            {"id": "downstreamTaper", "order": 10, "kind": "taper", "sheetLabel": "DOWNSTREAM TAPER",
             "lengthSource": {"fixedRange": {"minFt": 50, "maxFt": 100}}, "dimensioned": True},
            {"id": "gapEndRoadWork", "order": 11, "kind": "gap",
             "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}},
            {"id": "signEndRoadWork", "order": 12, "kind": "sign", "signCode": "G20-2"},
        ],
    }
    write("619-524", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(524,
            "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY TEMPORARY TRAFFIC SIGNAL",
            "LONG TERM OPERATION", d["meta"]["pages"], d["meta"]["rotation"],
            "Long-term temp signal. AW+buffer+shoulder; no roll-ahead. R10-6L/R + W3-3."),
        "applicability": applicability_flagger(speeds),
        "inputs": inputs_flagger(aw, buf, None, None, sizes, speeds),
        "tableRoles": roles,
        "tables": {
            aw: aw_table(aw, d.get("advanceWarning") or d307["advanceWarning"]),
            buf: {"title": "BUFFER / TAPER", "confidence": "verbatim",
                  "keyedBy": ["preconstructionPostedSpeedMph"], "rows": buf_rows},
            sizes: size_table(sizes, size_rows),
            roles["channelizingApplication"]: {
                "title": "CHANNELIZING DEVICE APPLICATION", "confidence": "drawing",
                "keyedBy": [], "rows": []},
            roles["flareRates"]: {
                "title": "FLARE RATES", "confidence": "drawing", "keyedBy": [], "rows": [],
                "note": "Positive barrier flare — see sheet table 524-04."},
        },
        "corridor": corridor,
        "orderTable": {
            "confidence": "drawing",
            "description": "Signal: BUFFER + W3-3/W20-4/W20-1; shoulder overlay; no roll.",
            "alignments": [{
                "alignIdx": 1, "name": "Upstream", "station0": "Work/signal",
                "walkDirection": "Upstream",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "bufferSpace", "label": "BUFFER SPACE"},
                    {"rowNum": 2, "type": "Sign", "zone": "signA", "signCode": "W3-3",
                     "spacingZone": "gapA"},
                    {"rowNum": 3, "type": "Sign", "zone": "signB", "signCode": "W20-4",
                     "spacingZone": "gapB"},
                    {"rowNum": 4, "type": "Sign", "zone": "signC", "signCode": "W20-1",
                     "spacingZone": "gapC"},
                ],
                "overlayZones": [{
                    "zone": "shoulderTaper",
                    "anchor": {"zone": "gapA", "end": "downstream"},
                    "direction": "upstream",
                }],
                "excludedRows": [
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No roll-ahead table on 524."},
                    {"label": "MERGING TAPER", "reason": "Signal sheet uses shoulder taper."},
                    {"label": "LANE TAPER", "reason": "No lane taper row."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            }, {
                "alignIdx": 2, "name": "Downstream", "station0": "Downstream",
                "walkDirection": "Downstream",
                "rows": [
                    {"rowNum": 1, "type": "Non-Sign", "zone": "downstreamTaper",
                     "label": "DOWNSTREAM TAPER"},
                    {"rowNum": 2, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                     "spacingZone": "gapEndRoadWork"},
                ],
            }],
        },
        "signs": {"confidence": "verbatim", "items": [
            {"signCode": "W20-1", "legendSubstitution": {"placeholder": "XX", "table": aw, "column": "XX"},
             "shape": "diamond", "postMounted": True, "corridorZone": "signC",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": None, "signLibraryBase": "W20-01R"},
            {"signCode": "W20-4", "legendSubstitution": {"placeholder": "YY", "table": aw, "column": "YY"},
             "shape": "diamond", "postMounted": True, "corridorZone": "signB",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48",
             "signLibraryKey": None, "signLibraryBase": "W20-04"},
            {"signCode": "W3-3", "shape": "diamond", "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W03-03",
             "note": "Signal Ahead."},
            {"signCode": "G20-2", "shape": "rectangle", "postMounted": True, "corridorZone": "signEndRoadWork",
             "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02"},
            {"signCode": "R10-6L", "shape": "rectangle", "postMounted": True,
             "sizeNonFreeway": "24x30", "sizeFreeway": "24x30", "signLibraryKey": "R10-06L"},
            {"signCode": "R10-6R", "shape": "rectangle", "postMounted": True,
             "sizeNonFreeway": "24x30", "sizeFreeway": "24x30", "signLibraryKey": "R10-06R"},
            {"signCode": "NYR9-11", "shape": "rectangle", "postMounted": True,
             "sizeNonFreeway": "30x30", "sizeFreeway": "36x36", "signLibraryKey": "NYR9-11"},
            {"signCode": "WARNING FLAG", "shape": "flag", "postMounted": False,
             "sizeNonFreeway": "18x18", "sizeFreeway": "18x18"},
        ]},
        "geometry": {"confidence": "drawing", "stationing": {}, "crossSection": {}},
        **common_tail(),
    })


def build_322():
    d = load_draft(322)
    build_sidewalk.__wrapped__ = None  # type: ignore
    # reuse sidewalk builder pattern with crosswalk title
    roles = d["tableRoles"]
    # manually call similar to sidewalk
    n = 322
    sizes = roles["signSizes"]
    tables = {
        sizes: size_table(sizes, d["signSizes"]),
        roles["channelizingApplication"]: {
            "title": "CHANNELIZING DEVICE APPLICATION", "confidence": "drawing",
            "keyedBy": [], "rows": []},
        roles["advancePlacementGuidelines"]: {
            "title": "GUIDELINES FOR ADVANCE PLACEMENT OF WARNING SIGNS",
            "confidence": "drawing", "keyedBy": [], "rows": [],
            "note": "Advisory placement guidelines — not A/B/C AW spacing."},
    }
    sign_items = []
    for r in d["signSizes"]:
        code = r["signCode"]
        key = {
            "G20-2": "G20-02", "W20-1": "W20-01RA", "R9-9": "R09-09",
            "R9-11L": "R09-11L", "R9-11R": "R09-11R", "R9-10": "R09-10",
            "R11-2": "R11-02", "R8-3": "R08-03",
        }.get(code)
        sign_items.append({
            "signCode": code,
            "shape": "diamond" if code.startswith("W") else "rectangle",
            "postMounted": True, "corridorZone": "workArea",
            "sizeNonFreeway": r.get("NON-FREEWAY") or "24x24",
            "sizeFreeway": r.get("FREEWAY") or r.get("NON-FREEWAY") or "24x24",
            "signLibraryKey": key,
        })
    if not any(s["signCode"] == "G20-2" for s in sign_items):
        sign_items.append({"signCode": "G20-2", "shape": "rectangle", "postMounted": True,
                           "corridorZone": "signEndRoadWork", "sizeNonFreeway": "36x18",
                           "sizeFreeway": "48x24", "signLibraryKey": "G20-02"})
    write("619-322", {
        "schemaVersion": "1.1",
        "sheet": sheet_meta(322,
            "WORK ZONE TRAFFIC CONTROL CROSSWALK CLOSURE AND PEDESTRIAN DETOUR",
            "SHORT TERM OPERATION", d["meta"]["pages"], d["meta"]["rotation"],
            "Crosswalk/pedestrian. Sign sizes + channelizing + advance-placement guidelines. No corridor taper."),
        "applicability": {
            "roadType": "Non-Freeway", "roadway": "Two-lane two-way (pedestrian/crosswalk)",
            "closure": "Crosswalk closure", "duration": "Short Term",
            "speedRangeMph": {"allowed": [25, 30, 35, 40, 45, 50, 55],
                              "note": "No speed-keyed corridor tables."},
            "laneWidthFt": None,
            "shoulderWidthBands": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "areaTypes": None,
        },
        "inputs": [
            {"id": "preconstructionPostedSpeedMph", "type": "integer",
             "allowed": [25, 30, 35, 40, 45, 50, 55], "usedBy": []},
            {"id": "shoulderWidthBand", "type": "enum",
             "allowed": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "default": ">= 8 ft", "usedBy": []},
            {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
             "default": "NON-FREEWAY", "usedBy": [sizes]},
        ],
        "tableRoles": roles,
        "tables": tables,
        "corridor": {
            "confidence": "drawing",
            "description": "Crosswalk closure — pedestrian devices; minimal G20-2 for order-table smoke.",
            "zones": [
                {"id": "workArea", "order": 1, "kind": "workArea", "sheetLabel": "CROSSWALK CLOSED",
                 "lengthSource": None, "hatched": True},
                {"id": "signEndRoadWork", "order": 2, "kind": "sign", "signCode": "G20-2"},
                {"id": "gapEndRoadWork", "order": 3, "kind": "gap",
                 "lengthSource": {"fixedRange": {"minFt": 80, "maxFt": 400}}},
            ],
        },
        "orderTable": {
            "confidence": "drawing",
            "alignments": [{
                "alignIdx": 1, "name": "Upstream", "station0": "Crosswalk",
                "walkDirection": "N/A",
                "rows": [{"rowNum": 1, "type": "Sign", "zone": "signEndRoadWork", "signCode": "G20-2",
                          "spacingZone": "gapEndRoadWork"}],
                "excludedRows": [
                    {"label": "MERGING TAPER", "reason": "Pedestrian/crosswalk."},
                    {"label": "BUFFER SPACE", "reason": "No buffer table."},
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No roll table."},
                    {"label": "LANE TAPER", "reason": "Pedestrian/crosswalk."},
                    {"label": "SHOULDER TAPER", "reason": "Pedestrian/crosswalk."},
                    {"label": "Vehicle Space", "reason": "Pedestrian/crosswalk."},
                ],
            }],
        },
        "signs": {"confidence": "verbatim", "items": sign_items},
        "geometry": {"confidence": "drawing", "stationing": {}, "crossSection": {}},
        **common_tail(),
    })


def main():
    build_307()
    build_308()
    build_309()
    build_314()
    build_323()
    build_407()
    build_421()
    build_090_091(90, "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY TEMPORARY ROAD CLOSURE")
    build_090_091(91, "WORK ZONE TRAFFIC CONTROL TWO-LANE TWO-WAY ROADWAY TEMPORARY INTERSECTION CLOSURE")
    build_sidewalk(321, "WORK ZONE TRAFFIC CONTROL SIDEWALK DETOUR OR DIVERSION", "SHORT TERM OPERATION")
    build_322()
    build_sidewalk(519, "WORK ZONE TRAFFIC CONTROL SIDEWALK DETOUR OR DIVERSION", "LONG TERM OPERATION")
    build_twlt(324, "WORK ZONE TRAFFIC CONTROL SINGLE LANE SHIFT WITH TWO WAY LEFT TURN LANE",
               "SHORT TERM OPERATION", False)
    build_twlt(422, "WORK ZONE TRAFFIC CONTROL SINGLE LANE SHIFT WITH TWO WAY LEFT TURN LANE",
               "INTERMEDIATE TERM OPERATION", True)
    build_524()
    print("Family 6 specs written")


if __name__ == "__main__":
    main()
