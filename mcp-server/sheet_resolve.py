"""Load and resolve Data/sheet-specs/<sheet>.json.

This is the Python-side owner of standard-sheet knowledge. It turns a sheet
spec plus the plan's inputs (speed, lane width, shoulder band, area type)
into the concrete station rows the VBA order table needs.

Why here and not in VBA: the spec is a static JSON data file. Routing a file
read through the MicroStation bridge (as get_sheet_requirements does for
sheet-registry.tsv) buys nothing, and VBA has no usable JSON parser. VBA keeps
receiving resolved numbers as bridge parameters -- the same shape it already
receives signRowsTSV in.

Engineering-judgment boundary (see wztc_ops.py): this module does not invent
numbers. Every value it returns is either read straight from a sheet table or
is a documented lookup on one. If a sheet has no spec, callers must fall back
to WZTCRules and say so, not guess.

Part of the sheet_spec split (2026-08-04): this module owns "what does this
sheet need" -- spec loading, table lookups (resolve), and the order-table/
station-walk payloads built from a resolved result. sheet_compile.py turns
those into absolute-coordinate drawing primitives; sheet_rules.py validates
compiled primitives before they reach the bridge. sheet_spec.py re-exports
all three so every existing `sheet_spec.X` call site keeps working unchanged.
"""
from __future__ import annotations

import json
import pathlib
import re
from typing import Optional

from pydantic import ValidationError

import sheet_schema

SPEC_DIR = pathlib.Path(__file__).resolve().parent.parent / "Data" / "sheet-specs"

# Table 311-03 style legend placeholders map onto SignLibrary's message-variant
# suffixes. The sheet is what disambiguates these -- resolve_sign_code cannot,
# which is why it returns every variant as a 'candidate'.
_LEGEND_SUFFIX = {"AHEAD": "A", "FT": "F", "FEET": "F", "MILE": "M", "MI": "M"}


class SpecError(Exception):
    pass


def _band_bounds(label: str) -> tuple[float, float]:
    """Parse a printed band label into numeric bounds.

    '<= 4 ft' -> (0, 4)   '5 - 7 ft' -> (5, 7)   '>= 8 ft' -> (8, inf)
    """
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", label)]
    if not nums:
        raise SpecError(f"cannot parse band label {label!r}")
    if "<=" in label or "<" in label:
        return (0.0, nums[0])
    if ">=" in label or ">" in label:
        return (nums[0], float("inf"))
    if len(nums) >= 2:
        return (nums[0], nums[1])
    return (nums[0], nums[0])


def allowed_speeds(spec: dict) -> list[int]:
    """A sheet's speedRangeMph is either a uniform {min,max,increment} range
    (619-311: every 5mph step 25-55 exists) or an explicit 'allowed' list for
    a sheet with a genuine gap (619-302's Table 302-02 skips 60mph even though
    619-011's master taper table has it -- confirmed by direct extraction,
    not assumed). Prefer 'allowed' when present so a gap is never silently
    papered over by a min/max/increment range that implies it doesn't exist."""
    rng = spec["applicability"]["speedRangeMph"]
    if "allowed" in rng:
        return list(rng["allowed"])
    return list(range(rng["min"], rng["max"] + 1, rng["increment"]))


def shoulder_band(spec: dict, shoulder_key: str) -> str:
    """Map the app's shoulder dropdown value onto one of the sheet's printed
    bands.

    The app offers per-foot choices ('8 ft'...'12 ft'); the sheet prints three
    bands and never distinguishes above 8 ft. ComputeSpacing invents per-foot
    values in that range -- at 45 mph / 12 ft shoulder it returns 160 ft where
    Table 311-02 says 120 ft. Collapsing to the printed band here is what stops
    that fabricated number reaching the drawing.
    """
    bands = spec["applicability"].get("shoulderWidthBands") or []
    if not bands:
        return shoulder_key
    if shoulder_key in bands:
        return shoulder_key
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", shoulder_key)]
    if not nums:
        raise SpecError(f"cannot interpret shoulder width {shoulder_key!r}; "
                        f"sheet bands are {bands}")
    width = nums[0] if len(nums) == 1 else sum(nums[:2]) / 2
    for b in bands:
        lo, hi = _band_bounds(b)
        if lo <= width <= hi:
            return b
    raise SpecError(f"shoulder width {shoulder_key!r} falls outside {bands}")


def spec_path(sheet_num: str) -> pathlib.Path:
    return SPEC_DIR / f"{sheet_num}.json"


def has_spec(sheet_num: str) -> bool:
    return spec_path(sheet_num).is_file()


def build_guide_path(sheet_num: str, spec: Optional[dict] = None) -> Optional[pathlib.Path]:
    """Resolve Data/sheet-specs/<sheet>.build.md (or sheet.buildGuide override).

    Machine prefs stay in the JSON; the companion .build.md holds live-build
    tips/QA/gotchas so the next session does not re-learn them from agent-log.
    """
    name = None
    if spec is None and has_spec(sheet_num):
        try:
            spec = load(sheet_num)
        except SpecError:
            spec = None
    if isinstance(spec, dict):
        sheet_meta = spec.get("sheet") or {}
        if isinstance(sheet_meta, dict):
            name = sheet_meta.get("buildGuide") or None
    if not name:
        name = f"{sheet_num}.build.md"
    # Only allow basename under SPEC_DIR (no path traversal).
    base = pathlib.Path(str(name)).name
    if not base.lower().endswith(".md"):
        return None
    path = SPEC_DIR / base
    return path if path.is_file() else None


def load_build_guide(sheet_num: str) -> Optional[dict]:
    """Load the sheet build playbook markdown, if present.

    Returns dict with path (repo-relative), absolutePath, sheetNum, text,
    charCount — or None when no companion guide exists.
    """
    path = build_guide_path(sheet_num)
    if path is None:
        return None
    text = path.read_text(encoding="utf-8")
    try:
        rel = str(path.relative_to(SPEC_DIR.parent.parent))
    except ValueError:
        rel = str(path)
    return {
        "sheetNum": sheet_num,
        "path": rel.replace("\\", "/"),
        "absolutePath": str(path),
        "charCount": len(text),
        "text": text,
    }


def load(sheet_num: str) -> Optional[dict]:
    """Load + Pydantic-validate a sheet spec. Returns the original dict on
    success (callers keep working against plain dicts); raises SpecError
    when the JSON exists but fails the schema gate."""
    p = spec_path(sheet_num)
    if not p.is_file():
        return None
    raw = json.loads(p.read_text(encoding="utf-8"))
    try:
        sheet_schema.validate_sheet_dict(raw)
    except ValidationError as e:
        raise SpecError(
            f"sheet {sheet_num} failed Pydantic schema: "
            f"{sheet_schema.format_validation_error(e)}"
        ) from e
    return raw


def load_raw_path(path: pathlib.Path) -> dict:
    """Load a spec from an explicit path (validator script). Same schema
    gate as load()."""
    raw = json.loads(path.read_text(encoding="utf-8"))
    try:
        sheet_schema.validate_sheet_dict(raw)
    except ValidationError as e:
        raise SpecError(
            f"{path.name} failed Pydantic schema: "
            f"{sheet_schema.format_validation_error(e)}"
        ) from e
    return raw


def _band(row: dict, speed: int) -> bool:
    lo, hi = row.get("minMph"), row.get("maxMph")
    return (lo is None or speed >= lo) and (hi is None or speed <= hi)


def _matches(sheet_value: str, given: str) -> bool:
    """Table keys are the sheet's verbatim wording ('LANE CLOSURE OR
    ENCROACHMENT'). Accept a caller's shorter form as long as it identifies
    exactly one row -- the caller checks the match count."""
    a, b = sheet_value.upper(), given.strip().upper()
    return a == b or b in a


def legend_suffix(legend: str) -> str:
    """'1000 FT.' -> 'F', 'AHEAD' -> 'A'. Raises rather than defaulting to
    'Ahead', because a silent wrong variant is exactly the failure the sheet
    exists to prevent."""
    for token in legend.upper().replace(".", "").split():
        if token in _LEGEND_SUFFIX:
            return _LEGEND_SUFFIX[token]
    raise SpecError(f"cannot map sign legend {legend!r} to a SignLibrary variant suffix")


def resolve(spec: dict, speed: int, lane_width: int, shoulder_width: str,
            area_type: Optional[str] = None, closure_type: Optional[str] = None,
            exposure: Optional[str] = None,
            protective_vehicle_gvw: Optional[int] = None) -> dict:
    """Resolve every table lookup this sheet needs for one set of inputs.

    shoulder_width accepts either a printed band ('>= 8 ft') or the app's
    per-foot dropdown value ('12 ft'), which is collapsed onto a band.

    area_type is required only when the sheet has an advanceWarningSpacing table.
    protective_vehicle_gvw is used when rollAheadDistance rows are keyed by GVW
    (619-301) rather than posted speed (619-302).
    """
    t = spec["tables"]
    roles = spec["tableRoles"]
    allowed = allowed_speeds(spec)
    if speed not in allowed:
        raise SpecError(
            f"sheet {spec['sheet']['number']} covers {allowed} mph only; got {speed}. "
            f"{spec['applicability']['speedRangeMph'].get('note', '')}")
    band = shoulder_band(spec, shoulder_width)
    lane_allowed = spec["applicability"].get("laneWidthFt") or []
    if lane_allowed and lane_width not in lane_allowed:
        raise SpecError(f"lane width {lane_width} not in {lane_allowed}")

    r02 = None
    if "taperAndBuffer" in roles and roles["taperAndBuffer"] in t:
        r02 = next(r for r in t[roles["taperAndBuffer"]]["rows"] if r["speedMph"] == speed)
        out = {
            "bufferFt": r02["longitudinalBufferSpace"]["ft"],
            "shoulderBand": band,
        }
        if "shoulderTaper" in r02 and band in r02["shoulderTaper"]:
            out["shoulderTaper"] = r02["shoulderTaper"][band]
        if "laneTaper" in r02 and lane_width is not None:
            out["laneTaper"] = r02["laneTaper"][str(lane_width)]
        # Optional lateral-shift column (ramp sheets 315/415)
        if "lateralShiftTaper" in r02 and lane_width is not None:
            out["lateralShiftTaper"] = r02["lateralShiftTaper"][str(lane_width)]
    else:
        # Short-duration mobile sheets (e.g. 619-205) have no taper/buffer table.
        out = {"shoulderBand": band}

    aw_role = roles.get("advanceWarningSpacing")
    if aw_role and aw_role in t:
        if not area_type:
            raise SpecError(
                f"sheet {spec['sheet']['number']} needs area_type/roadType for "
                f"advance-warning table {aw_role}")
        aw_table = t[aw_role]
        key_field = "areaType" if "areaType" in aw_table["rows"][0] else "roadType"
        r03 = next(r for r in aw_table["rows"]
                   if r[key_field] == area_type and _band(r, speed))
        # Short-duration / TWLT sheets may print only A/B (two advance signs).
        out["signGapFt"] = {k: r03[k] for k in ("A", "B", "C") if k in r03}
        out["legend"] = {k: r03[k] for k in ("XX", "YY") if k in r03}
    else:
        out["signGapFt"] = {}
        out["legend"] = {}

    if "rollAheadDistance" in roles and roles["rollAheadDistance"] in t:
        rad_rows = t[roles["rollAheadDistance"]]["rows"]
        if rad_rows and ("minGvwLbs" in rad_rows[0] or "gvwBand" in rad_rows[0]):
            gvw = protective_vehicle_gvw if protective_vehicle_gvw is not None else 22000
            hits = [r for r in rad_rows
                    if gvw >= r["minGvwLbs"]
                    and (r.get("maxGvwLbs") is None or gvw <= r["maxGvwLbs"])]
            if len(hits) != 1:
                raise SpecError(
                    f"roll-ahead GVW lookup for {gvw} lbs matched {len(hits)} rows")
            r04 = hits[0]
        else:
            r04 = next(r for r in rad_rows if _band(r, speed))
        out["rollAheadFt"] = {"min": r04["min"]["ft"], "max": r04["max"]["ft"]}

    if closure_type and exposure:
        if "protectiveVehicle" not in roles or roles["protectiveVehicle"] not in t:
            raise SpecError(
                f"sheet {spec['sheet']['number']} has no protectiveVehicle table "
                f"(long-term/barrier sheets); do not pass closure/exposure")
        pv_table = t[roles["protectiveVehicle"]]
        rows = pv_table["rows"]
        hits = [r for r in rows
                if _matches(r["closureType"], closure_type)
                and _matches(r["exposureCondition"], exposure)]
        if len(hits) != 1:
            raise SpecError(
                f"protective vehicle lookup for closure={closure_type!r} "
                f"exposure={exposure!r} matched {len(hits)} rows of table "
                f"{roles['protectiveVehicle']}. "
                f"Closure types: {sorted({r['closureType'] for r in rows})}. "
                f"Exposure conditions: {sorted({r['exposureCondition'] for r in rows})}")
        if pv_table.get("speedBands"):
            pv_band = next(b["id"] for b in pv_table["speedBands"] if _band(b, speed))
            out["protectiveVehicle"] = hits[0][pv_band]
            out["protectiveVehicleBand"] = pv_band
        elif hits[0].get("FREEWAY"):
            # Short-duration sheets may only print a FREEWAY column.
            out["protectiveVehicle"] = hits[0]["FREEWAY"]
            out["protectiveVehicleBand"] = "FREEWAY"
        else:
            raise SpecError(
                f"protective vehicle row has no speedBands and no FREEWAY cell")
    return out


def sign_library_key(item: dict, resolved: dict) -> str:
    """SignLibrary.bas key for one spec sign item."""
    if item.get("signLibraryKey"):
        return item["signLibraryKey"]
    base = item.get("signLibraryBase")
    sub = item.get("legendSubstitution")
    if not base or not sub:
        raise SpecError(
            f"sign {item['signCode']}: no signLibraryKey and no "
            f"signLibraryBase/legendSubstitution to derive one")
    return base + legend_suffix(resolved["legend"][sub["placeholder"]])


def zone_length(spec: dict, zone_id: str, resolved: dict,
                range_pick: str = "min") -> float:
    """Length in feet for a corridor zone. Ranges resolve by range_pick.

    Optional lengthSource.scale multiplies the looked-up value (619-303's
    '2L' span between successive merging tapers is scale=2 on laneTaper).
    """
    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    z = zones[zone_id]
    ls = z.get("lengthSource")
    if not ls:
        raise SpecError(f"zone {zone_id} has no length")
    scale = float(ls.get("scale", 1))
    if "fixedFt" in ls:
        return float(ls["fixedFt"]) * scale
    if "fixedRange" in ls:
        fr = ls["fixedRange"]
        return float(fr["minFt"] if range_pick == "min" else fr["maxFt"]) * scale
    tbl = ls["table"]
    roles = spec["tableRoles"]
    col = ls.get("column")
    if tbl == roles.get("rollAheadDistance"):
        return float(resolved["rollAheadFt"][range_pick]) * scale
    if tbl == roles.get("advanceWarningSpacing"):
        return float(resolved["signGapFt"][z["sheetLabel"]]) * scale
    if zone_id == "bufferSpace" or col == "longitudinalBufferSpace":
        return float(resolved["bufferFt"]) * scale
    if (zone_id in ("laneTaper", "mergingTaperUpstream", "mergingTaperDownstream")
            or col == "laneTaper"):
        if "laneTaper" not in resolved:
            raise SpecError(f"zone {zone_id} needs laneTaper but sheet has none")
        return float(resolved["laneTaper"]["ft"]) * scale
    if zone_id == "shoulderTaper" or col == "shoulderTaper":
        return float(resolved["shoulderTaper"]["ft"]) * scale
    raise SpecError(f"no resolver for zone {zone_id} (table {tbl}, column {col})")


def canonical_order_label(label: str) -> str:
    """Map sheet-spec Non-Sign labels onto the exact strings WZTCRules /
    PerpPlacement expect (title case from GetDefaultUpstreamItems).

    Specs often print ALL CAPS plan callouts ('SHOULDER TAPER', 'LANE TAPER').
    VBA Select Case on those strings is case-sensitive and misses them, so
    place_order_table_labels returns 0 rows; channelizing also looks for
    'Merging/Shifting Taper' not 'LANE TAPER'. Normalize at emit time.
    """
    key = " ".join((label or "").strip().upper().split())
    aliases = {
        "ROLL AHEAD DISTANCE": "Roll Ahead Distance",
        "ROLL AHEAD": "Roll Ahead Distance",
        "VEHICLE SPACE": "Vehicle Space",
        "BUFFER SPACE": "Buffer Space",
        "BUFFER": "Buffer Space",
        "SHOULDER TAPER": "Shoulder Taper",
        "DOWNSTREAM TAPER": "Downstream Taper",
        "WORK AREA": "Work Area",
        "MERGING/SHIFTING TAPER": "Merging/Shifting Taper",
        "MERGING TAPER": "Merging/Shifting Taper",
        "LANE TAPER": "Merging/Shifting Taper",
        "SHIFTING TAPER": "Merging/Shifting Taper",
        "UPSTREAM TAPER TEMP BARRIER": "Upstream Taper Temp Barrier",
        "UPSTREAM TAPER BOX/CORR BEAM": "Upstream Taper Box/Corr Beam",
    }
    return aliases.get(key, label)


def order_table_rows(spec: dict, resolved: dict, side: str = "One Side",
                     size_class: str = "NON-FREEWAY",
                     range_pick: str = "min") -> dict:
    """Build the bridge payload for BUILD_WZTC_ORDER_TABLE.

    Returns non-sign rows as 'alignIdx:label:spacing' and sign rows as
    'alignIdx:signKey:side:spacing:size', both in walk order. Overlay zones
    (e.g. 619-311's shoulder taper, which lies inside gap A) are returned
    separately -- they are drawn but consume no station.
    """
    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    signs = {s["signCode"]: s for s in spec["signs"]["items"]}
    size_key = "sizeNonFreeway" if size_class == "NON-FREEWAY" else "sizeFreeway"

    non_sign, sign, overlays = [], [], []
    for al in spec["orderTable"]["alignments"]:
        a = al["alignIdx"]
        for r in al["rows"]:
            if r["type"] == "Sign":
                item = signs[r["signCode"]]
                key = sign_library_key(item, resolved)
                spacing = zone_length(spec, r["spacingZone"], resolved, range_pick)
                size = (item.get(size_key) or item.get("sizeFreeway")
                        or item.get("sizeNonFreeway") or "")
                size = str(size).replace("x", '" x ') + '"' if size else '48" x 48"'
                sign.append(f"{a}:{key}:{side}:{spacing:g}:{size}")
            else:
                spacing = zone_length(spec, r["zone"], resolved, range_pick)
                label = canonical_order_label(r["label"])
                non_sign.append(f"{a}:{label}:{spacing:g}")
        for o in al.get("overlayZones", []):
            z = zones[o["zone"]]
            overlays.append({
                "alignIdx": a,
                "zone": o["zone"],
                "label": z["sheetLabel"],
                "lengthFt": zone_length(spec, o["zone"], resolved, range_pick),
                "anchorZone": o["anchor"]["zone"],
                "anchorEnd": o["anchor"]["end"],
                "direction": o["direction"],
            })
    return {"nonSignRows": non_sign, "signRows": sign, "overlays": overlays}


def station_walk(spec: dict, resolved: dict, range_pick: str = "min") -> list[dict]:
    """Cumulative stations per alignment -- for showing an engineer what will
    be drawn before anything is drawn."""
    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    out = []
    for al in spec["orderTable"]["alignments"]:
        sta = 0.0
        for r in al["rows"]:
            zid = r.get("spacingZone") or r["zone"]
            ln = zone_length(spec, zid, resolved, range_pick)
            sta += ln
            out.append({"alignIdx": al["alignIdx"], "alignName": al["name"],
                        "rowNum": r["rowNum"], "type": r["type"], "zone": r["zone"],
                        "lengthZone": zid,
                        "item": r.get("signCode") or r.get("label"),
                        "lengthFt": ln, "stationFt": sta})
        for o in al.get("overlayZones", []):
            anchor_sta = 0.0
            for r in al["rows"]:
                zid = r.get("spacingZone") or r["zone"]
                anchor_sta += zone_length(spec, zid, resolved, range_pick)
                if zid == o["anchor"]["zone"]:
                    break
            ln = zone_length(spec, o["zone"], resolved, range_pick)
            out.append({"alignIdx": al["alignIdx"], "alignName": al["name"],
                        "rowNum": None, "type": "Overlay", "zone": o["zone"],
                        "item": f"{zones[o['zone']]['sheetLabel']} (overlay)",
                        "lengthFt": ln, "stationFt": anchor_sta, "direction": o["direction"],
                        "note": f"runs {o['direction']} from station {anchor_sta:g} "
                                f"inside {zones[o['zone']].get('containedIn')}"})
    return out


# Spec input id -> DesignerInputs field (PlanSession lock).
_SPEC_INPUT_TO_SESSION = {
    "preconstructionPostedSpeedMph": "speed",
    "laneWidthFt": "lane_width",
    "shoulderWidthBand": "shoulder_width",
    "areaType": "area_type",
    "exposureCondition": "exposure_condition",
    "closureType": "closure_type",
    "signSizeClass": "road_type",
}


def _norm_token(v) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(v or "").lower())


def _try_derive(spec: dict, item: dict) -> Optional[tuple]:
    """Return (value, cite) when the spec already determines this input."""
    iid = item.get("id") or ""
    app = spec.get("applicability") or {}
    allowed = item.get("allowed")
    if iid == "closureType":
        c = str(app.get("closure") or "").lower()
        if "shoulder" in c and "lane" not in c:
            val = "SHOULDER CLOSURE OR ENCROACHMENT"
        elif "lane" in c:
            val = "LANE CLOSURE OR ENCROACHMENT"
        else:
            return None
        if allowed and val not in allowed:
            return None
        return (val, f"applicability.closure={app.get('closure')!r}")
    if iid == "signSizeClass":
        rt = str(app.get("roadType") or "")
        default = item.get("default")
        if not default:
            return None
        if "non-freeway" in rt.lower() and _norm_token(default) == "nonfreeway":
            return (default, f"applicability.roadType={rt!r} + inputs.default")
        if rt.lower() == "freeway" and _norm_token(default) == "freeway":
            return (default, f"applicability.roadType={rt!r} + inputs.default")
        return None
    return None


def _choice_options(item: dict, *, max_shown: int = 4) -> list[dict]:
    """ask_user_choice payload from allowed[]. Other only lists in-domain leftovers."""
    allowed = list(item.get("allowed") or [])
    default = item.get("default")
    labels = [str(v) for v in allowed]
    recommended = str(default) if default is not None and str(default) in labels else None
    if recommended is None and item.get("id") == "preconstructionPostedSpeedMph" and "45" in labels:
        recommended = "45"
    shown = list(labels)
    other_rest: list[str] = []
    if len(shown) > max_shown:
        numeric = all(
            isinstance(v, (int, float)) or str(v).lstrip("-").isdigit()
            for v in allowed
        )
        prefer: list[str] = []
        if recommended:
            prefer.append(recommended)
        if numeric and labels:
            for x in (labels[0], labels[-1]):
                if x not in prefer:
                    prefer.append(x)
        for x in shown:
            if x not in prefer:
                prefer.append(x)
        shown = prefer[: max_shown - 1]
        other_rest = [x for x in labels if x not in shown]
    options = []
    for lab in shown:
        opt: dict = {"label": lab, "description": item.get("label") or lab}
        if recommended and lab == recommended:
            opt["label"] = f"{lab} (Recommended)"
            opt["recommended"] = True
        options.append(opt)
    if other_rest:
        options.append({
            "label": "Other",
            "description": (
                "Type an exact in-domain value: " + "/".join(other_rest)
                + ". Values outside this sheet's allowed list are rejected."
            ),
        })
    return options


def _fallback_inputs_from_applicability(spec: dict) -> list[dict]:
    app = spec.get("applicability") or {}
    out: list[dict] = []
    if app.get("speedRangeMph"):
        out.append({
            "id": "preconstructionPostedSpeedMph",
            "label": "Preconstruction posted speed limit (MPH)",
            "type": "integer",
            "allowed": allowed_speeds(spec),
            "usedBy": [],
        })
    lanes = app.get("laneWidthFt")
    if lanes:
        out.append({
            "id": "laneWidthFt",
            "label": "Lane width (ft)",
            "type": "integer",
            "allowed": list(lanes),
            "usedBy": [],
        })
    bands = app.get("shoulderWidthBands")
    if bands:
        out.append({
            "id": "shoulderWidthBand",
            "label": "Shoulder width band",
            "type": "enum",
            "allowed": list(bands),
            "usedBy": [],
        })
    areas = app.get("areaTypes")
    if areas:
        out.append({
            "id": "areaType",
            "label": "Area type",
            "type": "enum",
            "allowed": list(areas),
            "usedBy": [],
        })
    return out


def required_designer_inputs(spec: dict, locked: Optional[dict] = None) -> dict:
    """Deterministic ask-list from spec['inputs'] (or applicability fallback)."""
    locked = locked or {}
    sheet = (spec.get("sheet") or {}).get("number") or ""
    raw = spec.get("inputs")
    items = list(raw) if raw else _fallback_inputs_from_applicability(spec)
    to_ask: list[dict] = []
    derived: list[dict] = []
    already: list[dict] = []
    ask_payloads: list[dict] = []
    for item in items:
        iid = item.get("id") or ""
        sess_key = _SPEC_INPUT_TO_SESSION.get(iid, iid)
        locked_val = locked.get(sess_key)
        if locked_val not in (None, "", 0):
            already.append({
                "id": iid, "sessionKey": sess_key, "value": locked_val,
                "status": "locked",
            })
            continue
        der = _try_derive(spec, item)
        if der is not None:
            val, cite = der
            derived.append({
                "id": iid, "sessionKey": sess_key, "value": val,
                "status": "derived", "cite": cite,
            })
            continue
        options = _choice_options(item)
        rec = {
            "id": iid,
            "sessionKey": sess_key,
            "label": item.get("label") or iid,
            "type": item.get("type"),
            "allowed": item.get("allowed"),
            "default": item.get("default"),
            "usedBy": item.get("usedBy"),
            "note": item.get("note"),
            "status": "ask",
            "askUserChoice": {
                "question": item.get("label") or iid,
                "options": options,
            },
        }
        to_ask.append(rec)
        ask_payloads.append(rec["askUserChoice"])
    return {
        "status": "OK",
        "sheetNum": sheet,
        "askCount": len(to_ask),
        "toAsk": to_ask,
        "derived": derived,
        "locked": already,
        "askUserChoice": ask_payloads,
        "highwayKinds": highway_kinds(spec),
        "highwayRoadway": (spec.get("applicability") or {}).get("roadway") or "",
        "note": (
            "Call ask_user_choice once per toAsk item using the provided "
            "options. Do not offer values outside allowed[]. Do not re-ask "
            "locked. Apply derived values and cite them in the journal. "
            "If the placed/locked highway kind is not in highwayKinds, "
            "caution the engineer before building — wrong sheet for that road."
        ),
    }


def validate_designer_input_value(spec: dict, input_id: str, value) -> dict:
    """Reject out-of-domain values (e.g. 60 mph on 619-311) before order table."""
    items = spec.get("inputs") or _fallback_inputs_from_applicability(spec)
    item = next((i for i in items if i.get("id") == input_id), None)
    if item is None:
        return {"ok": False, "note": f"unknown input id {input_id!r}"}
    allowed = item.get("allowed")
    if not allowed:
        return {"ok": True}
    if value in allowed or str(value) in [str(a) for a in allowed]:
        return {"ok": True, "value": value}
    note = (spec.get("applicability") or {}).get("speedRangeMph", {}).get("note")
    return {
        "ok": False,
        "value": value,
        "allowed": allowed,
        "note": note or f"{value!r} is not in allowed {allowed}",
    }


def normalize_placed_highway_kind(placed: str) -> str:
    """Map place_* / last-road tokens onto sheet highwayKinds."""
    p = (placed or "").strip().lower().replace("-", "_").replace(" ", "_")
    if p in ("two_way", "two_way_undivided", "undivided", "twoway"):
        return "two_way_undivided"
    if p in ("divided", "median"):
        return "divided"
    if p in ("freeway", "freeway_divided"):
        return "freeway"
    if p in ("twlt", "twlt_undivided"):
        return "twlt"
    if p in ("one_way", "oneway"):
        return "one_way"
    if p in ("ramp", "exit_ramp"):
        return "ramp"
    if p in ("parkway",):
        return "parkway"
    if p in ("any", "all"):
        return "any"
    return p


def highway_kinds(spec: dict) -> list[str]:
    """Canonical highway kinds this sheet applies to.

    Prefer applicability.highwayKinds when authored. Otherwise parse
    applicability.roadway (and roadType) so every 619 spec gets a check
    without editing 90 JSON files.
    """
    app = spec.get("applicability") or {}
    explicit = app.get("highwayKinds")
    if isinstance(explicit, list) and explicit:
        return [normalize_placed_highway_kind(str(x)) for x in explicit]
    road = str(app.get("roadway") or "").lower()
    rtype = str(app.get("roadType") or "").lower()
    kinds: set[str] = set()
    if "all roadway" in road:
        return ["any"]
    if "twlt" in road:
        kinds.add("twlt")
    elif "undivided" in road or "two-lane" in road or "two lane" in road:
        kinds.add("two_way_undivided")
    elif "multilane two-way" in road or "multilane two way" in road:
        kinds.add("two_way_undivided")
    if "divided" in road:
        kinds.add("divided")
    if "freeway" in road:
        kinds.add("freeway")
        kinds.add("divided")
    if "ramp" in road:
        kinds.add("ramp")
    if "parkway" in road:
        kinds.add("parkway")
    if "one-way" in road or "one way" in road:
        kinds.add("one_way")
    if not kinds and rtype == "freeway":
        kinds.update(["freeway", "divided"])
    if not kinds and rtype == "non-freeway":
        kinds.add("two_way_undivided")
    if not kinds:
        kinds.add("unknown")
    return sorted(kinds)


def highway_kind_match(spec: dict, placed_kind: str = "") -> dict:
    """Caution payload when the placed/locked road is the wrong type.

    Does not refuse the build — the agent must warn and ask. No placed
    road → compatible (abstract ticks / no striping yet).
    """
    kinds = highway_kinds(spec)
    roadway = (spec.get("applicability") or {}).get("roadway") or ""
    sheet = (spec.get("sheet") or {}).get("number") or ""
    placed = normalize_placed_highway_kind(placed_kind)
    if not placed or placed in ("unknown",):
        return {
            "mismatch": False,
            "sheetNum": sheet,
            "highwayKinds": kinds,
            "roadway": roadway,
            "placedKind": "",
        }
    ok = False
    if "any" in kinds or "unknown" in kinds:
        ok = True
    elif placed in kinds:
        ok = True
    elif placed == "divided" and ("freeway" in kinds or "divided" in kinds):
        ok = True
    elif placed == "freeway" and ("freeway" in kinds or "divided" in kinds):
        ok = True
    elif placed == "one_way" and "ramp" in kinds:
        ok = True
    elif placed == "ramp" and ("ramp" in kinds or "one_way" in kinds):
        ok = True
    out = {
        "mismatch": not ok,
        "sheetNum": sheet,
        "highwayKinds": kinds,
        "roadway": roadway,
        "placedKind": placed,
    }
    if not ok:
        out["caution"] = (
            f"Sheet {sheet} is for {roadway or kinds} "
            f"(highwayKinds={kinds}). The current road is {placed}. "
            "This is the wrong highway type for this sheet. Ask the "
            "engineer before building — switch sheets or place the "
            "matching highway."
        )
        out["askUserChoice"] = {
            "question": (
                f"{sheet} is a {roadway} sheet. The placed road is "
                f"{placed}. Continue anyway, switch sheets, or place "
                "the matching highway?"
            ),
            "options": [
                {
                    "label": "Stop — wrong highway (Recommended)",
                    "description": "Do not build. Pick a matching 619 sheet or place the right road type.",
                },
                {
                    "label": "Place matching highway first",
                    "description": "Draw the sheet's highway kind, then rebuild.",
                },
                {
                    "label": "Continue anyway",
                    "description": "Engineer overrode the caution. Build on this road.",
                },
            ],
        }
    return out


