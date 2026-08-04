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
"""
from __future__ import annotations

import json
import math
import pathlib
import re
from typing import Optional

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


def load(sheet_num: str) -> Optional[dict]:
    p = spec_path(sheet_num)
    if not p.is_file():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


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


# ============================================================
# PLACEMENT-PLAN COMPILER, STAGE 1 (Data/sheet-specs/STATUS.md)
# ============================================================
# Scoped to stations/dimensions/labels only -- channelizing/symbols/hatch
# are later stages. Replicates Modules/PerpPlacement.bas's
# PlaceOrderTableDimensions / PlaceOrderTableLabels geometry exactly (same
# PERP_HALF_LEN=40 tick-tip measurement, same outward-unit/offset math) so
# a golden-file test can diff this against 619-311's already-live-drawn
# output. The difference: this computes every point in pure Python via
# alignment_geometry.station_to_xy() from one GET_ALIGNMENT_VERTICES fetch,
# instead of VBA's GetPointAndTangent being called once per point across
# the bridge.

PERP_HALF_LEN_FT = 40.0  # Modules/PerpPlacement.bas:47 -- keep in sync

_ORDER_LABEL_KINDS = {
    "ROLL AHEAD": "RollAhead",
    "VEHICLE SPACE": "VehicleSpace",
    "BUFFER": "Buffer",
    "SHOULDER TAPER": "ShoulderTaper",
    "DOWNSTREAM TAPER": "DownstreamTaper",
    "MERGING": "MergingTaper",
    "SHIFTING TAPER": "MergingTaper",
    "LANE TAPER": "MergingTaper",
    "WORK AREA": "WorkArea",
}


def _order_label_kind(label: str) -> str:
    """Mirrors Modules/PerpPlacement.bas's OrderLabelKind() exactly --
    case-insensitive substring match, first hit wins in the same order."""
    u = " ".join((label or "").strip().upper().split())
    for needle, kind in _ORDER_LABEL_KINDS.items():
        if needle in u:
            return kind
    return ""


def _should_annotate_non_sign_label(label: str, sheet_elements: str) -> bool:
    """Mirrors Modules/PerpPlacement.bas's ShouldAnnotateNonSignLabel()
    exactly. sheet_elements is the pipe-list from get_sheet_requirements'
    'elements' field (e.g. 'MergingTaper|ShoulderTaper|...')."""
    kind = _order_label_kind(label)
    elems = sheet_elements or ""
    if kind in ("RollAhead", "VehicleSpace", "Buffer"):
        return True
    if kind == "MergingTaper":
        return "MergingTaper" in elems
    if kind == "ShoulderTaper":
        return "ShoulderTaper" in elems
    if kind == "DownstreamTaper":
        return "DownstreamTaper" in elems
    return False


def _outward_unit(tan_x: float, tan_y: float, outward_sign: float) -> tuple[float, float]:
    """Mirrors Modules/PerpPlacement.bas's OutwardUnit() exactly."""
    if outward_sign >= 0:
        return (-tan_y, tan_x)
    return (tan_y, -tan_x)


def compile_plan(spec: dict, resolved: dict, align_idx: int, segments,
                  outward_sign: float = -1.0, offset_dist: float = 15.0,
                  text_extra_along: float = 20.0, sheet_elements: str = "",
                  range_pick: str = "min") -> list[dict]:
    """Compile one alignment's stations/dimensions/labels into an explicit
    list of placement primitives in absolute model coordinates.

    segments: alignment_geometry.PathSegment list for this align_idx, from
    alignment_geometry.parse_vertices(wztc_ops.get_alignment_vertices(align_idx)).

    Each primitive is one of:
      {"kind": "station", "rowNum", "item", "stationFt", "x", "y", "tanX", "tanY"}
      {"kind": "dimension", "tip1": (x,y), "tip2": (x,y), "offset": (x,y), "text"}
      {"kind": "label", "x", "y", "text"}
    "dimension"/"label" primitive shapes are chosen to map directly onto
    PLACE_DIMENSION (x1,y1,x2,y2,ox,oy) and PLACE_TEXT_LABEL (text,x,y) --
    execution is a thin loop over these, not a translation layer.

    Does not call the bridge and does not require MicroStation to be open
    once `segments` has been fetched -- pure Python from there.
    """
    import alignment_geometry as ag  # local import: mcp-server-only dependency

    walk = [w for w in station_walk(spec, resolved, range_pick) if w["alignIdx"] == align_idx]
    primitives: list[dict] = []

    prev_x, prev_y, _, _ = ag.station_to_xy(segments, 0.0)
    for row in walk:
        if row["rowNum"] is None:
            continue  # overlay zones don't consume a station in this walk
        x, y, tan_x, tan_y = ag.station_to_xy(segments, row["stationFt"])
        primitives.append({
            "kind": "station", "rowNum": row["rowNum"], "item": row["item"],
            "stationFt": row["stationFt"], "x": x, "y": y, "tanX": tan_x, "tanY": tan_y,
        })

        out_x, out_y = _outward_unit(tan_x, tan_y, outward_sign)

        # Dimension every consecutive tick pair (tip-to-tip), Sign and
        # Non-Sign rows alike, same gate as PlaceOrderTableDimensions
        # ("spacing > 0" -- skip zero-length overlay-adjacent artifacts).
        if row["lengthFt"] > 0:
            t1x = prev_x + out_x * PERP_HALF_LEN_FT
            t1y = prev_y + out_y * PERP_HALF_LEN_FT
            t2x = x + out_x * PERP_HALF_LEN_FT
            t2y = y + out_y * PERP_HALF_LEN_FT
            ox = 0.5 * (t1x + t2x) + out_x * offset_dist
            oy = 0.5 * (t1y + t2y) + out_y * offset_dist
            primitives.append({
                "kind": "dimension", "tip1": (t1x, t1y), "tip2": (t2x, t2y),
                "offset": (ox, oy), "text": f"{row['lengthFt']:g}",
            })

        # Name labels below the dim line: Non-Sign rows only, gated on
        # whether this sheet's elements list actually has this feature
        # (same as ShouldAnnotateNonSignLabel).
        if (row["type"] == "Non-Sign" and row["lengthFt"] > 0
                and _should_annotate_non_sign_label(row["item"], sheet_elements)):
            mid_x = 0.5 * (prev_x + x) + out_x * PERP_HALF_LEN_FT
            mid_y = 0.5 * (prev_y + y) + out_y * PERP_HALF_LEN_FT
            label_out = offset_dist + text_extra_along
            tx = mid_x + out_x * label_out
            ty = mid_y + out_y * label_out
            txt = f"{row['item']} {row['lengthFt']:g}'"
            primitives.append({"kind": "label", "x": tx, "y": ty, "text": txt})

        prev_x, prev_y = x, y

    return primitives


# ============================================================
# PLACEMENT-PLAN COMPILER, STAGE 2 (Data/sheet-specs/STATUS.md)
# ============================================================
# Real counted channelizing devices, replacing
# Modules/PerpPlacement.bas's PlaceOrderTableChannelizing, which draws only
# THREE bare 2-point polylines total (no device count, no spacing) and uses
# a laneWidthFt*0.35 fudge factor for the shoulder-taper lateral offset --
# both the missing device count and the fudge factor are exactly the
# defects 619-311.json's knownCodeDeviations already documents (see
# `PerpPlacement.PlaceOrderTableChannelizing` and `rules.taper-continuity`
# there). This compiles real cone positions from the spec's own
# deviceCountSource, and connects the shoulder taper to the lane taper's
# tip at the exact shared point (zero lateral offset, no jog) instead of
# the fudge factor -- taper-continuity by construction, not by luck.

def _overlay_span(w: dict) -> tuple[float, float]:
    anchor = w["stationFt"]
    far = anchor + w["lengthFt"] if w["direction"] == "upstream" else anchor - w["lengthFt"]
    return (anchor, far)  # order preserved: [0]=anchor(offset 0), [1]=far end(full offset)


def _zone_station_ranges(spec: dict, resolved: dict, align_idx: int,
                          range_pick: str = "min") -> dict[str, tuple[float, float]]:
    """zone id -> (lo, hi) station range within one alignment's own walk,
    including overlay zones. Shared by compile_channelizing and
    compile_symbols so both stages agree on where a zone actually is."""
    walk = [w for w in station_walk(spec, resolved, range_pick) if w["alignIdx"] == align_idx]
    zone_range: dict[str, tuple[float, float]] = {}
    prev_sta = 0.0
    for w in walk:
        if w["type"] == "Overlay":
            continue
        lo, hi = sorted((prev_sta, w["stationFt"]))
        zone_range[w["zone"]] = (lo, hi)
        prev_sta = w["stationFt"]
    for w in walk:
        if w["type"] != "Overlay":
            continue
        anchor, far = _overlay_span(w)
        zone_range[w["zone"]] = (min(anchor, far), max(anchor, far))
    return zone_range


def _anchor_station(zone_range: dict[str, tuple[float, float]], zone_id: str, end: str) -> float | None:
    """Resolve a stationAnchor {zone, end} to a station. 'upstream' end =
    the higher station (station increases moving away from the work area,
    per this corridor model's convention throughout); 'downstream' = lower.
    'both' (e.g. workAreaHatch) has no single station -- returns None,
    caller decides (Stage 4 territory)."""
    rng = zone_range.get(zone_id)
    if rng is None or end not in ("upstream", "downstream"):
        return None
    lo, hi = rng
    return hi if end == "upstream" else lo


def compile_channelizing(spec: dict, resolved: dict, align_idx: int, segments,
                          lane_width_ft: float, shoulder_width_ft: float | None = None,
                          outward_sign: float = -1.0, range_pick: str = "min") -> list[dict]:
    """Cone primitives for one alignment's channelizingDevices symbol, in
    station-and-offset order along each run (endpoints included). Each
    primitive: {"kind": "cone", "run": run_id, "x", "y"} -- maps directly
    onto PLACE_CELL (or a future PLACE_ELEMENT_RUN-of-cones op); this stage
    does not decide which.

    shoulder_width_ft: the actual numeric shoulder width already supplied
    by the caller (e.g. build_wztc_order_table's shoulder_width param,
    before it's collapsed to a band for the table lookup) -- NOT invented
    here. Required only if the sheet has a shoulderTaperRun; other sheets
    ignore it. Falls back to lane_width_ft only when the sheet needs a
    shoulder run but the caller didn't pass one, with a note on the
    returned primitive marking the substitution (never silent)."""
    import alignment_geometry as ag

    sym = next((s for s in spec["symbols"]["items"] if s["id"] == "channelizingDevices"), None)
    if not sym:
        return []
    long_spacing = float((sym.get("longitudinalSpacing") or {}).get("maxFt", 40.0))

    walk = [w for w in station_walk(spec, resolved, range_pick) if w["alignIdx"] == align_idx]
    zone_range = _zone_station_ranges(spec, resolved, align_idx, range_pick)

    zone_offset_ends: dict[str, str] = {}  # zone id -> "anchor_is_lo"/"anchor_is_hi"
    for w in walk:
        if w["type"] != "Overlay":
            continue
        anchor, far = _overlay_span(w)
        # anchor end is always the shared connection point (offset 0);
        # record which physical end (lo/hi) that corresponds to.
        zone_offset_ends[w["zone"]] = "anchor_is_lo" if anchor <= far else "anchor_is_hi"

    def span_station_range(span: str):
        stas = []
        for p in span.split(".."):
            if p in zone_range:
                stas.extend(zone_range[p])
            elif p == "workArea":
                stas.append(0.0)
            else:
                return None
        return (min(stas), max(stas)) if stas else None

    primitives: list[dict] = []
    for run in sym.get("runs", []):
        zone_id = run["zone"]
        rng = span_station_range(zone_id) if ".." in zone_id else zone_range.get(zone_id)
        if rng is None:
            continue
        lo, hi = rng
        is_taper_run = run["id"] in ("laneTaperRun", "shoulderTaperRun")

        count_source = run.get("deviceCountSource")
        if count_source:
            col_key, col_field = count_source["column"].split(".")
            count = int(resolved[col_key][col_field])
            n_steps = max(count - 1, 1)
            stations = [lo + (hi - lo) * i / n_steps for i in range(count)]
        else:
            length = hi - lo
            n_steps = max(int(round(length / long_spacing)), 1) if length > 0 else 0
            stations = [lo + (hi - lo) * i / n_steps for i in range(n_steps + 1)] if length > 0 else [lo]

        # Lateral offset per station. Taper runs interpolate from 0 (at the
        # shared/tip end) to the full width (at the toe/far end); the
        # non-taper runs (longitudinal, downstream) hold a constant offset
        # equal to the closed-lane width -- they don't taper, the cone line
        # just runs straight down the closed lane.
        note = None
        if run["id"] == "laneTaperRun":
            # Convention matching every 619 sheet's own taper drawing (and
            # PerpPlacement's original "align at upstream tip" comment): the
            # HIGHER station (upstream, away from the work area) is the tip
            # (offset 0); the LOWER station (toward the work area) is the
            # toe (full lane width).
            off_lo, off_hi = lane_width_ft, 0.0
        elif run["id"] == "shoulderTaperRun":
            width = shoulder_width_ft
            if width is None:
                width = lane_width_ft
                note = "shoulder_width_ft not supplied -- substituted lane_width_ft, not a real shoulder measurement"
            anchor_is_lo = zone_offset_ends.get(zone_id) == "anchor_is_lo"
            off_lo, off_hi = (0.0, width) if anchor_is_lo else (width, 0.0)
        else:
            off_lo = off_hi = lane_width_ft

        for sta in stations:
            t = 0.0 if hi == lo else (sta - lo) / (hi - lo)
            offset = off_lo + t * (off_hi - off_lo)
            x, y, tan_x, tan_y = ag.station_to_xy(segments, sta)
            out_x, out_y = _outward_unit(tan_x, tan_y, outward_sign)
            prim = {"kind": "cone", "run": run["id"], "stationFt": sta,
                    "x": x + out_x * offset, "y": y + out_y * offset}
            if note:
                prim["note"] = note
            primitives.append(prim)

    return primitives


def check_taper_continuity(primitives: list[dict], tol_ft: float = 0.01) -> list[str]:
    """rules.taper-continuity as an executable check on compiled cone
    primitives, shipped with the stage that produces them rather than
    deferred to a later rules-engine pass (Stage 5 hardens this into a
    pre-draw gate; this is the check itself, usable standalone now).

    Finds cone runs that share a station (the shoulderTaper/laneTaper
    junction, or a taper toe meeting a longitudinal run) and asserts they
    land at the same XY within tol_ft. Returns failure strings, empty if
    every shared station is continuous."""
    by_run: dict[str, dict[float, tuple[float, float]]] = {}
    for p in primitives:
        if p["kind"] != "cone":
            continue
        by_run.setdefault(p["run"], {})[p["stationFt"]] = (p["x"], p["y"])

    fails = []
    run_ids = list(by_run.keys())
    for i, run_a in enumerate(run_ids):
        for run_b in run_ids[i + 1:]:
            shared = set(by_run[run_a]) & set(by_run[run_b])
            for sta in shared:
                xa, ya = by_run[run_a][sta]
                xb, yb = by_run[run_b][sta]
                dist = ((xa - xb) ** 2 + (ya - yb) ** 2) ** 0.5
                if dist > tol_ft:
                    fails.append(
                        f"{run_a} and {run_b} share station {sta:g} but land "
                        f"{dist:.3f} ft apart: ({xa:.3f},{ya:.3f}) vs ({xb:.3f},{yb:.3f})")
    return fails


# ============================================================
# PLACEMENT-PLAN COMPILER, STAGE 3 (Data/sheet-specs/STATUS.md)
# ============================================================
# Symbols (protective vehicles, arrow panel, vehicle-mounted signs),
# replacing Modules/PerpPlacement.bas's PlaceSheetSymbolCells, which only
# ever places ONE protective vehicle (Vehicle Space bay, Buffer Space
# fallback) and one arrow panel, with no concept of a sheet needing more
# than one PV (619-302 needs three) or of the arrow-panel/VEH#1 "OR" choice
# being an actual choice rather than always drawing the panel.
#
# This does NOT re-derive PV count from a table legend + closed-lane count
# -- that derivation already happened once, correctly, by a human/agent
# during spec authoring (see AUTHORING.md), and re-deriving it again here
# from scratch would be the exact kind of guessing sheet_spec.py's own
# module docstring says not to do. Instead: compile exactly the protective
# vehicles the spec's own symbols.items already lists (however many that
# is per sheet -- one on 619-311, three on 619-302), each from its own
# stationAnchor. That's what "PV count from the spec" means in practice.

def compile_symbols(spec: dict, resolved: dict, align_idx: int, segments,
                     outward_sign: float = -1.0, range_pick: str = "min") -> list[dict]:
    """Protective-vehicle / arrow-panel primitives for one alignment.

    Each primitive: {"kind": "protectiveVehicle"|"arrowPanel", "id",
    "cellName", "x", "y", "angleDeg", "stationFt", "requiredNote"}.
    requiredNote carries the spec's own conditional-requirement text
    verbatim (e.g. "only when the shoulder width is >= 8 ft") -- this
    compiler does not evaluate that condition, it surfaces it so the
    caller/engineer decides, the same "explicit choice, not auto-decided"
    principle as the arrow-panel/VEH#1 alternative below.

    An item with an "alternative" (e.g. arrowPanel's 'OR VEH #1') gets an
    altGroup tag cross-referencing its partner item (matched by the
    alternative's option text against another symbol's sheetLabel) --
    both primitives are compiled and returned; picking one is left to the
    caller. Vehicle-mounted signs (signs.items with postMounted:false and
    mountedOn:<symbol id>) become a 'vehicleMountedSign' primitive at that
    vehicle's own computed position, not an independent post."""
    import alignment_geometry as ag

    zone_range = _zone_station_ranges(spec, resolved, align_idx, range_pick)
    signs_mounted_on: dict[str, dict] = {
        s["mountedOn"]: s for s in spec["signs"]["items"]
        if s.get("postMounted") is False and s.get("mountedOn")
    }

    primitives: list[dict] = []
    prim_by_id: dict[str, dict] = {}

    for item in spec["symbols"]["items"]:
        anchor = item.get("stationAnchor")
        if not anchor:
            continue
        is_vehicle = bool(item.get("cellHint"))
        is_arrow_panel = (item["id"] == "arrowPanel")
        if not (is_vehicle or is_arrow_panel):
            continue  # spotter / workAreaHatch / etc. -- Stage 4 territory

        sta = _anchor_station(zone_range, anchor["zone"], anchor["end"])
        if sta is None:
            continue
        x, y, tan_x, tan_y = ag.station_to_xy(segments, sta)
        out_x, out_y = _outward_unit(tan_x, tan_y, outward_sign)
        px = x + out_x * PERP_HALF_LEN_FT
        py = y + out_y * PERP_HALF_LEN_FT
        angle_deg = math.degrees(math.atan2(tan_y, tan_x))

        prim = {
            "kind": "protectiveVehicle" if is_vehicle else "arrowPanel",
            "id": item["id"], "cellName": item.get("cellHint") or "TWZAP_P",
            "x": px, "y": py, "angleDeg": angle_deg, "stationFt": sta,
            "requiredNote": item.get("required"),
        }
        primitives.append(prim)
        prim_by_id[item["id"]] = prim

        if item["id"] in signs_mounted_on:
            sign = signs_mounted_on[item["id"]]
            primitives.append({
                "kind": "vehicleMountedSign", "signCode": sign["signCode"],
                "mountedOn": item["id"], "x": px, "y": py, "angleDeg": angle_deg,
            })

    # Second pass: link "OR" alternatives between already-compiled
    # primitives instead of synthesizing a duplicate entry.
    alt_counter = 0
    for item in spec["symbols"]["items"]:
        alt = item.get("alternative")
        if not alt or item["id"] not in prim_by_id:
            continue
        option_text = (alt.get("option") or "").strip().upper()
        partner_id = None
        for other in spec["symbols"]["items"]:
            label = (other.get("sheetLabel") or "").strip().upper()
            if other["id"] != item["id"] and label and label == option_text and other["id"] in prim_by_id:
                partner_id = other["id"]
                break
        alt_counter += 1
        group = f"alt{alt_counter}"
        prim_by_id[item["id"]]["altGroup"] = group
        prim_by_id[item["id"]]["altDescription"] = alt.get("description")
        if partner_id:
            prim_by_id[partner_id]["altGroup"] = group
        else:
            prim_by_id[item["id"]]["altPartnerNote"] = (
                f"alternative option {alt.get('option')!r} did not match any other "
                f"symbol's sheetLabel -- alternative described but not cross-linked")

    return primitives


# ============================================================
# PLACEMENT-PLAN COMPILER, STAGE 4 (Data/sheet-specs/STATUS.md)
# ============================================================
# Work-area hatch boundary as an explicit polygon, plus the conditional
# Detail-A/Note-N transverse device rows. Replaces
# Modules/PerpPlacement.bas's PlaceOrderTableWorkspace, whose current
# bounds run from path start through the Vehicle Space station -- which
# wrongly includes the roll ahead distance inside the hatch (exactly what
# rules.no-occupancy-buffer-rollahead exists to catch).
#
# Design correction made while writing this: the work area is NOT reachable
# by walking either alignment's own positive-station direction at all.
# orderTable.alignments[0].station0 ("Upstream" align) is literally defined
# as "Upstream edge of the WORK AREA", and alignments[1].station0
# ("Downstream" align) as "Downstream edge of the WORK AREA" -- each
# alignment starts AT its own edge of the work area and walks AWAY from it
# (upstream/downstream respectively). So the work area's length is not a
# number to invent OR accept as a bare external parameter: it is already
# implicit in how the engineer/agent committed the two alignments -- it is
# literally the real-world distance between align1's station-0 point and
# align2's station-0 point. Using a single alignment's frame with a
# "work_area_length_ft" the caller supplies (an earlier version of this
# function did that) would have needed the caller to already know a number
# that duplicates information the two committed alignments already encode,
# with no way to keep the two in sync.

def compile_hatch(spec: dict, resolved: dict, align1_segments, align2_segments,
                   lane_width_ft: float, shoulder_width_ft: float | None = None,
                   outward_sign: float = -1.0) -> list[dict]:
    """Work-area hatch boundary (kind='hatch') plus conditional transverse
    device rows (kind='transverseRun') when the sheet's
    channelizingDevices.transverse condition is met.

    align1_segments/align2_segments: alignment_geometry.PathSegment lists
    for the Upstream and Downstream alignments respectively (from
    GET_ALIGNMENT_VERTICES on each). Their own station-0 points ARE the
    work area's two edges -- see the module note above for why this isn't
    a separate numeric input. If the two alignments were committed with
    materially different tangents at station 0, that's a real drawing
    problem this function surfaces (workAreaLengthFt would reflect the
    straight-line distance between two points that should coincide with
    the roadway edge, not something this function silently papers over)."""
    import alignment_geometry as ag

    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    wa = zones.get("workArea")
    if not wa or not wa.get("hatched"):
        return []

    p1x, p1y, tan1x, tan1y = ag.station_to_xy(align1_segments, 0.0)
    p2x, p2y, _, _ = ag.station_to_xy(align2_segments, 0.0)
    length = math.hypot(p2x - p1x, p2y - p1y)

    spans_shoulder = "shoulder" in (wa.get("note") or "").lower()
    width = lane_width_ft + (shoulder_width_ft or 0.0) if spans_shoulder else lane_width_ft
    out_x, out_y = _outward_unit(tan1x, tan1y, outward_sign)

    def offset_pt(x: float, y: float, off: float) -> tuple[float, float]:
        return (x + out_x * off, y + out_y * off)

    boundary = [offset_pt(p1x, p1y, 0.0), offset_pt(p2x, p2y, 0.0),
                offset_pt(p2x, p2y, width), offset_pt(p1x, p1y, width)]
    primitives: list[dict] = [{
        "kind": "hatch", "id": "workAreaHatch", "boundary": boundary,
        "workAreaLengthFt": length, "widthFt": width,
    }]

    sym = next((s for s in spec["symbols"]["items"] if s["id"] == "channelizingDevices"), None)
    transverse = (sym or {}).get("transverse")
    if transverse and str(transverse.get("required")).lower() == "conditional":
        max_spacing = float(transverse.get("maxSpacingFt", 800))
        # Condition text is sheet-specific prose (e.g. "a paved shoulder >= 8'
        # closed for a distance greater than 800'"); this compiler checks
        # the one numeric part of it (work area length vs maxSpacingFt) and
        # surfaces the rest as a note for the caller/engineer to confirm --
        # it does not parse or evaluate the shoulder-width clause.
        if length > max_spacing:
            n_rows = int(length // max_spacing)
            for k in range(1, n_rows + 1):
                t = (k * max_spacing) / length
                if t >= 1.0:
                    break
                sx, sy = p1x + (p2x - p1x) * t, p1y + (p2y - p1y) * t
                primitives.append({
                    "kind": "transverseRun", "run": "transverse", "stationFromP1Ft": k * max_spacing,
                    "tip1": offset_pt(sx, sy, 0.0), "tip2": offset_pt(sx, sy, width),
                    "note": transverse.get("sheetText"),
                })

    return primitives


# ============================================================
# PLACEMENT-PLAN COMPILER, STAGE 5 (Data/sheet-specs/STATUS.md)
# ============================================================
# A hard pre-draw gate over the compiled primitives from Stages 1-4,
# checking the subset of each sheet's own rules[] that's mechanically
# checkable from geometry alone (no MicroStation needed). This is the
# "wrong is caught at compile time, not discovered in a screenshot three
# passes later" goal from the original design discussion.
#
# Honest scope: this does NOT cover every rules[] entry on every sheet --
# some (e.g. sign-spacing-source, speed-range, shoulder-band-collapse) are
# about table-lookup correctness, already covered by sheet_spec.resolve()
# and scripts/validate_sheet_spec.py's domain-invariant pass, not by
# anything geometric a compiled plan could re-check. And this stage does
# NOT retire Modules/PerpPlacement.bas's PlaceOrderTable* functions -- the
# original plan's own bar for that ("golden test passes for at least one
# reference sheet per family, 9 families") is real, substantial follow-up
# work, not something this pass claims to have done. What IS delivered:
# the gate mechanism itself, covering every geometric rule practical to
# check today, run against two different sheet families (619-311, 619-302)
# to prove it isn't 619-311-specific.

def run_rules_gate(spec: dict, resolved: dict, align_idx: int,
                    plan_primitives: list[dict], channelizing_primitives: list[dict],
                    symbol_primitives: list[dict], hatch_primitives: list[dict] | None = None) -> list[str]:
    """Returns failure strings (empty = gate passes). Call before any
    primitive list reaches the bridge."""
    fails: list[str] = []

    # taper-continuity (already built in Stage 2 -- reused, not duplicated)
    fails += [f"taper-continuity: {f}" for f in check_taper_continuity(channelizing_primitives)]

    # cone-spacing: no consecutive pair of cones in the same run may exceed
    # 40 ft of STATION separation (the sheet's own longitudinal spacing cap;
    # taper runs are exact by construction from the device count, so this
    # mainly catches a run that came out with too few points).
    by_run: dict[str, list[float]] = {}
    for p in channelizing_primitives:
        if p["kind"] == "cone":
            by_run.setdefault(p["run"], []).append(p["stationFt"])
    for run_id, stas in by_run.items():
        stas.sort()
        for a, b in zip(stas, stas[1:]):
            if b - a > 40.0 + 0.01:
                fails.append(f"cone-spacing: {run_id} has a {b - a:.1f} ft gap "
                             f"(stations {a:g}->{b:g}), exceeds the 40 ft cap")

    # sign-order: on the alignment station_walk gives us, Sign rows should
    # appear in strictly increasing station order walking upstream (this is
    # true by construction from the spec's own rowNum ordering -- checked
    # here as a real assertion, not assumed, in case a future spec authors
    # rows out of order).
    sign_stations = [(p["rowNum"], p["stationFt"]) for p in plan_primitives
                     if p["kind"] == "station"]
    for (n1, s1), (n2, s2) in zip(sign_stations, sign_stations[1:]):
        if s2 <= s1:
            fails.append(f"sign-order: row {n2} (station {s2:g}) is not further "
                         f"upstream than row {n1} (station {s1:g})")

    # arrow-panel-anchor: arrow panel and its station-matched taper zone tip
    # must be the same station (this is also what altGroup cross-linking to
    # a same-station vehicle alternative already implies, checked directly
    # here regardless of whether an alternative exists).
    zone_range = _zone_station_ranges(spec, resolved, align_idx)
    ap = next((p for p in symbol_primitives if p["kind"] == "arrowPanel"), None)
    if ap is not None:
        lane_rng = zone_range.get("laneTaper")
        if lane_rng is not None:
            expected_sta = max(lane_rng)  # upstream end = tip, per _anchor_station's convention
            if abs(ap["stationFt"] - expected_sta) > 0.01:
                fails.append(f"arrow-panel-anchor: arrow panel at station {ap['stationFt']:g}, "
                             f"expected laneTaper's upstream end at {expected_sta:g}")

    # no-occupancy-buffer-rollahead: compile_hatch's corrected design (see
    # its own module note) builds the hatch from align1/align2's own
    # station-0 points, which sit on the opposite side of station 0 from
    # rollAheadDistance/bufferSpace's positive-station ranges -- so an
    # overlap is structurally impossible by construction now, not just
    # checked for. What IS worth guarding: a degenerate (near-zero-length)
    # hatch, which usually means the two alignments were committed with
    # coincident or badly mismatched station-0 points rather than at the
    # work area's actual two edges.
    hatch = next((p for p in (hatch_primitives or []) if p.get("kind") == "hatch"), None)
    if hatch is not None and float(hatch["workAreaLengthFt"]) < 1.0:
        fails.append(f"no-occupancy-buffer-rollahead: work area length "
                     f"{hatch['workAreaLengthFt']:.2f} ft is near zero -- align1/align2 "
                     f"station-0 points may not be committed at the work area's real edges")

    return fails
