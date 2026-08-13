"""Placement-plan compiler: turns a sheet_resolve.resolve() result into
explicit drawing primitives in absolute model coordinates (Stages 1-4 of
the compiler, see Data/sheet-specs/STATUS.md).

Part of the sheet_spec split (2026-08-04): sheet_resolve.py owns "what does
this sheet need" (table lookups); this module owns "turn that into
coordinates"; sheet_rules.py validates the primitives this module produces
before they reach the bridge. sheet_spec.py re-exports all three so every
existing `sheet_spec.X` call site keeps working unchanged.
"""
from __future__ import annotations

import math

from sheet_resolve import station_walk

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

# Defaults match today's hardcoded engineer-reference style. Sheet JSON may
# override via top-level "annotationStyle" (see Data/sheet-specs/619-311.json).
_DEFAULT_ANNOTATION_STYLE = {
    "dimensionText": "lengthOnly",
    "featureLabel": "nameOnly",
    "overlayDimSide": "opposite",
    "offsetsFt": {
        "dimOutward": 15.0,
        "labelExtra": 20.0,
        "symbolLabel": 25.0,
    },
}

_DEFAULT_CHANNELIZING_REPRESENTATION = {
    "mode": "markers",
    "markerHalfSizeFt": 1.5,
}


def annotation_style(spec: dict | None) -> dict:
    """Merge sheet annotationStyle over defaults (deep-merge offsetsFt)."""
    style = {
        "dimensionText": _DEFAULT_ANNOTATION_STYLE["dimensionText"],
        "featureLabel": _DEFAULT_ANNOTATION_STYLE["featureLabel"],
        "overlayDimSide": _DEFAULT_ANNOTATION_STYLE["overlayDimSide"],
        "offsetsFt": dict(_DEFAULT_ANNOTATION_STYLE["offsetsFt"]),
    }
    raw = (spec or {}).get("annotationStyle") or {}
    for key in ("dimensionText", "featureLabel", "overlayDimSide"):
        if key in raw and raw[key] is not None:
            style[key] = raw[key]
    offs = raw.get("offsetsFt") or {}
    for key in ("dimOutward", "labelExtra", "symbolLabel"):
        if key in offs and offs[key] is not None:
            style["offsetsFt"][key] = float(offs[key])
    return style


def channelizing_representation(sym: dict | None) -> dict:
    """Merge channelizingDevices.representation over marker defaults."""
    out = dict(_DEFAULT_CHANNELIZING_REPRESENTATION)
    raw = (sym or {}).get("representation") or {}
    if raw.get("mode"):
        out["mode"] = str(raw["mode"])
    if raw.get("markerHalfSizeFt") is not None:
        out["markerHalfSizeFt"] = float(raw["markerHalfSizeFt"])
    return out


def _primitive_id(align_idx: int, ref: str, kind: str) -> str:
    """Stable id: {align}:{zone|runId|symbolId}:{kind}."""
    clean = (ref or "unknown").replace(" ", "")
    return f"{int(align_idx)}:{clean}:{kind}"


def _dim_text(length_ft: float, style: dict) -> str:
    policy = style.get("dimensionText") or "lengthOnly"
    if policy == "lengthOnly":
        return f"{length_ft:g}"
    # Unknown policies fall back to length-only (never invent names on dims).
    return f"{length_ft:g}"


def _label_text(item: str, length_ft: float, style: dict) -> str:
    name = _feature_label_text(item)
    policy = style.get("featureLabel") or "nameOnly"
    if policy == "nameAndLength":
        return f"{name} {length_ft:g}'"
    return name


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


def _feature_label_text(item: str) -> str:
    """Plan feature name only — length lives on the DimensionElement.
    Strips the station_walk '(overlay)' suffix for overlay zones."""
    t = " ".join((item or "").strip().split())
    if t.upper().endswith("(OVERLAY)"):
        t = t[: t.upper().rfind("(OVERLAY)")].strip()
    return t


def _outward_unit(tan_x: float, tan_y: float, outward_sign: float) -> tuple[float, float]:
    """Mirrors Modules/PerpPlacement.bas's OutwardUnit() exactly."""
    if outward_sign >= 0:
        return (-tan_y, tan_x)
    return (tan_y, -tan_x)


def _dim_tip_path(segments, sta0: float, sta1: float, outward_sign: float,
                  tip_off: float, step_ft: float = 20.0,
                  align_idx: int = 1) -> list[tuple[float, float]]:
    """Sample tip points (align + local outward * tip_off) along a station span.

    This is the polyline a curved plan dimension should hug — parallel to the
    alignment at tick-tip offset, not the straight chord between ends.
    Align2 flips tan so tips stay on the closed shoulder.
    """
    import alignment_geometry as ag

    def _closed(tx: float, ty: float) -> tuple[float, float]:
        if int(align_idx) == 2:
            tx, ty = -tx, -ty
        return _outward_unit(tx, ty, outward_sign)

    span = abs(sta1 - sta0)
    if span < 1e-9:
        x, y, tx, ty = ag.station_to_xy(segments, sta0)
        ox, oy = _closed(tx, ty)
        return [(x + ox * tip_off, y + oy * tip_off)]
    n = max(1, int(math.ceil(span / max(step_ft, 1.0))))
    out: list[tuple[float, float]] = []
    for i in range(n + 1):
        t = i / n
        sta = sta0 + (sta1 - sta0) * t
        x, y, tx, ty = ag.station_to_xy(segments, sta)
        ox, oy = _closed(tx, ty)
        pt = (x + ox * tip_off, y + oy * tip_off)
        if out and math.hypot(pt[0] - out[-1][0], pt[1] - out[-1][1]) < 1e-6:
            continue
        out.append(pt)
    return out


def _path_sagitta(path: list[tuple[float, float]], tip1, tip2) -> float:
    """Max distance from path mid-samples to the tip1→tip2 chord."""
    if len(path) < 3:
        return 0.0
    x1, y1 = float(tip1[0]), float(tip1[1])
    x2, y2 = float(tip2[0]), float(tip2[1])
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)
    if L < 1e-9:
        return 0.0
    ux, uy = dx / L, dy / L
    best = 0.0
    for x, y in path[1:-1]:
        # perpendicular distance to infinite line through tips
        sx, sy = x - x1, y - y1
        along = sx * ux + sy * uy
        px, py = x1 + ux * along, y1 + uy * along
        best = max(best, math.hypot(x - px, y - py))
    return best


def _path_length(path: list[tuple[float, float]]) -> float:
    total = 0.0
    for i in range(1, len(path)):
        total += math.hypot(path[i][0] - path[i - 1][0],
                            path[i][1] - path[i - 1][1])
    return total


def compile_plan(spec: dict, resolved: dict, align_idx: int, segments,
                  outward_sign: float = -1.0, offset_dist: float | None = None,
                  text_extra_along: float | None = None, sheet_elements: str = "",
                  range_pick: str = "min",
                  tip_half_len_ft: float | None = None) -> list[dict]:
    """Compile one alignment's stations/dimensions/labels into an explicit
    list of placement primitives in absolute model coordinates.

    segments: alignment_geometry.PathSegment list for this align_idx, from
    alignment_geometry.parse_vertices(wztc_ops.get_alignment_vertices(align_idx)).

    tip_half_len_ft: real-road locked half_len (lane+shoulder) so dim/label
    tips sit on the same EOP as signs/ticks. Default PERP_HALF_LEN_FT=40
    for abstract order-table ticks.

    Each primitive is one of:
      {"kind": "station", "rowNum", "item", "stationFt", "x", "y", "tanX", "tanY"}
      {"kind": "dimension", "tip1": (x,y), "tip2": (x,y), "offset": (x,y), "text"}
      {"kind": "label", "x", "y", "text"}
    Offsets / label text / overlay side come from spec["annotationStyle"]
    (defaults = engineer reference style). Each primitive carries
    primitiveId + specRef for the placement registry.

    Does not call the bridge and does not require MicroStation to be open
    once `segments` has been fetched -- pure Python from there.
    """
    import alignment_geometry as ag  # local import: mcp-server-only dependency

    style = annotation_style(spec)
    offs = style["offsetsFt"]
    if offset_dist is None:
        offset_dist = float(offs["dimOutward"])
    if text_extra_along is None:
        text_extra_along = float(offs["labelExtra"])
    tip_hl = (
        float(tip_half_len_ft)
        if tip_half_len_ft is not None and float(tip_half_len_ft) > 0
        else PERP_HALF_LEN_FT
    )

    walk = [w for w in station_walk(spec, resolved, range_pick) if w["alignIdx"] == align_idx]
    primitives: list[dict] = []
    zone_defs = {z["id"]: z for z in spec.get("corridor", {}).get("zones", [])}

    def _zone_dimensioned(zone_id: str | None) -> bool:
        # Sheet's own "dimensioned" flag wins; default True for zones that
        # don't specify it (e.g. taper/buffer/roll-ahead all say True anyway).
        return bool(zone_defs.get(zone_id, {}).get("dimensioned", True))

    def _closed_out(tx: float, ty: float, o_sign: float) -> tuple[float, float]:
        # Align2 tan points +travel; flip to Align1-equivalent basis so
        # dim/label tips stay on the closed shoulder (same as signs).
        if int(align_idx) == 2:
            tx, ty = -tx, -ty
        return _outward_unit(tx, ty, o_sign)

    prev_x, prev_y, prev_tx, prev_ty = ag.station_to_xy(segments, 0.0)
    prev_sta = 0.0
    for row in walk:
        if row["rowNum"] is None:
            continue  # overlay zones don't consume a station in this walk
        zone = str(row.get("zone") or row.get("item") or f"row{row['rowNum']}")
        x, y, tan_x, tan_y = ag.station_to_xy(segments, row["stationFt"])
        primitives.append({
            "kind": "station", "rowNum": row["rowNum"], "item": row["item"],
            "stationFt": row["stationFt"], "x": x, "y": y, "tanX": tan_x, "tanY": tan_y,
            "primitiveId": _primitive_id(align_idx, zone, "station"),
            "specRef": {"zone": zone, "run": None, "alignIdx": align_idx},
        })

        # Per-station outward so dim tips stay local-normal on curved
        # corridors (reusing the far tick's normal at both ends skewed
        # tip-to-tip dims across bends).
        out_prev_x, out_prev_y = _closed_out(prev_tx, prev_ty, outward_sign)
        out_x, out_y = _closed_out(tan_x, tan_y, outward_sign)
        mid_ox = 0.5 * (out_prev_x + out_x)
        mid_oy = 0.5 * (out_prev_y + out_y)
        mid_mag = math.hypot(mid_ox, mid_oy)
        if mid_mag > 1e-9:
            mid_ox /= mid_mag
            mid_oy /= mid_mag
        else:
            mid_ox, mid_oy = out_x, out_y

        # Dimension every consecutive tick pair (tip-to-tip), Sign and
        # Non-Sign rows alike, same gate as PlaceOrderTableDimensions
        # ("spacing > 0" -- skip zero-length overlay-adjacent artifacts) --
        # unless the zone that actually determines this row's length is
        # marked dimensioned=False on the sheet (e.g. gapEndRoadWork, which
        # the real sheet expresses as a text callout, not a dimension line).
        #
        # On curved corridors: path-hugging dim along the tip-offset roadside
        # (sheet length text). Linear Size chords cut through the pavement and
        # measure the wrong length (live QA 2026-08-13).
        if row["lengthFt"] > 0 and _zone_dimensioned(row.get("lengthZone", zone)):
            t1x = prev_x + out_prev_x * tip_hl
            t1y = prev_y + out_prev_y * tip_hl
            t2x = x + out_x * tip_hl
            t2y = y + out_y * tip_hl
            tip_path = _dim_tip_path(
                segments, prev_sta, float(row["stationFt"]),
                outward_sign, tip_hl, step_ft=10.0, align_idx=align_idx)
            chord = math.hypot(t2x - t1x, t2y - t1y)
            path_len = _path_length(tip_path)
            sag = _path_sagitta(tip_path, (t1x, t1y), (t2x, t2y))
            sheet_len = float(row["lengthFt"])
            # Hug when the tip path bows OR when a Linear Size chord would
            # disagree with the sheet/table length (e.g. downstream taper
            # showing ~45' instead of 50').
            curved = (
                len(tip_path) >= 2
                and (
                    sag > 0.5
                    or path_len > chord * 1.005 + 0.25
                    or abs(chord - sheet_len) > 1.0
                )
            )
            if curved and tip_path:
                mid_i = len(tip_path) // 2
                mx, my = tip_path[mid_i]
                mid_sta = 0.5 * (prev_sta + float(row["stationFt"]))
                _, _, mtx, mty = ag.station_to_xy(segments, mid_sta)
                mox, moy = _closed_out(mtx, mty, outward_sign)
                ox = mx + mox * offset_dist
                oy = my + moy * offset_dist
            else:
                ox = 0.5 * (t1x + t2x) + mid_ox * offset_dist
                oy = 0.5 * (t1y + t2y) + mid_oy * offset_dist
            dim_prim: dict = {
                "kind": "dimension", "tip1": (t1x, t1y), "tip2": (t2x, t2y),
                "offset": (ox, oy), "text": _dim_text(row["lengthFt"], style),
                "curved": curved,
                "primitiveId": _primitive_id(align_idx, zone, "dimension"),
                "specRef": {"zone": zone, "run": None, "alignIdx": align_idx},
            }
            if curved:
                dim_prim["path"] = tip_path
            primitives.append(dim_prim)

        # Name labels below the dim line: Non-Sign rows only, gated on
        # whether this sheet's elements list actually has this feature
        # (same as ShouldAnnotateNonSignLabel). Length stays on the dim —
        # do not duplicate it in the label text (engineer reference style).
        if (row["type"] == "Non-Sign" and row["lengthFt"] > 0
                and _should_annotate_non_sign_label(row["item"], sheet_elements)):
            mid_sta = 0.5 * (prev_sta + float(row["stationFt"]))
            mx, my, mtx, mty = ag.station_to_xy(segments, mid_sta)
            mox, moy = _closed_out(mtx, mty, outward_sign)
            mid_x = mx + mox * tip_hl
            mid_y = my + moy * tip_hl
            label_out = offset_dist + text_extra_along
            tx = mid_x + mox * label_out
            ty = mid_y + moy * label_out
            primitives.append({
                "kind": "label", "x": tx, "y": ty,
                "text": _label_text(row["item"], row["lengthFt"], style),
                "primitiveId": _primitive_id(align_idx, zone, "label"),
                "specRef": {"zone": zone, "run": None, "alignIdx": align_idx},
            })

        prev_x, prev_y, prev_tx, prev_ty = x, y, tan_x, tan_y
        prev_sta = float(row["stationFt"])

    # Overlay zones (e.g. SHOULDER TAPER inside gap A): dimension + label on
    # the opposite side of the main sign/dim column by default
    # (annotationStyle.overlayDimSide). Still no sequential station tick.
    overlay_sign = (
        -outward_sign if style.get("overlayDimSide", "opposite") == "opposite"
        else outward_sign
    )
    for row in walk:
        if row.get("type") != "Overlay" or float(row.get("lengthFt") or 0) <= 0:
            continue
        label = _feature_label_text(row["item"])
        if not _should_annotate_non_sign_label(label, sheet_elements):
            continue
        zone = str(row.get("zone") or label)
        anchor, far = _overlay_span(row)
        ax, ay, atx, aty = ag.station_to_xy(segments, anchor)
        fx, fy, ftx, fty = ag.station_to_xy(segments, far)
        mx, my, mtx, mty = ag.station_to_xy(segments, 0.5 * (anchor + far))
        out_a_x, out_a_y = _closed_out(atx, aty, overlay_sign)
        out_f_x, out_f_y = _closed_out(ftx, fty, overlay_sign)
        out_x, out_y = _closed_out(mtx, mty, overlay_sign)
        t1x = ax + out_a_x * tip_hl
        t1y = ay + out_a_y * tip_hl
        t2x = fx + out_f_x * tip_hl
        t2y = fy + out_f_y * tip_hl
        tip_path = _dim_tip_path(
            segments, float(anchor), float(far),
            overlay_sign, tip_hl, step_ft=10.0, align_idx=align_idx)
        chord = math.hypot(t2x - t1x, t2y - t1y)
        path_len = _path_length(tip_path)
        sag = _path_sagitta(tip_path, (t1x, t1y), (t2x, t2y))
        sheet_len = float(row["lengthFt"])
        curved = (
            len(tip_path) >= 2
            and (
                sag > 0.5
                or path_len > chord * 1.005 + 0.25
                or abs(chord - sheet_len) > 1.0
            )
        )
        if curved and tip_path:
            mid_i = len(tip_path) // 2
            px, py = tip_path[mid_i]
            ox = px + out_x * offset_dist
            oy = py + out_y * offset_dist
        else:
            ox = 0.5 * (t1x + t2x) + out_x * offset_dist
            oy = 0.5 * (t1y + t2y) + out_y * offset_dist
        dim_prim: dict = {
            "kind": "dimension", "tip1": (t1x, t1y), "tip2": (t2x, t2y),
            "offset": (ox, oy), "text": _dim_text(row["lengthFt"], style),
            "curved": curved,
            "primitiveId": _primitive_id(align_idx, zone, "dimension"),
            "specRef": {"zone": zone, "run": None, "alignIdx": align_idx,
                        "overlay": True},
        }
        if curved:
            dim_prim["path"] = tip_path
        primitives.append(dim_prim)
        label_out = offset_dist + text_extra_along
        primitives.append({
            "kind": "label",
            "x": mx + out_x * (tip_hl + label_out),
            "y": my + out_y * (tip_hl + label_out),
            "text": _label_text(row["item"], row["lengthFt"], style),
            "primitiveId": _primitive_id(align_idx, zone, "label"),
            "specRef": {"zone": zone, "run": None, "alignIdx": align_idx,
                        "overlay": True},
        })

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
    representation = channelizing_representation(sym)

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
    # Adjacent runs (e.g. shoulderTaperRun/laneTaperRun, laneTaperRun/
    # longitudinalRun) share their boundary station by construction (the
    # taper-continuity point), which used to place two device primitives
    # stacked at the exact same (x, y) -- one from each run's endpoint-
    # inclusive station list. Track physical points already used so the
    # second run's shared endpoint is skipped instead of doubled (real
    # miss found live 2026-08-10: "two sets of channelizing devices").
    placed_points: set[tuple[float, float]] = set()
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
            n_steps = max(math.ceil(length / long_spacing), 1) if length > 0 else 0
            stations = [lo + (hi - lo) * i / n_steps for i in range(n_steps + 1)] if length > 0 else [lo]

        # Lateral offset from the ALIGNMENT, which assemble_corridor /
        # compile_hatch treat as the LEFT EDGE of the closed lane
        # (longitudinal channelizing / lane-line between open and closed).
        # Live miss 2026-08-10 real-road 619-311: offsets assumed a
        # centerline-style align and put longitudinal cones at +lane_width
        # (fog line / "middle of the road") while hatch correctly spanned
        # align→EOP. Sheet lateralOffsets.channelizingDeviceLine: cones run
        # JUST OUTBOARD of that lane line (= offset ~0 on this align).
        #
        # Lane taper (direction of travel approach): tip UPSTREAM at the
        # OUTER edge of the closed travel lane (+lane_width); toe toward
        # the work area on the channelizing line (0). Shoulder taper
        # continues that line from travel-outer (+lane) out to paved
        # shoulder EOP (+lane+shoulder). Downstream reopens: work-end on
        # channelizing (0), far end at travel-outer. Align2's outward is
        # flipped in world XY vs Align1, so Align2 uses negative offsets
        # to stay on the closed side.
        note = None
        if run["id"] == "laneTaperRun":
            # hi station = upstream tip at outer travel edge; lo = toe on
            # channelizing line (align).
            off_lo, off_hi = 0.0, lane_width_ft
        elif run["id"] == "shoulderTaperRun":
            width = shoulder_width_ft
            if width is None:
                width = lane_width_ft
                note = "shoulder_width_ft not supplied -- substituted lane_width_ft, not a real shoulder measurement"
            # Shared continuity point with lane taper tip = +lane_width.
            # Far upstream end = outer EOP = lane + shoulder.
            anchor_off = lane_width_ft
            far_off = lane_width_ft + float(width)
            anchor_is_lo = zone_offset_ends.get(zone_id) == "anchor_is_lo"
            off_lo, off_hi = (anchor_off, far_off) if anchor_is_lo else (far_off, anchor_off)
        elif run["id"] == "downstreamRun":
            # Align 2 travels opposite Align 1 — same outward_sign flips
            # world side. Negative offsets keep cones on the closed side.
            # Work-area end (nearer |sta|=0) on channelizing; far end at
            # outer travel edge so the lane reopens.
            if abs(lo) <= abs(hi):
                off_lo, off_hi = 0.0, -lane_width_ft
            else:
                off_lo, off_hi = -lane_width_ft, 0.0
        else:
            # longitudinalRun (buffer / roll-ahead / work): on the
            # channelizing line (align).
            off_lo = off_hi = 0.0

        for sta in stations:
            t = 0.0 if hi == lo else (sta - lo) / (hi - lo)
            offset = off_lo + t * (off_hi - off_lo)
            x, y, tan_x, tan_y = ag.station_to_xy(segments, sta)
            out_x, out_y = _outward_unit(tan_x, tan_y, outward_sign)
            point_key = (round(x + out_x * offset, 3), round(y + out_y * offset, 3))
            if point_key in placed_points:
                continue
            placed_points.add(point_key)
            prim = {
                "kind": "cone", "run": run["id"], "stationFt": sta,
                "x": x + out_x * offset, "y": y + out_y * offset,
                "representation": dict(representation),
                "primitiveId": _primitive_id(align_idx, run["id"], "cone"),
                "specRef": {
                    "zone": zone_id, "run": run["id"], "alignIdx": align_idx,
                },
            }
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

def _lateral_offset_ft(lateral_anchor: str | None, lane_width_ft: float | None,
                        shoulder_width_ft: float | None) -> tuple[float, str | None]:
    """Interprets a symbol's spec-authored `lateralAnchor` prose (e.g. "On
    the paved shoulder, outboard of the closed travel lane" / "In the
    closed travel lane, parallel to traffic") into a lateral offset from
    the alignment line, using the sheet's own lane/shoulder widths.

    Previously compile_symbols placed every vehicle/arrow-panel primitive
    at the fixed PERP_HALF_LEN_FT=40ft reference-tick offset regardless of
    lateralAnchor -- that constant sizes the unrelated perpendicular
    reference tick lines, not a real lane/shoulder position. Confirmed
    live 2026-08-04 as the root cause of QA findings "PV in wrong bay" /
    "mispositioned arrow panel". Returns (offsetFt, warningNote);
    warningNote is set whenever the text doesn't match a recognized
    lane/shoulder pattern (or the needed width wasn't supplied), so the
    caller surfaces the fallback rather than silently trusting a guess."""
    text = (lateral_anchor or "").lower()
    lw = lane_width_ft or 0.0
    if "shoulder" in text:
        if shoulder_width_ft is None:
            return PERP_HALF_LEN_FT, (
                f"lateralAnchor={lateral_anchor!r} references the shoulder but no "
                f"shoulder_width_ft was supplied -- fell back to the {PERP_HALF_LEN_FT}ft "
                f"reference-tick offset instead of a real shoulder-centered position")
        return lw + shoulder_width_ft / 2.0, None
    if "lane" in text:
        return lw / 2.0, None
    return PERP_HALF_LEN_FT, (
        f"lateralAnchor={lateral_anchor!r} did not match a recognized lane/shoulder "
        f"pattern -- fell back to the {PERP_HALF_LEN_FT}ft reference-tick offset")


def compile_symbols(spec: dict, resolved: dict, align_idx: int, segments,
                     outward_sign: float = -1.0, range_pick: str = "min",
                     lane_width_ft: float | None = None,
                     shoulder_width_ft: float | None = None,
                     tip_half_len_ft: float | None = None) -> list[dict]:
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
    vehicle's own computed position, not an independent post.

    tip_half_len_ft: perp tip distance for arrow-panel post base (default
    40 abstract; real-road = lane+shoulder from channelizing-line align).
    """
    import alignment_geometry as ag

    zone_range = _zone_station_ranges(spec, resolved, align_idx, range_pick)
    style = annotation_style(spec)
    symbol_label_out = float(style["offsetsFt"]["symbolLabel"])
    signs_mounted_on: dict[str, dict] = {
        s["mountedOn"]: s for s in spec["signs"]["items"]
        if s.get("postMounted") is False and s.get("mountedOn")
    }
    ap_tip = float(tip_half_len_ft) if tip_half_len_ft and tip_half_len_ft > 0 else PERP_HALF_LEN_FT

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
        angle_deg = math.degrees(math.atan2(tan_y, tan_x))
        kind = "protectiveVehicle" if is_vehicle else "arrowPanel"

        if is_arrow_panel:
            # Base = tip of the station perp (same attachment as roadside
            # signs), NOT the alignment center. Live miss 2026-08-10: base
            # at centerline put the stem/face/label ~half_len short of the
            # diamond-sign Y axis. tip_half_len_ft matches run_sheet_build /
            # resolve_sheet_lateral (40 abstract; lane+shoulder on real road).
            stem_gap_ft = 50.0
            panel_half_extent_guess_ft = 12.0
            base_x = x + out_x * ap_tip
            base_y = y + out_y * ap_tip
            label_dist = (ap_tip + stem_gap_ft
                          + panel_half_extent_guess_ft + symbol_label_out)
            prim = {
                "kind": kind,
                "id": item["id"], "cellName": item.get("cellHint") or "TWZAP_P",
                "x": base_x, "y": base_y, "dirX": out_x, "dirY": out_y,
                "angleDeg": angle_deg, "stationFt": sta,
                "requiredNote": item.get("required"),
                "primitiveId": _primitive_id(align_idx, item["id"], kind),
                "specRef": {
                    "zone": anchor.get("zone"), "run": None,
                    "symbolId": item["id"], "alignIdx": align_idx,
                },
            }
            primitives.append(prim)
            prim_by_id[item["id"]] = prim
            # Name beyond the cell (engineer reference style) — not a dim.
            primitives.append({
                "kind": "label",
                "x": x + out_x * label_dist,
                "y": y + out_y * label_dist,
                "text": "ARROW PANEL",
                "primitiveId": _primitive_id(align_idx, "arrowPanel", "label"),
                "specRef": {
                    "zone": anchor.get("zone"), "run": None,
                    "symbolId": "arrowPanel", "alignIdx": align_idx,
                },
            })
            continue

        offset_ft, offset_warning = _lateral_offset_ft(
            item.get("lateralAnchor"), lane_width_ft, shoulder_width_ft)
        px = x + out_x * offset_ft
        py = y + out_y * offset_ft
        prim = {
            "kind": kind,
            "id": item["id"], "cellName": item.get("cellHint") or "TWZAP_P",
            "x": px, "y": py, "angleDeg": angle_deg, "stationFt": sta,
            "requiredNote": item.get("required"),
            "lateralOffsetFt": offset_ft,
            "primitiveId": _primitive_id(align_idx, item["id"], kind),
            "specRef": {
                "zone": anchor.get("zone"), "run": None,
                "symbolId": item["id"], "alignIdx": align_idx,
            },
        }
        if offset_warning:
            prim["lateralOffsetWarning"] = offset_warning
        primitives.append(prim)
        prim_by_id[item["id"]] = prim

        if item["id"] in signs_mounted_on:
            sign = signs_mounted_on[item["id"]]
            primitives.append({
                "kind": "vehicleMountedSign", "signCode": sign["signCode"],
                "mountedOn": item["id"], "x": px, "y": py, "angleDeg": angle_deg,
                "primitiveId": _primitive_id(align_idx, sign["signCode"], "vehicleMountedSign"),
                "specRef": {
                    "zone": anchor.get("zone"), "run": None,
                    "signNum": sign["signCode"], "alignIdx": align_idx,
                },
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
                   outward_sign: float = -1.0,
                   work_bay_vertices: list | None = None) -> list[dict]:
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
    the roadway edge, not something this function silently papers over).

    work_bay_vertices: optional polyline from Align1 sta0 → Align2 sta0
    along the closed-lane / roadway edge. When provided (curved corridor),
    the hatch follows that path with local outward normals. When omitted,
    the hatch is the historical straight parallelogram between the two
    sta0 points (chord)."""
    import alignment_geometry as ag

    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    wa = zones.get("workArea")
    if not wa or not wa.get("hatched"):
        return []

    p1x, p1y, tan1x, tan1y = ag.station_to_xy(align1_segments, 0.0)
    p2x, p2y, _, _ = ag.station_to_xy(align2_segments, 0.0)

    spans_shoulder = "shoulder" in (wa.get("note") or "").lower()
    width = lane_width_ft + (shoulder_width_ft or 0.0) if spans_shoulder else lane_width_ft

    bay_segs = None
    if work_bay_vertices and len(work_bay_vertices) >= 2:
        try:
            bay_segs = ag.segments_from_polyline(work_bay_vertices)
        except ag.AlignmentGeometryError:
            bay_segs = None

    if bay_segs is not None and ag.total_length(bay_segs) >= 1.0:
        length = ag.total_length(bay_segs)
        # Densify enough that CreateShapeElement1 follows the lane curve
        # (6 pts on a 100 ft bay looked like a parallelogram).
        step = min(5.0, max(2.0, length / 40.0))
        n = max(4, int(math.ceil(length / step)))
        inner: list[tuple[float, float]] = []
        outer: list[tuple[float, float]] = []
        for i in range(n + 1):
            sta = length * (i / n)
            bx, by, btx, bty = ag.station_to_xy(bay_segs, sta)
            # Match Align1 outward: use −bay_travel as the tan basis.
            ox, oy = _outward_unit(-btx, -bty, outward_sign)
            inner.append((bx, by))
            outer.append((bx + ox * width, by + oy * width))
        boundary = list(inner) + list(reversed(outer))
        # Topology diagnostic still uses Align1 tan + chord to Align2 sta0.
        out_x, out_y = _outward_unit(tan1x, tan1y, outward_sign)

        def offset_pt(x: float, y: float, off: float) -> tuple[float, float]:
            return (x + out_x * off, y + out_y * off)
    else:
        length = math.hypot(p2x - p1x, p2y - p1y)
        out_x, out_y = _outward_unit(tan1x, tan1y, outward_sign)

        def offset_pt(x: float, y: float, off: float) -> tuple[float, float]:
            return (x + out_x * off, y + out_y * off)

        boundary = [offset_pt(p1x, p1y, 0.0), offset_pt(p2x, p2y, 0.0),
                    offset_pt(p2x, p2y, width), offset_pt(p1x, p1y, width)]

    # Diagnostic for sheet_rules.run_rules_gate's corridor-topology check
    # (see its own comment): where does align2's station-0 point (p2)
    # actually sit relative to align1's own line through p1? A correctly-
    # committed corridor has align2's edge off to the side somewhere, not
    # sitting on align1's own path -- if p2 projects onto align1's line at
    # a POSITIVE station (same direction align1's own stations increase)
    # and is close to that line (small perpendicular offset), align2 was
    # placed by walking further along align1's own corridor rather than at
    # a geometrically distinct work-area edge (confirmed live 2026-08-04:
    # this is exactly what an inverted-topology 619-311 build looked like --
    # Downstream offset +1000ft along the same blank line as Upstream).
    dx, dy = p2x - p1x, p2y - p1y
    p2_projected_station = dx * tan1x + dy * tan1y
    p2_perp_dist_ft = abs(dx * (-tan1y) + dy * tan1x)

    primitives: list[dict] = [{
        "kind": "hatch", "id": "workAreaHatch", "boundary": boundary,
        "workAreaLengthFt": length, "widthFt": width,
        "align2ProjectedStationOnAlign1Ft": p2_projected_station,
        "align2PerpDistFromAlign1Ft": p2_perp_dist_ft,
        "curvedWorkBay": bay_segs is not None and ag.total_length(bay_segs) >= 1.0,
        "primitiveId": "0:workArea:hatch",
        "specRef": {"zone": "workArea", "run": None, "alignIdx": 0},
    }]

    sym = next((s for s in spec["symbols"]["items"] if s["id"] == "channelizingDevices"), None)
    transverse = (sym or {}).get("transverse")
    if transverse and str(transverse.get("required")).lower() == "conditional":
        max_spacing = float(transverse.get("maxSpacingFt", 800))
        # Condition text is sheet-specific prose (e.g. "a paved shoulder >= 8'
        # closed for a distance greater than 800'"). All 15 sheet specs that
        # carry this clause phrase it as "8' or wider/greater" (confirmed via
        # grep across Data/sheet-specs), so minShoulderWidthFt defaults to 8.0
        # rather than parsing the prose; a spec can still override it
        # explicitly via a numeric "minShoulderWidthFt" key on the transverse
        # block. Both the length AND shoulder-width clauses must hold --
        # previously only length was checked, which fired transverse runs on
        # narrow-shoulder closures the sheet note doesn't actually cover.
        min_shoulder_ft = float(transverse.get("minShoulderWidthFt", 8.0))
        shoulder_wide_enough = (shoulder_width_ft or 0.0) >= min_shoulder_ft
        if length > max_spacing and shoulder_wide_enough:
            n_rows = int(length // max_spacing)
            for k in range(1, n_rows + 1):
                t = (k * max_spacing) / length
                if t >= 1.0:
                    break
                if bay_segs is not None and ag.total_length(bay_segs) >= 1.0:
                    sx, sy, stx, sty = ag.station_to_xy(bay_segs, k * max_spacing)
                    ox, oy = _outward_unit(-stx, -sty, outward_sign)
                    tip1 = (sx, sy)
                    tip2 = (sx + ox * width, sy + oy * width)
                else:
                    sx, sy = p1x + (p2x - p1x) * t, p1y + (p2y - p1y) * t
                    tip1 = offset_pt(sx, sy, 0.0)
                    tip2 = offset_pt(sx, sy, width)
                primitives.append({
                    "kind": "transverseRun", "run": "transverse",
                    "stationFromP1Ft": k * max_spacing,
                    "tip1": tip1, "tip2": tip2,
                    "note": transverse.get("sheetText"),
                    "primitiveId": _primitive_id(0, f"transverse{k}", "transverseRun"),
                    "specRef": {"zone": "workArea", "run": "transverse", "alignIdx": 0},
                })

    return primitives
