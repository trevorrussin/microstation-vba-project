"""Dashed white pavement elevation-break lines vs red AC at ACSR_P_T stations.

In the photogrammetry references, the pair of white dashed/dotted lines that
sandwich the red AC alignment live on elevation-break levels (PELB, BELB, PE,
PET, and other *EB / *ELB levels). At each station: resolve chainage on the AC
polyline, then measure a **same-source sandwich**: complete *EB pair if both
sides hit within MAX_SIDE_OFFSET_FT; else complete basemap Level 2 pair; else
fill a single missing side from the other source. Never mix sources per side
when either source already forms a full pair. PC/PT from ACSC_P_T attach to
the next ACSR_P_T station row at or above each event (round up). Curve PI labels
are omitted (tangency PIs only). Offset = ABS((north-south)/2) when both sides
hit; if only one side hits, offset = that side and bf_north_south = that side.
Output: station, point, north, south,
Offset From Alignment Best Fit Line to Photogrammetry (ft), bf_north_south.

Off-corridor PELB/RGCB fragments are excluded by a corridor lateral filter and
a parallel-to-AC check.
"""
from __future__ import annotations

import csv
import math
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
OFFSET_COL = "Offset From Alignment Best Fit Line to Photogrammetry (ft)"
sys.path.insert(0, str(ROOT / "mcp-server"))

import com_cache  # noqa: F401 — relocate pywin32 cache before COM generation
import pythoncom
from alignment_geometry import PathSegment, nearest_station, segments_from_polyline, station_to_xy, total_length
from ms_connect import _candidate_apps, get_microstation_app
from win32com.client import Dispatch

# Red alignment geometry: ACSRT_P in older files; level AC in 091026+.
AC_LEVEL_CANDIDATES = ("ACSRT_P", "AC")
# Station +50 ticks: ACSR_P_T historically; AC_C in 091026+.
STATION_LEVEL_CANDIDATES = ("ACSR_P_T", "AC_C")
# PC/PT events: ACSC_P_T historically; also embedded on AC_C in 091026+.
PC_PT_LEVEL_CANDIDATES = ("ACSC_P_T", "AC_C")
AC_LEVEL_NAME = "ACSRT_P"  # resolved at runtime via _resolve_ac_level()

STATION_CHAINAGE_OFFSET = 1950.0  # station_ft - chainage on AC (ch 0 at 19+50)
EARLY_CHAINAGE_OFFSET = 1950.0
CHAINAGE_FORMULA_SWITCH_FT = 0.0 # Use 1950 for everything
MIN_TRUSTED_CHAINAGE_FT = 0.0  # allow station 19+50 (ch 0) to use label XY
FIRST_STATION_FT = 1950.0
MAX_LABEL_PERP_FT = 150.0  # tie-break duplicate labels only; do not drop labels
DASHED_LEVELS = ("PELB", "BELB", "PE", "PET", "RGCB", "RG", "RGC", "RELB", "LELB")
BASEMAP_EXTRA_LEVELS = ("Level 2",)  # scanned in ALL attachments now
MAX_CORRIDOR_LATERAL_FT = 50.0
MAX_SIDE_OFFSET_FT = 8.0
# Allow measuring through dashed-line gaps: if the AC perp foot falls off a
# dash segment but within this along-track distance of it, use the infinite-
# line lateral offset (matches eye-measure through dash gaps).
MAX_DASH_ALONG_GAP_FT = 12.0
MIN_PARALLEL_COS = math.cos(math.radians(30.0))


def _app():
    matches: list = []
    any_apps: list = []
    for _open_file, app in _candidate_apps():
        any_apps.append(app)
        try:
            app.VBE.VBProjects("Test")
            matches.append(app)
        except Exception:
            continue
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        print(
            f"warning: {len(matches)} MicroStation instances with Test loaded; using first",
            flush=True,
        )
        return matches[0]
    if len(any_apps) == 1:
        return any_apps[0]
    if len(any_apps) > 1:
        print(
            f"warning: {len(any_apps)} MicroStation instances (no Test); using first",
            flush=True,
        )
        return any_apps[0]
    return get_microstation_app()


def _level_has_elements(mr, adf, level_name: str) -> bool:
    try:
        lv = adf.Levels(level_name)
    except Exception:
        return False
    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    o_scan.ExcludeAllLevels()
    o_scan.IncludeLevel(lv)
    ee = mr.Scan(o_scan)
    return bool(ee.MoveNext())


def _resolve_level(mr, adf, candidates: tuple[str, ...], kind: str) -> str:
    for name in candidates:
        if _level_has_elements(mr, adf, name):
            print(f"using {kind} level {name}", flush=True)
            return name
    raise RuntimeError(f"No {kind} elements on any of {candidates}")


def _parse_station_label(text: str) -> float | None:
    """Parse ##+## from bare ticks or prefixed labels (ET/WT/PC/PT …)."""
    s = text.strip().replace(" ", "")
    m = re.search(r"(\d+)\+(\d+(?:\.\d+)?)", s)
    if not m:
        return None
    return int(m.group(1)) * 100.0 + float(m.group(2))


def _format_station(sta_ft: float) -> str:
    major = int(sta_ft) // 100
    minor = sta_ft - major * 100
    if abs(minor - round(minor)) < 1e-9:
        return f"{major}+{int(round(minor)):02d}"
    # trim trailing zeros
    minor_s = f"{minor:.2f}".rstrip("0").rstrip(".")
    return f"{major}+{minor_s}"


def _label_side(text: str) -> str:
    """ET/WT (split EB/WB ticks) or empty for mainline +50 labels."""
    m = re.match(r"^\s*(ET|WT)\b", text.strip(), re.I)
    return m.group(1).upper() if m else ""


def _text_node_lines(el) -> list[str]:
    if el.IsTextElement:
        return [el.AsTextElement.Text.strip()]
    tn = el.AsTextNodeElement
    return [
        str(tn.TextLine(i)).strip()
        for i in range(1, int(tn.TextLinesCount) + 1)
    ]


def _transform_point(p, xf):
    return (
        xf.RowX.X * p.X + xf.RowX.Y * p.Y + xf.RowX.Z * p.Z + xf.TranslationX,
        xf.RowY.X * p.X + xf.RowY.Y * p.Y + xf.RowY.Z * p.Z + xf.TranslationY,
    )


def _attachment_label(att, ref_df) -> str:
    try:
        name = Path(ref_df.FullName).stem
        if name:
            return name
    except Exception:
        pass
    desc = (getattr(att, "Description", None) or "").strip()
    if desc:
        m = re.search(r"([^\\]+\.dgn)", desc, re.I)
        if m:
            return Path(m.group(1)).stem
    return (getattr(att, "Name", None) or "ref").replace(" ", "_")


def _working_attachments(mr) -> list[tuple[str, object, object]]:
    refs: list[tuple[str, object, object]] = []
    try:
        atts = mr.Attachments
        count = int(atts.Count)
    except Exception:
        return refs
    for i in range(1, count + 1):
        try:
            att = atts.Item(i).AsAttachment
            if not att.Name:
                continue
            ref_df = att.DesignFile
            refs.append((_attachment_label(att, ref_df), att, ref_df))
        except Exception:
            continue
    return refs


def _level_scan(mr, adf, level_name: str):
    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    o_scan.ExcludeAllLevels()
    try:
        lv = adf.Levels(level_name)
        o_scan.IncludeLevel(lv)
    except Exception:
        pass
    return mr.Scan(o_scan)


def _scan_level_segments(
    ref_mr,
    ref_df,
    level_name: str,
    xf,
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    try:
        lv = ref_df.Levels(level_name)
    except Exception:
        return []

    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    o_scan.ExcludeAllLevels()
    o_scan.IncludeLevel(lv)
    o_enum = ref_mr.Scan(o_scan)

    segs: list[tuple[tuple[float, float], tuple[float, float]]] = []
    while o_enum.MoveNext():
        el = o_enum.Current
        try:
            vl = el.AsVertexList
            pts = vl.GetVertices()
        except Exception:
            continue
        if len(pts) < 2:
            continue
        for j in range(len(pts) - 1):
            p0 = pts[j]
            p1 = pts[j + 1]
            if xf is not None:
                a = _transform_point(p0, xf)
                b = _transform_point(p1, xf)
            else:
                a = (float(p0.X), float(p0.Y))
                b = (float(p1.X), float(p1.Y))
            if math.hypot(b[0] - a[0], b[1] - a[1]) >= 0.01:
                segs.append((a, b))
    return segs


def _dist_point_to_ac(px: float, py: float, ac_segs: list[PathSegment]) -> float:
    best = float("inf")
    for seg in ac_segs:
        dx = seg.ex - seg.sx
        dy = seg.ey - seg.sy
        denom = dx * dx + dy * dy
        if denom < 1e-18:
            qx, qy = seg.sx, seg.sy
        else:
            u = ((px - seg.sx) * dx + (py - seg.sy) * dy) / denom
            u = max(0.0, min(1.0, u))
            qx = seg.sx + u * dx
            qy = seg.sy + u * dy
        best = min(best, math.hypot(qx - px, qy - py))
    return best


def _segment_parallel_to_ac(
    p0: tuple[float, float],
    p1: tuple[float, float],
    ac_segs: list[PathSegment],
) -> bool:
    mx = (p0[0] + p1[0]) / 2.0
    my = (p0[1] + p1[1]) / 2.0
    ch, _ = nearest_station(ac_segs, mx, my)
    _, _, tx, ty = station_to_xy(ac_segs, ch)
    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    ln = math.hypot(dx, dy)
    if ln < 1e-6:
        return False
    dot = abs((dx / ln) * tx + (dy / ln) * ty)
    return dot >= MIN_PARALLEL_COS


def _filter_corridor_segments(
    segs: list[tuple[tuple[float, float], tuple[float, float]]],
    ac_segs: list[PathSegment],
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    out: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for p0, p1 in segs:
        mx = (p0[0] + p1[0]) / 2.0
        my = (p0[1] + p1[1]) / 2.0
        if _dist_point_to_ac(mx, my, ac_segs) > MAX_CORRIDOR_LATERAL_FT:
            continue
        out.append((p0, p1))
    return out


def _segment_parallel_to_tangent(
    p0: tuple[float, float],
    p1: tuple[float, float],
    tx: float,
    ty: float,
) -> bool:
    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    ln = math.hypot(dx, dy)
    if ln < 1e-6:
        return False
    dot = abs((dx / ln) * tx + (dy / ln) * ty)
    return dot >= MIN_PARALLEL_COS


def _dashed_levels_for_ref(ref_df, ref_label: str = "") -> tuple[str, ...]:
    """Known sandwich levels plus any *EB / *ELB level in this reference."""
    names: set[str] = set(DASHED_LEVELS)
    if "basemap" in ref_label.lower():
        names.update(BASEMAP_EXTRA_LEVELS)
    try:
        count = int(ref_df.Levels.Count)
    except Exception:
        return tuple(sorted(names))
    for i in range(1, count + 1):
        try:
            level_name = ref_df.Levels.Item(i).Name
        except Exception:
            continue
        upper = level_name.upper()
        if upper.endswith("EB") or upper.endswith("ELB"):
            names.add(level_name)
    return tuple(sorted(names))


def _collect_dashed_segments(
    att,
    ref_df,
    ac_segs: list[PathSegment],
    ref_label: str = "",
) -> tuple[
    list[tuple[tuple[float, float], tuple[float, float]]],
    list[tuple[tuple[float, float], tuple[float, float]]],
    tuple[str, ...],
]:
    """Return (primary_eb_segs, fallback_basemap_segs, level_names).

    Primary = PELB/BELB/*EB parallel dashed sandwich lines.
    Fallback = basemap Level 2, also parallel-filtered — used only when
    primary misses a side at that station (east-end photo gaps).
    """
    xf = None
    try:
        xf = att.GetReferenceToMasterTransform()
    except Exception:
        xf = None
    ref_mr = ref_df.DefaultModelReference
    level_names = _dashed_levels_for_ref(ref_df, ref_label)
    primary: list[tuple[tuple[float, float], tuple[float, float]]] = []
    fallback: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for level_name in level_names:
        raw = _filter_corridor_segments(
            _scan_level_segments(ref_mr, ref_df, level_name, xf),
            ac_segs,
        )
        for seg in raw:
            if not _segment_parallel_to_ac(seg[0], seg[1], ac_segs):
                continue
            if level_name in BASEMAP_EXTRA_LEVELS:
                fallback.append(seg)
            else:
                primary.append(seg)
    return primary, fallback, level_names


def _side_offset(
    ax: float,
    ay: float,
    tx: float,
    ty: float,
    segs: list[tuple[tuple[float, float], tuple[float, float]]],
    side: str,
    max_dist: float,
) -> tuple[float, float, float] | None:
    """Nearest parallel dashed line on this side within max_dist of AC.

    Prefer the smallest lateral distance. On-dash feet and short dash-gap
    feet (MAX_DASH_ALONG_GAP_FT) compete equally — preferring on-dash only
    let a far Level 2 line beat a closer dashed gap at 212+50. Gap fill is
    capped so distant parallel segments cannot invent false offsets.
    """
    ln = math.hypot(tx, ty)
    if ln < 1e-12:
        return None
    tx /= ln
    ty /= ln
    nx, ny = -ty, tx

    best_on = float("inf")
    best_on_q = (ax, ay)
    best_gap = float("inf")
    best_gap_q = (ax, ay)

    for p0, p1 in segs:
        if not _segment_parallel_to_tangent(p0, p1, tx, ty):
            continue
        dx = p1[0] - p0[0]
        dy = p1[1] - p0[1]
        denom = dx * dx + dy * dy
        if denom < 1e-18:
            continue
        seg_len = math.sqrt(denom)
        u_raw = ((ax - p0[0]) * dx + (ay - p0[1]) * dy) / denom
        if u_raw < 0.0:
            along_off = abs(u_raw) * seg_len
        elif u_raw > 1.0:
            along_off = (u_raw - 1.0) * seg_len
        else:
            along_off = 0.0

        qx = p0[0] + u_raw * dx
        qy = p0[1] + u_raw * dy
        along_left = (qx - ax) * nx + (qy - ay) * ny
        if side == "left" and along_left <= 0.05:
            continue
        if side == "right" and along_left >= -0.05:
            continue
        dist = abs(along_left)
        if dist > max_dist:
            continue

        if along_off <= 1e-6:
            if dist < best_on:
                best_on = dist
                best_on_q = (qx, qy)
        elif along_off <= MAX_DASH_ALONG_GAP_FT and dist < best_gap:
            best_gap = dist
            best_gap_q = (qx, qy)

    if best_on != float("inf") and best_gap != float("inf"):
        # Prefer on-dash when nearly equal; gap must be clearly closer.
        if best_gap + 0.15 < best_on:
            return best_gap, best_gap_q[0], best_gap_q[1]
        return best_on, best_on_q[0], best_on_q[1]
    if best_on != float("inf"):
        return best_on, best_on_q[0], best_on_q[1]
    if best_gap != float("inf"):
        return best_gap, best_gap_q[0], best_gap_q[1]
    return None


def _measure_sides(
    sx: float,
    sy: float,
    tx: float,
    ty: float,
    primary_segs: list[tuple[tuple[float, float], tuple[float, float]]],
    fallback_segs: list[tuple[tuple[float, float], tuple[float, float]]],
) -> tuple[float | None, float | None]:
    """Measure N/S as one sandwich from the same source.

    1. Complete primary (*EB) pair → use both primary sides.
    2. Else complete basemap Level 2 pair → use both fallback sides.
    3. Else fill any remaining miss from the other source (photo gaps).

    Per-side closest-wins mixed Level 2 north with EB south (and the reverse)
    and broke stations like 241+50 / 210+50.
    """
    def _one(segs: list, side: str) -> float | None:
        hit = _side_offset(sx, sy, tx, ty, segs, side, MAX_SIDE_OFFSET_FT)
        return None if hit is None else round(hit[0], 3)

    pL = _one(primary_segs, "left")
    pR = _one(primary_segs, "right")
    fL = _one(fallback_segs, "left") if fallback_segs else None
    fR = _one(fallback_segs, "right") if fallback_segs else None

    if pL is not None and pR is not None:
        return pL, pR
    if fL is not None and fR is not None:
        return fL, fR
    return (pL if pL is not None else fL), (pR if pR is not None else fR)


def _offsets_at_chainage(
    chainage: float,
    ac_segs: list[PathSegment],
    physical_segs: list[PathSegment],
    primary_segs: list[tuple[tuple[float, float], tuple[float, float]]],
    fallback_segs: list[tuple[tuple[float, float], tuple[float, float]]],
) -> tuple[float | None, float | None]:
    """Measure both sides at the same XY. Never mix station_xy with physical snap."""
    sx, sy, tx, ty = station_to_xy(ac_segs, chainage)
    left, right = _measure_sides(sx, sy, tx, ty, primary_segs, fallback_segs)
    if left is not None and right is not None:
        return left, right
    # Only relocate when BOTH sides miss — one-sided fill from a different
    # XY (physical AC) caused bogus 200+ station asymmetries.
    if left is None and right is None and physical_segs:
        ax, ay, ptx, pty = _nearest_on_physical_ac(sx, sy, physical_segs)
        return _measure_sides(ax, ay, ptx, pty, primary_segs, fallback_segs)
    return left, right


def _tessellate_arc(arc) -> list[tuple[float, float]]:
    """Approximate an arc with line segments."""
    pts = []
    try:
        # Get start/end/center/radius/angles
        # This is a bit complex via COM, so we'll use a simpler approach:
        # Get 10 points along the arc using the COM API if possible
        # Or just use start and end for now if it's a small arc
        p0 = arc.StartPoint
        p1 = arc.EndPoint
        # For alignment purposes, even 2-3 segments per arc is often enough
        # if the arcs are large radius (highway curves).
        # Let's try to get the full vertex list if MicroStation supports it for arcs.
        try:
            return [(float(p.X), float(p.Y)) for p in arc.AsVertexList.GetVertices()]
        except:
            return [(float(p0.X), float(p0.Y)), (float(p1.X), float(p1.Y))]
    except:
        return []


def _ac_geometry_cloud() -> list[PathSegment]:
    """All AC / ACSRT_P edges as independent segments (no cross-element chaining)."""
    pythoncom.CoInitialize()
    try:
        app = _app()
        adf = app.ActiveDesignFile
        mr = adf.DefaultModelReference
        level_name = _resolve_level(mr, adf, AC_LEVEL_CANDIDATES, "AC-geometry")
        try:
            lv = adf.Levels(level_name)
        except Exception:
            return []

        o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
        o_scan.ExcludeNonGraphical()
        o_scan.IncludeLevel(lv)
        ee = mr.Scan(o_scan)

        polylines: list[list[tuple[float, float]]] = []

        def _add_el(e) -> None:
            try:
                if e.Type == 16:
                    pts = _tessellate_arc(e.AsArcElement)
                else:
                    pts = e.AsVertexList.GetVertices()
                    pts = [(float(p.X), float(p.Y)) for p in pts]
                if len(pts) >= 2:
                    polylines.append(pts)
            except Exception:
                return

        while ee.MoveNext():
            el = ee.Current
            if el.Type not in (3, 4, 12, 16):
                continue
            try:
                if el.Type == 12:
                    sub_ee = el.AsComplexStringElement.GetSubElements()
                    while sub_ee.MoveNext():
                        _add_el(sub_ee.Current)
                else:
                    _add_el(el)
            except Exception:
                continue

        segs: list[PathSegment] = []
        for pts in polylines:
            for i in range(len(pts) - 1):
                p0, p1 = pts[i], pts[i + 1]
                ln = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
                if ln < 1e-6:
                    continue
                segs.append(
                    PathSegment(
                        is_arc=False,
                        sx=p0[0],
                        sy=p0[1],
                        sz=0.0,
                        ex=p1[0],
                        ey=p1[1],
                        ez=0.0,
                        seg_len=ln,
                    )
                )
        return segs
    finally:
        pythoncom.CoUninitialize()


def _nearest_point_on_cloud(
    px: float,
    py: float,
    cloud: list[PathSegment],
) -> tuple[float, float, float, float, float]:
    """Return qx, qy, dist, tx, ty for nearest point on cloud."""
    best_dist = float("inf")
    best_x, best_y = px, py
    best_tx, best_ty = 1.0, 0.0
    for seg in cloud:
        dx = seg.ex - seg.sx
        dy = seg.ey - seg.sy
        denom = dx * dx + dy * dy
        if denom < 1e-18:
            qx, qy = seg.sx, seg.sy
            ln = 0.0
        else:
            u = ((px - seg.sx) * dx + (py - seg.sy) * dy) / denom
            u = max(0.0, min(1.0, u))
            qx = seg.sx + u * dx
            qy = seg.sy + u * dy
            ln = math.sqrt(denom)
        dist = math.hypot(qx - px, qy - py)
        if dist < best_dist:
            best_dist = dist
            best_x, best_y = qx, qy
            if ln >= 1e-12:
                best_tx, best_ty = dx / ln, dy / ln
    return best_x, best_y, best_dist, best_tx, best_ty


def _chain_ac_by_station_labels(
    cloud: list[PathSegment],
    label_xy: dict[float, tuple[float, float]],
    max_snap_ft: float = 250.0,
) -> list[PathSegment]:
    """Build one west→east AC path snapped to station labels in station order.

    Avoids greedy endpoint chaining that doubles back through EB/WB forks and
    inflates length (seen ~71k ft on a ~26k ft corridor in 091026).
    """
    if not cloud or not label_xy:
        return []
    ordered = sorted(label_xy.items(), key=lambda kv: kv[0])
    snapped: list[tuple[float, float]] = []
    for sta, (lx, ly) in ordered:
        qx, qy, dist, _tx, _ty = _nearest_point_on_cloud(lx, ly, cloud)
        if dist > max_snap_ft:
            print(f"warning: station {sta:g} AC snap {dist:.1f} ft — skipped", flush=True)
            continue
        if snapped:
            prev = snapped[-1]
            if math.hypot(qx - prev[0], qy - prev[1]) < 0.05:
                continue
            # Densify with cloud vertices near the chord, ordered along chord.
            ax, ay = prev
            bx, by = qx, qy
            chord = math.hypot(bx - ax, by - ay)
            if chord > 1.0:
                ux, uy = (bx - ax) / chord, (by - ay) / chord
                extras: list[tuple[float, float, float]] = []
                for seg in cloud:
                    for px, py in ((seg.sx, seg.sy), (seg.ex, seg.ey)):
                        t = (px - ax) * ux + (py - ay) * uy
                        if t <= 1.0 or t >= chord - 1.0:
                            continue
                        # lateral distance to chord
                        lat = abs((px - ax) * (-uy) + (py - ay) * ux)
                        if lat <= 40.0:
                            extras.append((t, px, py))
                extras.sort()
                # thin extras to ~25 ft
                last_t = 0.0
                for t, px, py in extras:
                    if t - last_t >= 25.0:
                        snapped.append((px, py))
                        last_t = t
        snapped.append((qx, qy))

    segs: list[PathSegment] = []
    for i in range(len(snapped) - 1):
        p0, p1 = snapped[i], snapped[i + 1]
        ln = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
        if ln < 1e-6:
            continue
        segs.append(
            PathSegment(
                is_arc=False,
                sx=p0[0],
                sy=p0[1],
                sz=0.0,
                ex=p1[0],
                ey=p1[1],
                ez=0.0,
                seg_len=ln,
            )
        )
    return segs


def _physical_ac_segments() -> list[PathSegment]:
    """Deprecated wrapper — prefer cloud + station-guided chain in main()."""
    return _ac_geometry_cloud()


def _nearest_on_physical_ac(
    px: float,
    py: float,
    physical_segs: list[PathSegment],
) -> tuple[float, float, float, float]:
    best_dist = float("inf")
    best_x = px
    best_y = py
    best_tx = 1.0
    best_ty = 0.0
    for seg in physical_segs:
        dx = seg.ex - seg.sx
        dy = seg.ey - seg.sy
        denom = dx * dx + dy * dy
        if denom < 1e-18:
            qx, qy = seg.sx, seg.sy
        else:
            u = ((px - seg.sx) * dx + (py - seg.sy) * dy) / denom
            u = max(0.0, min(1.0, u))
            qx = seg.sx + u * dx
            qy = seg.sy + u * dy
        dist = math.hypot(qx - px, qy - py)
        if dist < best_dist:
            best_dist = dist
            best_x = qx
            best_y = qy
            ln = math.hypot(dx, dy)
            if ln >= 1e-12:
                best_tx = dx / ln
                best_ty = dy / ln
    return best_x, best_y, best_tx, best_ty


def _label_origin(el) -> tuple[float, float] | None:
    try:
        o = el.AsTextElement.Origin
        return float(o.X), float(o.Y)
    except Exception:
        pass
    try:
        o = el.AsTextNodeElement.Origin
        return float(o.X), float(o.Y)
    except Exception:
        return None


def _chainage_for_station(
    sta_ft: float,
    label_x: float | None,
    label_y: float | None,
    ac_segs: list[PathSegment],
    ac_len: float,
) -> tuple[float, str]:
    """Resolve chainage on the AC stationing polyline for one station row."""
    if label_x is not None and label_y is not None:
        ch, perp = nearest_station(ac_segs, label_x, label_y)
        if perp <= MAX_LABEL_PERP_FT and ch >= MIN_TRUSTED_CHAINAGE_FT:
            return min(ch, ac_len), "label"
    if sta_ft >= CHAINAGE_FORMULA_SWITCH_FT:
        return min(max(0.0, sta_ft - STATION_CHAINAGE_OFFSET), ac_len), "formula2700"
    return min(max(0.0, sta_ft - EARLY_CHAINAGE_OFFSET), ac_len), "formula1950"


def _read_acsc_events() -> list[tuple[str, str, float, float | None, float | None]]:
    """PC and PT station labels (ACSC_P_T or AC_C single-line 'PC 19+50')."""
    pythoncom.CoInitialize()
    try:
        app = _app()
        adf = app.ActiveDesignFile
        mr = adf.DefaultModelReference
        level_name = _resolve_level(mr, adf, PC_PT_LEVEL_CANDIDATES, "PC/PT")
        ee = _level_scan(mr, adf, level_name)
        events: list[tuple[str, str, float, float | None, float | None]] = []
        while ee.MoveNext():
            el = ee.Current
            try:
                if not (el.IsTextElement or el.IsTextNodeElement):
                    continue
            except Exception:
                continue
            lines = _text_node_lines(el)
            if not lines:
                continue
            kind = ""
            sta_text = ""
            if len(lines) == 1 and "|" in lines[0]:
                kind, sta_text = [part.strip() for part in lines[0].split("|", 1)]
            elif len(lines) >= 2:
                kind, sta_text = lines[0], lines[1]
            else:
                # Single line e.g. "PC 19+50.00" or "PT 26+98.41"
                m = re.match(
                    r"^\s*(PC|PT)\s+(\d+\+\d+(?:\.\d+)?)\s*$",
                    lines[0],
                    re.I,
                )
                if not m:
                    continue
                kind, sta_text = m.group(1), m.group(2)
            kind = kind.upper().replace(".", "")
            if kind not in ("PC", "PT"):
                continue
            sta_ft = _parse_station_label(sta_text)
            if sta_ft is None:
                continue
            label = sta_text.strip().replace(" ", "")
            if label.endswith(".00"):
                label = label[:-3]
            origin = _label_origin(el)
            x, y = (origin if origin else (None, None))
            events.append((kind, label, sta_ft, x, y))
    finally:
        pythoncom.CoUninitialize()
    events.sort(key=lambda row: row[2])
    return events


def _build_output_stations(
    labels: list[tuple[str, float]],
    events: list[tuple[str, str, float, float | None, float | None]],
) -> list[tuple[str, str, float]]:
    """Merge mainline ticks + ET/WT split ticks + exact PC/PT events."""
    # Key by (sta, side) so ET and WT never collapse onto one row.
    rows: dict[tuple[float, str], tuple[str, str]] = {}

    for lab, sta in labels:
        side = _label_side(lab)
        display = _format_station(sta)
        rows[(sta, side)] = (display, side)

    for kind, lab, sta, _x, _y in events:
        if kind not in ("PC", "PT"):
            continue
        display = lab if lab else _format_station(sta)
        rows[(sta, kind)] = (display, kind)

    return sorted(
        ((lab, point, sta) for (sta, _side), (lab, point) in rows.items()),
        key=lambda row: (row[2], row[1]),
    )

def _photogrammetry_offset(north: float | None, south: float | None) -> float | None:
    if north is not None and south is not None:
        return round(abs((north - south) / 2), 3)
    if north is not None:
        return round(north, 3)
    if south is not None:
        return round(south, 3)
    return None


def _bf_north_south_from_sides(north: float | None, south: float | None) -> str:
    if north is not None and south is not None:
        half = (north - south) / 2
        if half < 0:
            return "south"
        if half > 0:
            return "north"
        return ""
    if north is not None:
        return "north"
    if south is not None:
        return "south"
    return ""


def _read_station_labels() -> list[tuple[str, float, float, float]]:
    pythoncom.CoInitialize()
    try:
        app = _app()
        adf = app.ActiveDesignFile
        mr = adf.DefaultModelReference
        level_name = _resolve_level(mr, adf, STATION_LEVEL_CANDIDATES, "station-label")
        ee = _level_scan(mr, adf, level_name)
        labels: list[tuple[str, float, float, float]] = []
        while ee.MoveNext():
            el = ee.Current
            try:
                if not (el.IsTextElement or el.IsTextNodeElement):
                    continue
            except Exception:
                continue
            try:
                lines = _text_node_lines(el)
                content = lines[0] if lines else ""
            except Exception:
                content = ""
            # Skip PC/PT event strings (handled via _read_acsc_events).
            # Keep ET/WT — those are the ~100 ft ticks on both sides east of Havemeyer.
            if re.match(r"^\s*(PC|PT)\b", content, re.I):
                continue
            sta = _parse_station_label(content)
            if sta is None:
                continue
            origin = _label_origin(el)
            if origin is None:
                continue
            labels.append((content, sta, origin[0], origin[1]))
    finally:
        pythoncom.CoUninitialize()
    labels.sort(key=lambda row: row[1])
    return labels


def _label_xy_map(
    labels: list[tuple[str, float, float, float]],
    events: list[tuple[str, str, float, float | None, float | None]],
    ac_segs: list[PathSegment],
) -> dict[float, tuple[float, float]]:
    """Map station ft -> label XY; prefer lowest-perp duplicate at same station."""
    best_perp: dict[float, float] = {}
    out: dict[float, tuple[float, float]] = {}

    def _register(sta: float, x: float, y: float) -> None:
        _, perp = nearest_station(ac_segs, x, y)
        if sta not in out or perp < best_perp[sta]:
            out[sta] = (x, y)
            best_perp[sta] = perp

    for _lab, sta, x, y in labels:
        _register(sta, x, y)
    for _kind, _lab, sta, x, y in events:
        if x is None or y is None:
            continue
        _register(sta, x, y)

    return out


def _interpolate_blanks(values: list[float | None], max_gap: int = 25) -> list[float | None]:
    out = list(values)
    n = len(out)
    i = 0
    while i < n:
        if out[i] is not None:
            i += 1
            continue
        j = i
        while j < n and out[j] is None:
            j += 1
        if i > 0 and j < n and (j - i) <= max_gap:
            v0 = out[i - 1]
            v1 = out[j]
            assert v0 is not None and v1 is not None
            span = j - i
            for k in range(span):
                t = (k + 1) / (span + 1)
                out[i + k] = round(v0 + t * (v1 - v0), 3)
        i = max(j, i + 1)
    return out


def main() -> Path:
    labels = _read_station_labels()
    if not labels:
        raise RuntimeError("No station labels found on ACSR_P_T or AC_C")

    acsc_events = _read_acsc_events()
    print(
        f"PC/PT events: pc={sum(1 for k, _, _, _, _ in acsc_events if k == 'PC')} "
        f"pt={sum(1 for k, _, _, _, _ in acsc_events if k == 'PT')}",
        flush=True,
    )

    labels = [(lab, sta, x, y) for lab, sta, x, y in labels if sta >= FIRST_STATION_FT]
    if not labels:
        raise RuntimeError(f"No stations at or after {FIRST_STATION_FT:g} ft")

    # AC geometry cloud + station-guided chain (avoids EB/WB fork double-backs).
    # Seed path with mainline ticks + ET only — WT sits on the other split
    # branch and would zigzag the chain if included.
    ac_cloud = _ac_geometry_cloud()
    if not ac_cloud:
        raise RuntimeError("No AC geometry found on ACSRT_P or AC")
    seed_xy: dict[float, tuple[float, float]] = {}
    et_n = wt_n = main_n = 0
    for lab, sta, x, y in labels:
        side = _label_side(lab)
        if side == "WT":
            wt_n += 1
            continue
        if side == "ET":
            et_n += 1
        else:
            main_n += 1
        seed_xy[sta] = (x, y)
    print(f"labels main={main_n} ET={et_n} WT={wt_n} (WT excluded from AC chain)", flush=True)
    ac_segs = _chain_ac_by_station_labels(ac_cloud, seed_xy)
    if not ac_segs:
        raise RuntimeError("Failed to chain AC along station labels")
    physical_segs = ac_cloud  # full geometry for optional both-miss snap
    ac_len = total_length(ac_segs)
    print(
        f"AC cloud edges={len(ac_cloud)} chained_len={ac_len:.1f} ft "
        f"label_span={max(seed_xy)-min(seed_xy):.1f} sta-ft",
        flush=True,
    )
    label_xy = _label_xy_map(labels, acsc_events, ac_segs)
    print(f"station labels with XY: {len(label_xy)}", flush=True)

    pythoncom.CoInitialize()
    try:
        mr = _app().ActiveDesignFile.DefaultModelReference
        ref_attachments = _working_attachments(mr)
    finally:
        pythoncom.CoUninitialize()

    primary_all: list[tuple[tuple[float, float], tuple[float, float]]] = []
    fallback_all: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for ref_label, att, ref_df in ref_attachments:
        pythoncom.CoInitialize()
        try:
            primary, fallback, level_names = _collect_dashed_segments(
                att, ref_df, ac_segs, ref_label,
            )
        finally:
            pythoncom.CoUninitialize()
        print(
            f"scanning {ref_label} levels={','.join(level_names)} "
            f"primary={len(primary)} fallback={len(fallback)} …",
            flush=True,
        )
        primary_all.extend(primary)
        fallback_all.extend(fallback)

    if not primary_all and not fallback_all:
        raise RuntimeError("No in-corridor dashed linework found in references")

    out_dir = ROOT / "outputs" / "727-alignment-offset"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"dashed_sandwich_offset_{stamp}.csv"
    xlsx_path = out_dir / f"dashed_sandwich_offset_{stamp}.xlsx"

    header = ["station", "point", "north", "south", OFFSET_COL, "bf_north_south"]
    label_rows = [(lab, sta) for lab, sta, _x, _y in labels]
    output_stations = _build_output_stations(label_rows, acsc_events)
    station_rows: list[tuple[float | None, float | None]] = []
    chainage_methods: dict[str, int] = {}
    missing_both = 0
    missing_one = 0
    for _station_label, _point, station_ft in output_stations:
        lx = ly = None
        if station_ft in label_xy:
            lx, ly = label_xy[station_ft]
        chainage, method = _chainage_for_station(station_ft, lx, ly, ac_segs, ac_len)
        chainage_methods[method] = chainage_methods.get(method, 0) + 1
        left, right = _offsets_at_chainage(
            chainage, ac_segs, physical_segs, primary_all, fallback_all,
        )
        if left is None and right is None:
            missing_both += 1
        elif left is None or right is None:
            missing_one += 1
        station_rows.append((left, right))
    print(
        f"chainage: {chainage_methods}  missing_both={missing_both} missing_one={missing_one}",
        flush=True,
    )

    records = [
        (output_stations[i][0], output_stations[i][1], station_rows[i][0], station_rows[i][1])
        for i in range(len(output_stations))
    ]

    filled_both = sum(
        1 for _, _, left, right in records if left is not None and right is not None
    )
    point_rows = sum(1 for _, point, _, _ in records if point)
    over_max = sum(
        1 for _, _, left, right in records
        if (left is not None and left > MAX_SIDE_OFFSET_FT)
        or (right is not None and right > MAX_SIDE_OFFSET_FT)
    )

    with csv_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(header)
        for station_label, point, left, right in records:
            offset = _photogrammetry_offset(left, right)
            direction = _bf_north_south_from_sides(left, right)
            writer.writerow([
                station_label,
                point,
                "" if left is None else left,
                "" if right is None else right,
                "" if offset is None else offset,
                direction,
            ])

    try:
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "AC vs dashed"
        ws.append(header)
        for row_idx, (station_label, point, left, right) in enumerate(records, start=2):
            ws.append([
                station_label,
                point,
                left,
                right,
                None,
                None,
            ])
            ws[f"E{row_idx}"] = (
                f'=IF(C{row_idx}="",IF(D{row_idx}="","",D{row_idx}),'
                f'IF(D{row_idx}="",C{row_idx},ABS((C{row_idx}-D{row_idx})/2)))'
            )
            ws[f"F{row_idx}"] = (
                f'=IF(AND(C{row_idx}<>"",D{row_idx}<>""),'
                f'IF((C{row_idx}-D{row_idx})/2<0,"south",IF((C{row_idx}-D{row_idx})/2>0,"north","")),'
                f'IF(C{row_idx}<>"","north",IF(D{row_idx}<>"","south","")))'
            )
        wb.save(xlsx_path)
    except ImportError:
        xlsx_path = None

    first = records[0]
    first_offset = _photogrammetry_offset(first[2], first[3])
    first_dir = _bf_north_south_from_sides(first[2], first[3])
    print(
        f"rows={len(records)} both_sides={filled_both} point_rows={point_rows} "
        f"over_{MAX_SIDE_OFFSET_FT:g}ft={over_max} "
        f"first={first[0]} point={first[1]} north={first[2]} south={first[3]} "
        f"offset={first_offset} bf={first_dir} ac_len={ac_len:.1f}"
    )
    print(f"csv={csv_path}")
    if xlsx_path:
        print(f"xlsx={xlsx_path}")
    return csv_path


if __name__ == "__main__":
    main()
