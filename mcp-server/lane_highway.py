"""Plan-view road striping geometry for WZTC general CAD.

Patterns (engineer live 2026-08-10 + sheet cross-sections):

One-way (`lane_highway_lines`) — Family 2/3 single carriageway core:
  2 solid white outer edges + (lanes-1) dashed white separators.

Two-way undivided (`two_way_highway_lines`) — Family 1 / 311-style:
  2 solid white outsides + 2 solid yellow center (yellow_gap_ft apart)
  + (L-1) dashed white each side (L = lanes/2). Even total lanes only.

Divided / freeway dual carriageway (`divided_highway_lines`) — 302-style:
  Each direction: solid white outer, (N-1) dashed white, solid yellow
  median edge. Median gap = median_width_ft between the two yellows.
  Optional outer shoulders.

TWLT undivided (`twlt_highway_lines`) — 312/412-style:
  Outer white + (L-1) dashed white each direction; center TWLT lane
  bounded by two dashed yellow lines twlt_width_ft apart.
  L = lanes_per_direction (travel lanes each way; TWLT not counted).

Optional shoulders: solid white EOP lines outside each travel outer edge
at shoulder_width_ft (sheet “paved shoulder outside outer travel lane”).

Dash pattern is real gaps (dash_ft solid / gap_ft empty), not a MS linestyle.

Path: every builder accepts either a straight first-travel-outer edge
`(x1,y1)->(x2,y2)` or a multi-point `vertices=[[x,y],…]` polyline (S-curves,
polylines). Offset uses local path tangent normals via alignment_geometry.
"""
from __future__ import annotations

import math
from typing import Literal, Sequence

from alignment_geometry import PathSegment, station_to_xy, total_length


def _corridor_frame(
    x1: float, y1: float, x2: float, y2: float, side: Literal["right", "left"],
) -> tuple[float, float, float, float, float]:
    dx = x2 - x1
    dy = y2 - y1
    length = math.hypot(dx, dy)
    if length < 1e-6:
        raise ValueError("corridor length must be > 0 (x1,y1) and (x2,y2) coincide")
    tx, ty = dx / length, dy / length
    nx, ny = _side_normal(tx, ty, side)
    return length, tx, ty, nx, ny


def _side_normal(
    tx: float, ty: float, side: Literal["right", "left"],
) -> tuple[float, float]:
    if side == "left":
        return -ty, tx
    if side == "right":
        return ty, -tx
    raise ValueError("side must be 'right' or 'left'")


def vertices_to_path_segments(
    vertices: Sequence[Sequence[float]],
) -> list[PathSegment]:
    """Build straight PathSegments from an ordered first-travel-outer polyline."""
    if vertices is None or len(vertices) < 2:
        raise ValueError("vertices must have at least 2 points [[x,y], …]")
    segs: list[PathSegment] = []
    for i in range(len(vertices) - 1):
        p0, p1 = vertices[i], vertices[i + 1]
        sx, sy = float(p0[0]), float(p0[1])
        ex, ey = float(p1[0]), float(p1[1])
        seg_len = math.hypot(ex - sx, ey - sy)
        if seg_len < 1e-9:
            continue
        segs.append(PathSegment(
            False, sx, sy, 0.0, ex, ey, 0.0, seg_len,
        ))
    if not segs:
        raise ValueError("corridor length must be > 0 (degenerate vertices)")
    return segs


def resolve_edge_path(
    x1: float, y1: float, x2: float, y2: float,
    vertices: Sequence[Sequence[float]] | None = None,
) -> list[PathSegment]:
    """Prefer multi-point vertices when provided; else two-point edge."""
    if vertices is not None:
        return vertices_to_path_segments(vertices)
    return vertices_to_path_segments([[x1, y1], [x2, y2]])


def path_length(segments: list[PathSegment]) -> float:
    return total_length(segments)


def fillet_polyline_segments(
    segments: list[PathSegment],
    radius_ft: float,
) -> list[PathSegment]:
    """Replace sharp corners with circular fillets (continuous curve path).

    Each interior vertex becomes line → arc → line. Skips a corner when the
    radius will not fit on either adjacent leg. radius_ft <= 0 leaves the
    path unchanged.
    """
    if radius_ft <= 1e-9 or len(segments) < 2:
        return list(segments)

    # Work from vertex list of the straight polyline only
    if any(s.is_arc for s in segments):
        return list(segments)

    verts = [[segments[0].sx, segments[0].sy]]
    for seg in segments:
        verts.append([seg.ex, seg.ey])

    out: list[PathSegment] = []
    n = len(verts)
    # trim_start[i] = how much to trim from start of edge i (edge from verts[i] to verts[i+1])
    trim_start = [0.0] * (n - 1)
    trim_end = [0.0] * (n - 1)
    arcs: dict[int, PathSegment] = {}  # keyed by vertex index

    for i in range(1, n - 1):
        x0, y0 = verts[i - 1]
        x1, y1 = verts[i]
        x2, y2 = verts[i + 1]
        l0 = math.hypot(x1 - x0, y1 - y0)
        l1 = math.hypot(x2 - x1, y2 - y1)
        if l0 < 1e-9 or l1 < 1e-9:
            continue
        t0x, t0y = (x1 - x0) / l0, (y1 - y0) / l0
        t1x, t1y = (x2 - x1) / l1, (y2 - y1) / l1
        cross = t0x * t1y - t0y * t1x
        dot = max(-1.0, min(1.0, t0x * t1x + t0y * t1y))
        phi = math.atan2(cross, dot)
        if abs(phi) < 1e-4:
            continue  # nearly colinear
        half = abs(phi) * 0.5
        inset = radius_ft * math.tan(half)
        if inset + 1e-6 > l0 - trim_start[i - 1] or inset + 1e-6 > l1:
            continue  # won't fit
        if inset + 1e-6 > l0 or inset + 1e-6 > l1 - (trim_end[i] if False else 0):
            pass
        # Reserve trim on both edges
        if inset > l0 - trim_start[i - 1] - 1e-6:
            continue
        if inset > l1 - 1e-6:
            continue
        trim_end[i - 1] = inset
        trim_start[i] = inset

        psx = x1 - t0x * inset
        psy = y1 - t0y * inset
        pex = x1 + t1x * inset
        pey = y1 + t1y * inset
        if phi > 0:
            nx, ny = -t0y, t0x  # left of incoming
        else:
            nx, ny = t0y, -t0x  # right of incoming
        cx, cy = psx + nx * radius_ft, psy + ny * radius_ft
        sa = math.atan2(psy - cy, psx - cx)
        ea = math.atan2(pey - cy, pex - cx)
        sweep = ea - sa
        # Normalize sweep to match turn direction
        if phi > 0:  # CCW sweep should be positive
            while sweep <= 0:
                sweep += 2 * math.pi
            while sweep > 2 * math.pi:
                sweep -= 2 * math.pi
        else:
            while sweep >= 0:
                sweep -= 2 * math.pi
            while sweep < -2 * math.pi:
                sweep += 2 * math.pi
        arc_len = abs(sweep) * radius_ft
        arcs[i] = PathSegment(
            True, psx, psy, 0.0, pex, pey, 0.0, arc_len,
            cx=cx, cy=cy, radius=radius_ft,
            start_angle=sa, sweep_angle=sweep,
        )

    # Emit trimmed straights + arcs
    for i in range(n - 1):
        x0, y0 = verts[i]
        x1, y1 = verts[i + 1]
        dx, dy = x1 - x0, y1 - y0
        L = math.hypot(dx, dy)
        if L < 1e-9:
            continue
        tx, ty = dx / L, dy / L
        a0 = trim_start[i]
        a1 = L - trim_end[i]
        if a1 - a0 > 1e-6:
            sx, sy = x0 + tx * a0, y0 + ty * a0
            ex, ey = x0 + tx * a1, y0 + ty * a1
            out.append(PathSegment(
                False, sx, sy, 0.0, ex, ey, 0.0, a1 - a0,
            ))
        # Arc after this edge when this edge ends at a filleted vertex
        v_end = i + 1
        if v_end in arcs:
            out.append(arcs[v_end])
    return out if out else list(segments)


def _append_solid(
    out: list[dict], *, kind: str, row: int,
    sx: float, sy: float, ex: float, ey: float,
    vertices: list[list[float]] | None = None,
) -> None:
    seg: dict = {
        "style": "solid", "kind": kind, "row": row,
        "x1": sx, "y1": sy, "x2": ex, "y2": ey,
    }
    if vertices is not None and len(vertices) >= 2:
        seg["vertices"] = vertices
    out.append(seg)


def _append_dashed_row(
    out: list[dict], *, kind: str, row: int,
    sx: float, sy: float, ex: float, ey: float,
    tx: float, ty: float, length: float, dash_ft: float, gap_ft: float,
) -> None:
    period = dash_ft + gap_ft
    t = 0.0
    seg_i = 0
    while t + 1e-9 < length:
        t1 = min(t + dash_ft, length)
        out.append({
            "style": "dashed", "kind": kind, "row": row, "seg": seg_i,
            "x1": sx + tx * t, "y1": sy + ty * t,
            "x2": sx + tx * t1, "y2": sy + ty * t1,
        })
        seg_i += 1
        t += period
        if gap_ft == 0 and t1 >= length:
            break


def _row_at(
    out: list[dict], *, style: str, kind: str, row: int, off: float,
    x1: float, y1: float, x2: float, y2: float,
    nx: float, ny: float, tx: float, ty: float, length: float,
    dash_ft: float, gap_ft: float,
) -> None:
    """Straight-corridor row."""
    sx = x1 + nx * off
    sy = y1 + ny * off
    ex = x2 + nx * off
    ey = y2 + ny * off
    if style == "solid":
        _append_solid(out, kind=kind, row=row, sx=sx, sy=sy, ex=ex, ey=ey)
    else:
        _append_dashed_row(
            out, kind=kind, row=row, sx=sx, sy=sy, ex=ex, ey=ey,
            tx=tx, ty=ty, length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )


def _path_vertices(segments: list[PathSegment]) -> list[list[float]]:
    if not segments:
        return []
    out = [[segments[0].sx, segments[0].sy]]
    for seg in segments:
        out.append([seg.ex, seg.ey])
    return out


def _continuous_offset_polyline(
    segments: list[PathSegment],
    off: float,
    side: Literal["right", "left"],
    *,
    step_ft: float = 10.0,
) -> list[list[float]]:
    """Sample a parallel offset along a (possibly filleted) path.

    Normals come from station_to_xy, so they stay continuous through arc
    fillets — no miter gaps/spikes at sharp polyline corners.

    Default step_ft=10 (not 1): a multi-thousand-ft sheet corridor at 1 ft
    densify produced ~5k-vertex PLACE_POLYLINE TSVs that VBA rejected
    (Unknown error) while short dash stubs still placed — live 619-311
    curved reverse-S smoke 2026-08-13.
    """
    if not segments:
        raise ValueError("no path segments")
    length = total_length(segments)
    if length < 1e-9:
        raise ValueError("path length must be > 0")

    stations: list[float] = [0.0]
    # Include every segment seam
    cum = 0.0
    for seg in segments:
        cum += seg.seg_len
        stations.append(cum)
        if seg.seg_len <= step_ft + 1e-9:
            continue
        n = max(1, int(math.ceil(seg.seg_len / step_ft)))
        start = cum - seg.seg_len
        for i in range(1, n):
            stations.append(start + seg.seg_len * (i / n))
    stations = sorted(set(round(s, 6) for s in stations))

    out: list[list[float]] = []
    for st in stations:
        x, y, tx, ty = station_to_xy(segments, st)
        nx, ny = _side_normal(tx, ty, side)
        out.append([x + nx * off, y + ny * off])
    # Dedupe consecutive
    deduped: list[list[float]] = []
    for v in out:
        if deduped and abs(deduped[-1][0] - v[0]) < 1e-9 and abs(deduped[-1][1] - v[1]) < 1e-9:
            continue
        deduped.append(v)
    if len(deduped) < 2:
        return out
    return deduped


def _prepare_path_segments(
    segments: list[PathSegment],
    *,
    fillet_radius_ft: float,
) -> list[PathSegment]:
    if len(segments) <= 1 or fillet_radius_ft <= 0:
        return segments
    return fillet_polyline_segments(segments, fillet_radius_ft)


def _poly_length(verts: list[list[float]]) -> float:
    total = 0.0
    for i in range(len(verts) - 1):
        total += math.hypot(
            verts[i + 1][0] - verts[i][0], verts[i + 1][1] - verts[i][1],
        )
    return total


def _point_along_poly(
    verts: list[list[float]], station: float,
) -> tuple[float, float, float, float]:
    """(x, y, tx, ty) at arc length station along polyline."""
    if len(verts) < 2:
        raise ValueError("polyline needs >= 2 vertices")
    if station <= 0:
        dx = verts[1][0] - verts[0][0]
        dy = verts[1][1] - verts[0][1]
        L = math.hypot(dx, dy) or 1e-9
        return verts[0][0], verts[0][1], dx / L, dy / L
    remaining = station
    for i in range(len(verts) - 1):
        dx = verts[i + 1][0] - verts[i][0]
        dy = verts[i + 1][1] - verts[i][1]
        L = math.hypot(dx, dy)
        if L < 1e-12:
            continue
        if remaining <= L + 1e-9:
            t = max(0.0, min(1.0, remaining / L))
            return (
                verts[i][0] + dx * t, verts[i][1] + dy * t,
                dx / L, dy / L,
            )
        remaining -= L
    dx = verts[-1][0] - verts[-2][0]
    dy = verts[-1][1] - verts[-2][1]
    L = math.hypot(dx, dy) or 1e-9
    return verts[-1][0], verts[-1][1], dx / L, dy / L


def _append_dashed_along_poly(
    out: list[dict], *, kind: str, row: int,
    verts: list[list[float]], dash_ft: float, gap_ft: float,
) -> None:
    length = _poly_length(verts)
    if length < 1e-9:
        return
    period = dash_ft + gap_ft
    t = 0.0
    seg_i = 0
    while t + 1e-9 < length:
        t1 = min(t + dash_ft, length)
        x1, y1, _, _ = _point_along_poly(verts, t)
        x2, y2, _, _ = _point_along_poly(verts, t1)
        out.append({
            "style": "dashed", "kind": kind, "row": row, "seg": seg_i,
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
        })
        seg_i += 1
        t += period
        if gap_ft == 0 and t1 >= length:
            break


def _emit_row(
    out: list[dict], *, style: str, kind: str, row: int, off: float,
    segments: list[PathSegment], side: Literal["right", "left"],
    dash_ft: float, gap_ft: float,
) -> None:
    """Emit one continuous striping row along a (filleted) path."""
    if len(segments) == 1 and not segments[0].is_arc:
        seg = segments[0]
        length, tx, ty, nx, ny = _corridor_frame(
            seg.sx, seg.sy, seg.ex, seg.ey, side,
        )
        _row_at(
            out, style=style, kind=kind, row=row, off=off,
            x1=seg.sx, y1=seg.sy, x2=seg.ex, y2=seg.ey,
            nx=nx, ny=ny, tx=tx, ty=ty, length=length,
            dash_ft=dash_ft, gap_ft=gap_ft,
        )
        return

    verts = _continuous_offset_polyline(segments, off, side)
    if len(verts) < 2:
        return
    if style == "solid":
        _append_solid(
            out, kind=kind, row=row,
            sx=verts[0][0], sy=verts[0][1],
            ex=verts[-1][0], ey=verts[-1][1],
            vertices=verts,
        )
    else:
        _append_dashed_along_poly(
            out, kind=kind, row=row, verts=verts,
            dash_ft=dash_ft, gap_ft=gap_ft,
        )


def _append_shoulders(
    out: list[dict], *, shoulder_width_ft: float,
    first_travel_off: float, last_travel_off: float,
    segments: list[PathSegment], side: Literal["right", "left"],
    next_row: int,
) -> int:
    """Add solid white EOP lines outside both travel outer edges. Returns next row id."""
    if shoulder_width_ft <= 0:
        return next_row
    for off in (first_travel_off - shoulder_width_ft,
                last_travel_off + shoulder_width_ft):
        _emit_row(
            out, style="solid", kind="shoulder", row=next_row, off=off,
            segments=segments, side=side, dash_ft=10.0, gap_ft=30.0,
        )
        next_row += 1
    return next_row


def _finish_shoulders(
    out: list[dict], *, shoulder_width_ft: float,
    first_travel_off: float, last_travel_off: float,
    segments: list[PathSegment], side: Literal["right", "left"],
    next_row: int,
) -> int:
    return _append_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=first_travel_off, last_travel_off=last_travel_off,
        segments=segments, side=side, next_row=next_row,
    )


def lane_highway_lines(
    lanes: int,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """N-lane one-way strip. First travel outer edge = (x1,y1)->(x2,y2) or vertices."""
    if lanes < 1:
        raise ValueError("lanes must be >= 1")
    if lane_width_ft <= 0:
        raise ValueError("lane_width_ft must be > 0")
    if shoulder_width_ft < 0:
        raise ValueError("shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    segments = _prepare_path_segments(
        resolve_edge_path(x1, y1, x2, y2, vertices),
        fillet_radius_ft=float(fillet_radius_ft),
    )
    out: list[dict] = []
    n_rows = lanes + 1
    for row in range(n_rows):
        off = row * lane_width_ft
        kind = "edge" if (row == 0 or row == lanes) else "lane"
        style = "solid" if kind == "edge" else "dashed"
        _emit_row(
            out, style=style, kind=kind, row=row, off=off,
            segments=segments, side=side, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _finish_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=0.0, last_travel_off=lanes * lane_width_ft,
        segments=segments, side=side, next_row=n_rows,
    )
    return out


def two_way_highway_lines(
    lanes: int,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """Even-N undivided two-way (double yellow center). First line = first travel outer."""
    if lanes < 2 or lanes % 2 != 0:
        raise ValueError("two-way lanes must be an even integer >= 2 (2, 4, 6, …)")
    if lane_width_ft <= 0:
        raise ValueError("lane_width_ft must be > 0")
    if yellow_gap_ft <= 0:
        raise ValueError("yellow_gap_ft must be > 0")
    if shoulder_width_ft < 0:
        raise ValueError("shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    segments = _prepare_path_segments(
        resolve_edge_path(x1, y1, x2, y2, vertices),
        fillet_radius_ft=float(fillet_radius_ft),
    )
    per_dir = lanes // 2
    dashed_per_side = per_dir - 1

    rows: list[tuple[str, str]] = [("solid", "edge")]
    for _ in range(dashed_per_side):
        rows.append(("dashed", "lane"))
    rows.append(("solid", "yellow"))
    rows.append(("solid", "yellow"))
    for _ in range(dashed_per_side):
        rows.append(("dashed", "lane"))
    rows.append(("solid", "edge"))

    out: list[dict] = []
    off = 0.0
    offsets: list[float] = []
    for i, (style, kind) in enumerate(rows):
        if i == 0:
            off = 0.0
        elif kind == "yellow" and rows[i - 1][1] == "yellow":
            off += yellow_gap_ft
        else:
            off += lane_width_ft
        offsets.append(off)
        _emit_row(
            out, style=style, kind=kind, row=i, off=off,
            segments=segments, side=side, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _finish_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        segments=segments, side=side, next_row=len(rows),
    )
    return out


def asymmetric_two_way_highway_lines(
    lanes_first: int,
    lanes_second: int,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    median_first_ft: float = 0.0,
    median_second_ft: float = 0.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """Undivided two-way with different lane counts each side of double yellow.

    Built from the first travel outer edge (see ``_first_edge_from_centerline``:
    that edge sits on the **-nx** / left-of-corridor side):
      [lanes_first] + [optional median_first] + double yellow
      + [optional median_second] + [lanes_second].

    Lane-drop sketch: first pack = away/out (2 + median near yellow),
    second pack = toward/approach (3). Across the box that reads
    3 into the intersection and 2 after.
    """
    a = int(lanes_first)
    b = int(lanes_second)
    if a < 1 or b < 1:
        raise ValueError("lanes_first and lanes_second must be >= 1")
    if lane_width_ft <= 0 or yellow_gap_ft <= 0:
        raise ValueError("lane_width_ft and yellow_gap_ft must be > 0")
    if median_first_ft < 0 or median_second_ft < 0 or shoulder_width_ft < 0:
        raise ValueError("median_* and shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    segments = _prepare_path_segments(
        resolve_edge_path(x1, y1, x2, y2, vertices),
        fillet_radius_ft=float(fillet_radius_ft),
    )
    rows: list[tuple[str, str, float]] = []
    rows.append(("solid", "edge", 0.0))
    for _ in range(a - 1):
        rows.append(("dashed", "lane", lane_width_ft))
    if median_first_ft > 0:
        # From last dash: finish last travel lane + median void, then yellow.
        rows.append(("skip", "median", lane_width_ft + median_first_ft))
        rows.append(("solid", "yellow", 0.0))
    else:
        rows.append(("solid", "yellow", lane_width_ft))
    rows.append(("solid", "yellow", yellow_gap_ft))
    if median_second_ft > 0:
        rows.append(("skip", "median", median_second_ft))
    for _ in range(b - 1):
        rows.append(("dashed", "lane", lane_width_ft))
    rows.append(("solid", "edge", lane_width_ft))

    out: list[dict] = []
    off = 0.0
    offsets: list[float] = []
    row_i = 0
    for i, (style, kind, delta) in enumerate(rows):
        if i > 0:
            off += delta
        if style == "skip":
            continue
        offsets.append(off)
        _emit_row(
            out, style=style, kind=kind, row=row_i, off=off,
            segments=segments, side=side, dash_ft=dash_ft, gap_ft=gap_ft,
        )
        row_i += 1
    if offsets:
        _finish_shoulders(
            out, shoulder_width_ft=shoulder_width_ft,
            first_travel_off=offsets[0], last_travel_off=offsets[-1],
            segments=segments, side=side, next_row=row_i,
        )
    return out


def asymmetric_two_way_width_ft(
    lanes_first: int,
    lanes_second: int,
    *,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    median_first_ft: float = 0.0,
    median_second_ft: float = 0.0,
) -> float:
    return (
        float(lanes_first) * float(lane_width_ft)
        + float(median_first_ft)
        + float(yellow_gap_ft)
        + float(median_second_ft)
        + float(lanes_second) * float(lane_width_ft)
    )


def divided_highway_lines(
    lanes_per_direction: int,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    median_width_ft: float,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """Dual carriageway with physical median gap (619-302-style).

    Layout from first travel outer edge:
      white solid outer
      (N-1) dashed white
      yellow solid (median edge)
      [median_width_ft empty]
      yellow solid (median edge)
      (N-1) dashed white
      white solid outer
    Optional solid white shoulders outside both outers.
    """
    n = int(lanes_per_direction)
    if n < 1:
        raise ValueError("lanes_per_direction must be >= 1")
    if median_width_ft <= 0:
        raise ValueError("median_width_ft must be > 0 (physical median gap)")
    if lane_width_ft <= 0:
        raise ValueError("lane_width_ft must be > 0")
    if shoulder_width_ft < 0:
        raise ValueError("shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    segments = _prepare_path_segments(
        resolve_edge_path(x1, y1, x2, y2, vertices),
        fillet_radius_ft=float(fillet_radius_ft),
    )
    dashed = n - 1

    # Direction A then B (style, kind, step_after_prev): first row step ignored.
    # Steps: lane_width between travel lines; median_width between yellows.
    rows: list[tuple[str, str]] = [("solid", "edge")]
    for _ in range(dashed):
        rows.append(("dashed", "lane"))
    rows.append(("solid", "yellow"))
    rows.append(("solid", "yellow"))  # after median gap
    for _ in range(dashed):
        rows.append(("dashed", "lane"))
    rows.append(("solid", "edge"))

    out: list[dict] = []
    off = 0.0
    offsets: list[float] = []
    for i, (style, kind) in enumerate(rows):
        if i == 0:
            off = 0.0
        elif kind == "yellow" and rows[i - 1][1] == "yellow":
            off += median_width_ft
        else:
            off += lane_width_ft
        offsets.append(off)
        _emit_row(
            out, style=style, kind=kind, row=i, off=off,
            segments=segments, side=side, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _finish_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        segments=segments, side=side, next_row=len(rows),
    )
    return out


def twlt_highway_lines(
    lanes_per_direction: int,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    twlt_width_ft: float = 12.0,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """Undivided multilane with center two-way left-turn lane (312-style).

    lanes_per_direction = travel lanes each way (TWLT not counted).
    Layout:
      white solid outer
      (L-1) dashed white
      dashed yellow (TWLT bound)
      [twlt_width_ft]
      dashed yellow (TWLT bound)
      (L-1) dashed white
      white solid outer
    """
    l = int(lanes_per_direction)
    if l < 1:
        raise ValueError("lanes_per_direction must be >= 1")
    if twlt_width_ft <= 0:
        raise ValueError("twlt_width_ft must be > 0")
    if lane_width_ft <= 0:
        raise ValueError("lane_width_ft must be > 0")
    if shoulder_width_ft < 0:
        raise ValueError("shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    segments = _prepare_path_segments(
        resolve_edge_path(x1, y1, x2, y2, vertices),
        fillet_radius_ft=float(fillet_radius_ft),
    )
    dashed = l - 1

    rows: list[tuple[str, str]] = [("solid", "edge")]
    for _ in range(dashed):
        rows.append(("dashed", "lane"))
    rows.append(("dashed", "yellow"))
    rows.append(("dashed", "yellow"))
    for _ in range(dashed):
        rows.append(("dashed", "lane"))
    rows.append(("solid", "edge"))

    out: list[dict] = []
    off = 0.0
    offsets: list[float] = []
    for i, (style, kind) in enumerate(rows):
        if i == 0:
            off = 0.0
        elif kind == "yellow" and rows[i - 1][1] == "yellow":
            off += twlt_width_ft
        else:
            off += lane_width_ft
        offsets.append(off)
        _emit_row(
            out, style=style, kind=kind, row=i, off=off,
            segments=segments, side=side, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _finish_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        segments=segments, side=side, next_row=len(rows),
    )
    return out


def travel_width_ft(
    road_type: str,
    *,
    lanes: int | None = None,
    lanes_per_direction: int | None = None,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    median_width_ft: float = 0.0,
    twlt_width_ft: float = 12.0,
) -> float:
    """Distance between the two outer travel edges (not including shoulders)."""
    rt = (road_type or "").strip().lower()
    lw = float(lane_width_ft)
    if rt in ("one_way", "one-way", "highway"):
        if lanes is None or lanes < 1:
            raise ValueError("one_way requires lanes >= 1")
        return float(lanes) * lw
    if rt in ("two_way", "two-way", "undivided"):
        if lanes is None or lanes < 2 or lanes % 2 != 0:
            raise ValueError("two_way requires even lanes >= 2")
        return float(lanes) * lw + float(yellow_gap_ft)
    if rt in ("divided", "freeway", "median"):
        if lanes_per_direction is None or lanes_per_direction < 1:
            raise ValueError("divided requires lanes_per_direction >= 1")
        if median_width_ft <= 0:
            raise ValueError("divided requires median_width_ft > 0")
        return 2.0 * float(lanes_per_direction) * lw + float(median_width_ft)
    if rt in ("twlt",):
        if lanes_per_direction is None or lanes_per_direction < 1:
            raise ValueError("twlt requires lanes_per_direction >= 1")
        return 2.0 * float(lanes_per_direction) * lw + float(twlt_width_ft)
    raise ValueError(
        "road_type must be one_way|two_way|divided|twlt "
        f"(got {road_type!r})"
    )


def build_strip_lines(
    road_type: str,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    *,
    lanes: int | None = None,
    lanes_per_direction: int | None = None,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    median_width_ft: float = 0.0,
    twlt_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    vertices: Sequence[Sequence[float]] | None = None,
    fillet_radius_ft: float = 150.0,
) -> list[dict]:
    """Dispatch to the matching strip builder."""
    rt = (road_type or "").strip().lower()
    kw = dict(
        lane_width_ft=lane_width_ft,
        shoulder_width_ft=shoulder_width_ft,
        dash_ft=dash_ft,
        gap_ft=gap_ft,
        side=side,
        vertices=vertices,
        fillet_radius_ft=fillet_radius_ft,
    )
    if rt in ("one_way", "one-way", "highway"):
        return lane_highway_lines(int(lanes or 0), x1, y1, x2, y2, **kw)
    if rt in ("two_way", "two-way", "undivided"):
        return two_way_highway_lines(
            int(lanes or 0), x1, y1, x2, y2, yellow_gap_ft=yellow_gap_ft, **kw,
        )
    if rt in ("divided", "freeway", "median"):
        return divided_highway_lines(
            int(lanes_per_direction or 0), x1, y1, x2, y2,
            median_width_ft=median_width_ft, **kw,
        )
    if rt in ("twlt",):
        return twlt_highway_lines(
            int(lanes_per_direction or 0), x1, y1, x2, y2,
            twlt_width_ft=twlt_width_ft, **kw,
        )
    raise ValueError(
        "road_type must be one_way|two_way|divided|twlt "
        f"(got {road_type!r})"
    )
