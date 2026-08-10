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
"""
from __future__ import annotations

import math
from typing import Literal


def _corridor_frame(
    x1: float, y1: float, x2: float, y2: float, side: Literal["right", "left"],
) -> tuple[float, float, float, float, float]:
    dx = x2 - x1
    dy = y2 - y1
    length = math.hypot(dx, dy)
    if length < 1e-6:
        raise ValueError("corridor length must be > 0 (x1,y1) and (x2,y2) coincide")
    tx, ty = dx / length, dy / length
    if side == "left":
        nx, ny = -ty, tx
    elif side == "right":
        nx, ny = ty, -tx
    else:
        raise ValueError("side must be 'right' or 'left'")
    return length, tx, ty, nx, ny


def _append_solid(out: list[dict], *, kind: str, row: int,
                  sx: float, sy: float, ex: float, ey: float) -> None:
    out.append({
        "style": "solid", "kind": kind, "row": row,
        "x1": sx, "y1": sy, "x2": ex, "y2": ey,
    })


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


def _append_shoulders(
    out: list[dict], *, shoulder_width_ft: float,
    first_travel_off: float, last_travel_off: float,
    x1: float, y1: float, x2: float, y2: float,
    nx: float, ny: float, next_row: int,
) -> int:
    """Add solid white EOP lines outside both travel outer edges. Returns next row id."""
    if shoulder_width_ft <= 0:
        return next_row
    for off in (first_travel_off - shoulder_width_ft,
                last_travel_off + shoulder_width_ft):
        _append_solid(
            out, kind="shoulder", row=next_row,
            sx=x1 + nx * off, sy=y1 + ny * off,
            ex=x2 + nx * off, ey=y2 + ny * off,
        )
        next_row += 1
    return next_row


def lane_highway_lines(
    lanes: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
) -> list[dict]:
    """N-lane one-way strip. (x1,y1)->(x2,y2) is the first travel outer edge."""
    if lanes < 1:
        raise ValueError("lanes must be >= 1")
    if lane_width_ft <= 0:
        raise ValueError("lane_width_ft must be > 0")
    if shoulder_width_ft < 0:
        raise ValueError("shoulder_width_ft must be >= 0")
    if dash_ft <= 0 or gap_ft < 0:
        raise ValueError("dash_ft must be > 0 and gap_ft >= 0")

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
    out: list[dict] = []
    n_rows = lanes + 1
    for row in range(n_rows):
        off = row * lane_width_ft
        kind = "edge" if (row == 0 or row == lanes) else "lane"
        style = "solid" if kind == "edge" else "dashed"
        _row_at(
            out, style=style, kind=kind, row=row, off=off,
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, tx=tx, ty=ty,
            length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _append_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=0.0, last_travel_off=lanes * lane_width_ft,
        x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, next_row=n_rows,
    )
    return out


def two_way_highway_lines(
    lanes: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
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

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
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
        _row_at(
            out, style=style, kind=kind, row=i, off=off,
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, tx=tx, ty=ty,
            length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _append_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, next_row=len(rows),
    )
    return out


def asymmetric_two_way_highway_lines(
    lanes_first: int,
    lanes_second: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    median_first_ft: float = 0.0,
    median_second_ft: float = 0.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
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

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
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
        _row_at(
            out, style=style, kind=kind, row=row_i, off=off,
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, tx=tx, ty=ty,
            length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )
        row_i += 1
    if offsets:
        _append_shoulders(
            out, shoulder_width_ft=shoulder_width_ft,
            first_travel_off=offsets[0], last_travel_off=offsets[-1],
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, next_row=row_i,
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
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    median_width_ft: float,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
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

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
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
        _row_at(
            out, style=style, kind=kind, row=i, off=off,
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, tx=tx, ty=ty,
            length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _append_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, next_row=len(rows),
    )
    return out


def twlt_highway_lines(
    lanes_per_direction: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    twlt_width_ft: float = 12.0,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
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

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
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
        _row_at(
            out, style=style, kind=kind, row=i, off=off,
            x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, tx=tx, ty=ty,
            length=length, dash_ft=dash_ft, gap_ft=gap_ft,
        )
    _append_shoulders(
        out, shoulder_width_ft=shoulder_width_ft,
        first_travel_off=offsets[0], last_travel_off=offsets[-1],
        x1=x1, y1=y1, x2=x2, y2=y2, nx=nx, ny=ny, next_row=len(rows),
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
    x1: float,
    y1: float,
    x2: float,
    y2: float,
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
) -> list[dict]:
    """Dispatch to the matching strip builder."""
    rt = (road_type or "").strip().lower()
    kw = dict(
        lane_width_ft=lane_width_ft,
        shoulder_width_ft=shoulder_width_ft,
        dash_ft=dash_ft,
        gap_ft=gap_ft,
        side=side,
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
