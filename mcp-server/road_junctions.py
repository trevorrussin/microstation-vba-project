"""Intersection and ramp-gore plan geometry composing lane_highway strips.

Orthogonal intersection (MUTCD 3B.11 / NYSDOT 685-style sketch):
  Primary and secondary arms STOP at the intersection box — solid edge /
  center / lane lines do NOT continue through intersecting approaches.
  Each approach gets a transverse crosswalk (white) and stop bar (white,
  4 ft beyond the far crosswalk line). Dotted centerline extension through
  the box when has_turning_lanes, TWLT, or dedicated turn pockets
  (lanes_in > lanes_out). Turn arrows use ny_plan_striping.cel with
  ACTIVE ANGLE 0 = +Y; SAL/SAR+SLONLY only for dedicated pockets.

  junction='plus' both secondary stubs; junction='tee' one stub on tee_side.

Ramp gore (Family 5 entrance/exit sketch, general CAD):
  Mainline one-way strip along (x1,y1)->(x2,y2).
  Ramp one-way strip diverges at ramp_angle_deg toward side at
  gore_station_ft along the mainline ramp-side edge; nose is the shared
  point. Optional solid white gore marks from the nose back along both
  corridors.
"""
from __future__ import annotations

import math
from typing import Literal

from lane_highway import (
    _corridor_frame,
    build_strip_lines,
    lane_highway_lines,
    travel_width_ft,
)

# MUTCD-ish defaults for the intersection sketch (plan view, not 685 CAD cells).
_CROSSWALK_WIDTH_FT = 8.0          # gap between the two transverse lines (>= 6)
_STOP_BAR_AFTER_CROSSWALK_FT = 4.0  # stop line in advance of nearest crosswalk
_DOTTED_EXT_DASH_FT = 2.0          # MUTCD dotted extension ~2 ft segments
_DOTTED_EXT_GAP_FT = 4.0           # common 2–6 ft gap; use 4
_ARROW_SETBACK_FT = 20.0           # arrow center outbound of stop bar
DEFAULT_STRIPING_CELL_LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"
# ny_plan_striping.cel lane-use arrows (descriptions confirmed live).
_CELL_ARROW_LEFT = "SAL"
_CELL_ARROW_RIGHT = "SAR"
_CELL_ARROW_STRAIGHT = "SAS"
_CELL_ARROW_LEFT_STRAIGHT = "SALS"
_CELL_ARROW_RIGHT_STRAIGHT = "SARS"
_CELL_ONLY = "SLONLY"
# No left+straight+right cell in ny_plan_striping.cel — emit SALS+SARS pair.
_CELL_LSR_PAIR = "__LSR_PAIR__"


def _rotate(tx: float, ty: float, deg: float) -> tuple[float, float]:
    """Rotate vector CCW by deg."""
    r = math.radians(deg)
    c, s = math.cos(r), math.sin(r)
    return tx * c - ty * s, tx * s + ty * c


def _tag(segs: list[dict], arm: str) -> list[dict]:
    out = []
    for s in segs:
        d = dict(s)
        d["arm"] = arm
        out.append(d)
    return out


def _first_edge_from_centerline(
    c0x: float, c0y: float, c1x: float, c1y: float,
    travel_w: float, side: Literal["right", "left"],
) -> tuple[float, float, float, float]:
    """Return (x1,y1,x2,y2) for strip first travel edge given centerline ends."""
    _, _, _, nx, ny = _corridor_frame(c0x, c0y, c1x, c1y, side)
    half = travel_w / 2.0
    return (
        c0x - nx * half, c0y - ny * half,
        c1x - nx * half, c1y - ny * half,
    )


def _append_transverse(
    out: list[dict], *, kind: str, arm: str, row: int,
    cx: float, cy: float, trx: float, try_: float, half_w: float,
) -> None:
    """Solid white line across the approach (stop bar / crosswalk bar)."""
    out.append({
        "style": "solid", "kind": kind, "row": row, "arm": arm,
        "x1": cx - trx * half_w, "y1": cy - try_ * half_w,
        "x2": cx + trx * half_w, "y2": cy + try_ * half_w,
    })


def _append_dotted_center(
    out: list[dict], *, arm: str,
    x1: float, y1: float, x2: float, y2: float,
) -> None:
    """Yellow dotted centerline extension through the intersection box."""
    dx, dy = x2 - x1, y2 - y1
    length = math.hypot(dx, dy)
    if length < 1e-6:
        return
    tx, ty = dx / length, dy / length
    period = _DOTTED_EXT_DASH_FT + _DOTTED_EXT_GAP_FT
    t = 0.0
    seg_i = 0
    while t + 1e-9 < length:
        t1 = min(t + _DOTTED_EXT_DASH_FT, length)
        out.append({
            "style": "dashed", "kind": "yellow", "row": 800, "arm": arm,
            "seg": seg_i,
            "x1": x1 + tx * t, "y1": y1 + ty * t,
            "x2": x1 + tx * t1, "y2": y1 + ty * t1,
        })
        seg_i += 1
        t += period


def _approach_stop_and_crosswalk(
    out: list[dict], *,
    jx: float, jy: float,
    out_dx: float, out_dy: float,
    box_edge_ft: float,
    travel_w: float,
    arm: str,
    crosswalk: bool,
    stop_bar: bool,
) -> float:
    """Place crosswalk then stop bar. Returns stop-bar station (outbound)."""
    trx, try_ = -out_dy, out_dx
    half = travel_w / 2.0
    row0 = 700
    stop_s = box_edge_ft

    if crosswalk:
        for i, s in enumerate((box_edge_ft, box_edge_ft + _CROSSWALK_WIDTH_FT)):
            cx = jx + out_dx * s
            cy = jy + out_dy * s
            _append_transverse(
                out, kind="crosswalk", arm=arm, row=row0 + i,
                cx=cx, cy=cy, trx=trx, try_=try_, half_w=half,
            )
        stop_s = box_edge_ft + _CROSSWALK_WIDTH_FT

    if stop_bar:
        if crosswalk:
            stop_s += _STOP_BAR_AFTER_CROSSWALK_FT
        cx = jx + out_dx * stop_s
        cy = jy + out_dy * stop_s
        _append_transverse(
            out, kind="stop_bar", arm=arm, row=row0 + 10,
            cx=cx, cy=cy, trx=trx, try_=try_, half_w=half,
        )
    return float(stop_s)


def _station(jx: float, jy: float, out_dx: float, out_dy: float,
             x: float, y: float) -> float:
    return (x - jx) * out_dx + (y - jy) * out_dy


def _clip_center_lane_before_stop(
    segs: list[dict], *,
    jx: float, jy: float, out_dx: float, out_dy: float, stop_s: float,
) -> list[dict]:
    """Keep edge/shoulder to the box; stop yellow/lane at the stop bar.

    Longitudinal center/lane markings must not enter the stop-bar/crosswalk
    zone. Edge lines still meet the intersection box so arms connect.
    """
    kept: list[dict] = []
    for seg in segs:
        kind = seg.get("kind") or ""
        if kind in ("edge", "shoulder", "crosswalk", "stop_bar", "gore"):
            kept.append(seg)
            continue
        if kind not in ("yellow", "lane"):
            kept.append(seg)
            continue

        s1 = _station(jx, jy, out_dx, out_dy, seg["x1"], seg["y1"])
        s2 = _station(jx, jy, out_dx, out_dy, seg["x2"], seg["y2"])
        lo, hi = (s1, s2) if s1 <= s2 else (s2, s1)
        if hi <= stop_s + 1e-6:
            continue  # entirely inside mark zone / box side of stop
        if lo >= stop_s - 1e-6:
            kept.append(seg)
            continue
        # Trim the endpoint on the box side of stop_s up to the stop bar.
        dx = seg["x2"] - seg["x1"]
        dy = seg["y2"] - seg["y1"]
        if abs(s2 - s1) < 1e-9:
            continue
        t_stop = (stop_s - s1) / (s2 - s1)
        t_stop = max(0.0, min(1.0, t_stop))
        xs = seg["x1"] + dx * t_stop
        ys = seg["y1"] + dy * t_stop
        d = dict(seg)
        if s1 < s2:
            d["x1"], d["y1"] = xs, ys
        else:
            d["x2"], d["y2"] = xs, ys
        # Drop zero-length
        if math.hypot(d["x2"] - d["x1"], d["y2"] - d["y1"]) < 1e-6:
            continue
        kept.append(d)
    return kept


def _lanes_toward(
    road_type: str, *,
    lanes: int | None, lanes_per_direction: int | None,
) -> int:
    rt = (road_type or "").strip().lower()
    if rt in ("one_way", "one-way", "highway"):
        return max(int(lanes or 1), 1)
    if rt in ("two_way", "two-way", "undivided"):
        return max(int(lanes or 2) // 2, 1)
    if rt in ("divided", "freeway", "median", "twlt"):
        return max(int(lanes_per_direction or 1), 1)
    return 1


def _ms_striping_arrow_angle_deg(travel_x: float, travel_y: float) -> float:
    """ACTIVE ANGLE for ny_plan_striping SAS/SAL/… cells.

    Live probe: angle 0 aligns cell +Y. Tip of SAS matches that axis.
    Convert travel (unit vector) with atan2(-tx, ty). Intersection QA with
    +180 left tips facing away — do not add 180.
    """
    return math.degrees(math.atan2(-travel_x, travel_y))


def _dedicated_turn_count(lanes_in: int, lanes_out: int | None) -> int:
    """Mandatory turn pockets = approach lanes minus through/receiving lanes."""
    lin = max(int(lanes_in), 1)
    if lanes_out is None:
        return 0
    lout = max(int(lanes_out), 1)
    return max(0, lin - lout)


def _center_clearance_ft(
    road_type: str, *,
    yellow_gap_ft: float = 2.0,
    median_width_ft: float = 0.0,
    twlt_width_ft: float = 12.0,
) -> float | None:
    """Half-width of non-travel center (yellow / median / TWLT), or None for one-way."""
    rt = (road_type or "").strip().lower()
    if rt in ("one_way", "one-way", "highway"):
        return None
    if rt in ("two_way", "two-way", "undivided"):
        return float(yellow_gap_ft) / 2.0
    if rt in ("divided", "freeway", "median"):
        return float(median_width_ft) / 2.0
    if rt == "twlt":
        return float(twlt_width_ft) / 2.0
    return float(yellow_gap_ft) / 2.0


def _allowed_turns_for_arm(
    *,
    junction: str,
    arm: str,
    tee_side: str,
) -> tuple[bool, bool, bool]:
    """(can_left, can_straight, can_right) relative to approach travel."""
    junc = (junction or "plus").strip().lower()
    if junc == "plus":
        return True, True, True
    # Tee: primary keeps straight; one cross turn toward the stub; stub has L+R only.
    ts = (tee_side or "right").strip().lower()
    if arm.startswith("secondary"):
        return True, False, True
    # primary_* — stub on geometric right of primary bearing is tee_side 'right'
    # primary_neg travel = +primary bearing; right turn reaches tee_side=right stub.
    if arm == "primary_neg":
        if ts == "right":
            return False, True, True
        return True, True, False
    if arm == "primary_pos":
        if ts == "right":
            return True, True, False
        return False, True, True
    return True, True, True


def _shared_through_cells(
    through_n: int,
    can_left: bool,
    can_straight: bool,
    can_right: bool,
) -> list[str]:
    """Lane-use cells for non-dedicated through lanes (no SLONLY)."""
    n = max(int(through_n), 0)
    if n <= 0:
        return []
    L, S, R = bool(can_left), bool(can_straight), bool(can_right)

    def _one() -> str:
        if L and S and R:
            return _CELL_LSR_PAIR
        if L and S:
            return _CELL_ARROW_LEFT_STRAIGHT
        if R and S:
            return _CELL_ARROW_RIGHT_STRAIGHT
        if L and R:
            return _CELL_LSR_PAIR
        if L:
            return _CELL_ARROW_LEFT
        if R:
            return _CELL_ARROW_RIGHT
        return _CELL_ARROW_STRAIGHT

    if n == 1:
        return [_one()]
    if n == 2:
        left = (
            _CELL_ARROW_LEFT_STRAIGHT if L and S else
            (_CELL_ARROW_LEFT if L and not S else _CELL_ARROW_STRAIGHT)
        )
        right = (
            _CELL_ARROW_RIGHT_STRAIGHT if R and S else
            (_CELL_ARROW_RIGHT if R and not S else _CELL_ARROW_STRAIGHT)
        )
        if not L and not R:
            return [_CELL_ARROW_STRAIGHT, _CELL_ARROW_STRAIGHT]
        return [left, right]
    # 3+ through: left shared L+S, centers straight only, right shared R+S
    out: list[str] = []
    for i in range(n):
        if i == 0:
            out.append(
                _CELL_ARROW_LEFT_STRAIGHT if L and S else
                (_CELL_ARROW_LEFT if L else _CELL_ARROW_STRAIGHT)
            )
        elif i == n - 1:
            out.append(
                _CELL_ARROW_RIGHT_STRAIGHT if R and S else
                (_CELL_ARROW_RIGHT if R else _CELL_ARROW_STRAIGHT)
            )
        else:
            out.append(_CELL_ARROW_STRAIGHT)
    return out


def _arrow_cell_placements(
    lanes_in: int,
    dedicated: int,
    can_left: bool,
    can_straight: bool,
    can_right: bool,
) -> list[tuple[str, int, bool]]:
    """(cellName, lane_index, place_only) for each approach lane.

    Dedicated pockets (lanes_in > lanes_out) get SAL/SAR + SLONLY on the
    drop lanes. Remaining through lanes get shared options (SALS / SAS /
    SARS / LSR pair) from what turns are legal — never SLONLY.
    """
    n = max(int(lanes_in), 1)
    d = max(0, min(int(dedicated), n))
    left_d = (d + 1) // 2
    right_d = d - left_d
    through_n = n - d
    # Once a left/right ONLY pocket exists, remaining lanes lose that turn.
    eff_l = can_left and left_d == 0
    eff_r = can_right and right_d == 0
    through_cells = _shared_through_cells(
        through_n, eff_l, can_straight, eff_r,
    )

    out: list[tuple[str, int, bool]] = []
    through_i = 0
    for i in range(n):
        if i < left_d:
            out.append((_CELL_ARROW_LEFT, i, True))
        elif right_d and i >= n - right_d:
            out.append((_CELL_ARROW_RIGHT, i, True))
        else:
            cell = through_cells[through_i]
            through_i += 1
            out.append((cell, i, False))
    return out


def _append_turn_arrow_metas(
    out: list[dict], *,
    jx: float, jy: float,
    out_dx: float, out_dy: float,
    stop_s: float,
    travel_w: float,
    lane_width_ft: float,
    lanes_toward: int,
    lanes_out: int | None,
    arm: str,
    road_type: str,
    yellow_gap_ft: float = 2.0,
    median_width_ft: float = 0.0,
    twlt_width_ft: float = 12.0,
    can_left: bool = True,
    can_straight: bool = True,
    can_right: bool = True,
) -> None:
    """Meta markers for ny_plan_striping.cel arrows upstream of the stop bar.

    Travel is toward the junction (-outbound). Approach lanes sit on the
    right-hand half of a two-way/divided strip (US). Striping cells use
    ACTIVE ANGLE 0 = +Y.
    """
    travel_x, travel_y = -out_dx, -out_dy
    angle = _ms_striping_arrow_angle_deg(travel_x, travel_y)
    left_x, left_y = -travel_y, travel_x
    right_x, right_y = travel_y, -travel_x
    s_arrow = stop_s + _ARROW_SETBACK_FT
    cx0 = jx + out_dx * s_arrow
    cy0 = jy + out_dy * s_arrow
    n = max(int(lanes_toward), 1)
    dedicated = _dedicated_turn_count(n, lanes_out)
    clearance = _center_clearance_ft(
        road_type,
        yellow_gap_ft=yellow_gap_ft,
        median_width_ft=median_width_ft,
        twlt_width_ft=twlt_width_ft,
    )

    def _lane_center(i: int) -> tuple[float, float]:
        # i=0 = leftmost approach lane (nearest centerline for two-way).
        if clearance is None:
            # one-way: full width, i=0 leftmost when facing travel
            off_left = travel_w / 2.0 - (i + 0.5) * lane_width_ft
            return cx0 + left_x * off_left, cy0 + left_y * off_left
        dist = clearance + (i + 0.5) * lane_width_ft
        return cx0 + right_x * dist, cy0 + right_y * dist

    def _emit(cell: str, i: int, place_only: bool, ax: float, ay: float) -> None:
        out.append({
            "style": "meta", "kind": "turn_arrow", "arm": arm, "row": 850 + i,
            "cellName": cell,
            "libraryPath": DEFAULT_STRIPING_CELL_LIB,
            "x": ax, "y": ay, "angleDeg": angle,
            "x1": ax, "y1": ay, "x2": ax, "y2": ay,
            "lanesIn": n,
            "lanesOut": n if lanes_out is None else int(lanes_out),
            "dedicated": dedicated,
        })
        if place_only:
            ox = ax + travel_x * (-8.0)
            oy = ay + travel_y * (-8.0)
            out.append({
                "style": "meta", "kind": "turn_arrow", "arm": arm, "row": 860 + i,
                "cellName": _CELL_ONLY,
                "libraryPath": DEFAULT_STRIPING_CELL_LIB,
                "x": ox, "y": oy, "angleDeg": angle,
                "x1": ox, "y1": oy, "x2": ox, "y2": oy,
            })

    for cell, i, place_only in _arrow_cell_placements(
        n, dedicated, can_left, can_straight, can_right,
    ):
        ax, ay = _lane_center(i)
        if cell == _CELL_LSR_PAIR:
            # No triple-head cell in the NY striping lib — stack L+S and R+S.
            _emit(
                _CELL_ARROW_LEFT_STRAIGHT, i, False,
                ax + travel_x * (-4.0), ay + travel_y * (-4.0),
            )
            _emit(
                _CELL_ARROW_RIGHT_STRAIGHT, i, False,
                ax + travel_x * (4.0), ay + travel_y * (4.0),
            )
        else:
            _emit(cell, i, place_only, ax, ay)


def _resolve_dotted_center(
    has_turning_lanes: bool | None,
    primary_road_type: str,
    secondary_road_type: str,
    *,
    primary_dedicated: int,
    secondary_dedicated: int,
) -> bool:
    """Dotted yellow through the box: explicit flag, TWLT, or real dedicated turns."""
    if has_turning_lanes is True:
        return True
    if has_turning_lanes is False:
        return False
    rts = {
        (primary_road_type or "").strip().lower(),
        (secondary_road_type or "").strip().lower(),
    }
    if "twlt" in rts:
        return True
    return primary_dedicated > 0 or secondary_dedicated > 0


def orthogonal_intersection_lines(
    junction_x: float,
    junction_y: float,
    *,
    primary_road_type: str,
    secondary_road_type: str,
    primary_length_ft: float,
    secondary_stub_ft: float,
    primary_bearing_deg: float = 0.0,
    junction: Literal["plus", "tee"] = "plus",
    tee_side: Literal["left", "right"] = "right",
    primary_lanes: int | None = None,
    secondary_lanes: int | None = None,
    primary_lanes_per_direction: int | None = None,
    secondary_lanes_per_direction: int | None = None,
    lane_width_ft: float = 12.0,
    yellow_gap_ft: float = 2.0,
    primary_median_width_ft: float = 0.0,
    secondary_median_width_ft: float = 0.0,
    primary_twlt_width_ft: float = 12.0,
    secondary_twlt_width_ft: float = 12.0,
    primary_shoulder_width_ft: float = 0.0,
    secondary_shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
    side: Literal["right", "left"] = "right",
    crosswalks: bool = True,
    stop_bars: bool = True,
    has_turning_lanes: bool | None = None,
    turn_arrows: bool = True,
    primary_lanes_out: int | None = None,
    secondary_lanes_out: int | None = None,
) -> list[dict]:
    """Return +/T intersection segments with MUTCD box rules.

    Edge lines meet the intersection box (arms connect). Yellow center /
    dashed lane lines stop at the stop bar (not through the stop box).
    Crosswalks + stop bars on every approach by default. Turn-arrow metas
    from ny_plan_striping.cel when turn_arrows=True. Dedicated SAL/SAR +
    SLONLY only when lanes_in > lanes_out (primary_lanes_out /
    secondary_lanes_out); equal in/out → through SAS only. Dotted yellow
    through the box when has_turning_lanes, TWLT, or dedicated > 0.
    """
    if primary_length_ft <= 0 or secondary_stub_ft <= 0:
        raise ValueError("primary_length_ft and secondary_stub_ft must be > 0")
    junc = (junction or "plus").strip().lower()
    if junc not in ("plus", "tee"):
        raise ValueError("junction must be 'plus' or 'tee'")

    mark_depth = 0.0
    if crosswalks:
        mark_depth += _CROSSWALK_WIDTH_FT
    if stop_bars:
        mark_depth += _STOP_BAR_AFTER_CROSSWALK_FT if crosswalks else 0.0
    arrow_depth = _ARROW_SETBACK_FT if turn_arrows else 0.0

    ptx, pty = _rotate(1.0, 0.0, primary_bearing_deg)
    stx, sty = _rotate(ptx, pty, 90.0)

    strip_kw_p = dict(
        lanes=primary_lanes,
        lanes_per_direction=primary_lanes_per_direction,
        lane_width_ft=lane_width_ft,
        yellow_gap_ft=yellow_gap_ft,
        median_width_ft=primary_median_width_ft,
        twlt_width_ft=primary_twlt_width_ft,
        shoulder_width_ft=primary_shoulder_width_ft,
        dash_ft=dash_ft, gap_ft=gap_ft, side=side,
    )
    strip_kw_s = dict(
        lanes=secondary_lanes,
        lanes_per_direction=secondary_lanes_per_direction,
        lane_width_ft=lane_width_ft,
        yellow_gap_ft=yellow_gap_ft,
        median_width_ft=secondary_median_width_ft,
        twlt_width_ft=secondary_twlt_width_ft,
        shoulder_width_ft=secondary_shoulder_width_ft,
        dash_ft=dash_ft, gap_ft=gap_ft, side=side,
    )

    pw = travel_width_ft(
        primary_road_type,
        lanes=primary_lanes,
        lanes_per_direction=primary_lanes_per_direction,
        lane_width_ft=lane_width_ft,
        yellow_gap_ft=yellow_gap_ft,
        median_width_ft=primary_median_width_ft,
        twlt_width_ft=primary_twlt_width_ft,
    )
    sw = travel_width_ft(
        secondary_road_type,
        lanes=secondary_lanes,
        lanes_per_direction=secondary_lanes_per_direction,
        lane_width_ft=lane_width_ft,
        yellow_gap_ft=yellow_gap_ft,
        median_width_ft=secondary_median_width_ft,
        twlt_width_ft=secondary_twlt_width_ft,
    )
    primary_half = pw / 2.0 + float(primary_shoulder_width_ft)
    secondary_half = sw / 2.0 + float(secondary_shoulder_width_ft)
    half_p = primary_length_ft / 2.0

    need = mark_depth + arrow_depth + 1.0
    if half_p - secondary_half < need:
        raise ValueError(
            f"primary_length_ft too short for intersection marks "
            f"(need each primary arm >= box half {secondary_half:.1f} + "
            f"{need:.1f})"
        )
    if secondary_stub_ft < need:
        raise ValueError(
            f"secondary_stub_ft ({secondary_stub_ft}) too short for "
            f"crosswalk/stop/arrows (need >= {need:.1f})"
        )

    out: list[dict] = []
    p_lanes_toward = _lanes_toward(
        primary_road_type, lanes=primary_lanes,
        lanes_per_direction=primary_lanes_per_direction,
    )
    s_lanes_toward = _lanes_toward(
        secondary_road_type, lanes=secondary_lanes,
        lanes_per_direction=secondary_lanes_per_direction,
    )
    # Default lanes_out = lanes_in (continuous through) → no dedicated turns.
    p_lanes_out = (
        p_lanes_toward if primary_lanes_out is None else int(primary_lanes_out)
    )
    s_lanes_out = (
        s_lanes_toward if secondary_lanes_out is None else int(secondary_lanes_out)
    )
    p_dedicated = _dedicated_turn_count(p_lanes_toward, p_lanes_out)
    s_dedicated = _dedicated_turn_count(s_lanes_toward, s_lanes_out)
    turning = _resolve_dotted_center(
        has_turning_lanes, primary_road_type, secondary_road_type,
        primary_dedicated=p_dedicated, secondary_dedicated=s_dedicated,
    )

    def _primary_arm(sign: float, arm: str) -> None:
        outer_s = sign * half_p
        inner_s = sign * secondary_half
        c0x = junction_x + ptx * outer_s
        c0y = junction_y + pty * outer_s
        c1x = junction_x + ptx * inner_s
        c1y = junction_y + pty * inner_s
        x1, y1, x2, y2 = _first_edge_from_centerline(c0x, c0y, c1x, c1y, pw, side)
        raw = _tag(
            build_strip_lines(primary_road_type, x1, y1, x2, y2, **strip_kw_p),
            arm,
        )
        out_dx, out_dy = sign * ptx, sign * pty
        stop_s = _approach_stop_and_crosswalk(
            out, jx=junction_x, jy=junction_y,
            out_dx=out_dx, out_dy=out_dy,
            box_edge_ft=secondary_half, travel_w=pw, arm=arm,
            crosswalk=crosswalks, stop_bar=stop_bars,
        )
        out.extend(_clip_center_lane_before_stop(
            raw, jx=junction_x, jy=junction_y,
            out_dx=out_dx, out_dy=out_dy, stop_s=stop_s,
        ))
        if turn_arrows:
            cl, cs, cr = _allowed_turns_for_arm(
                junction=junc, arm=arm, tee_side=tee_side,
            )
            _append_turn_arrow_metas(
                out, jx=junction_x, jy=junction_y,
                out_dx=out_dx, out_dy=out_dy, stop_s=stop_s,
                travel_w=pw, lane_width_ft=lane_width_ft,
                lanes_toward=p_lanes_toward, lanes_out=p_lanes_out,
                arm=arm,
                road_type=primary_road_type,
                yellow_gap_ft=yellow_gap_ft,
                median_width_ft=primary_median_width_ft,
                twlt_width_ft=primary_twlt_width_ft,
                can_left=cl, can_straight=cs, can_right=cr,
            )

    _primary_arm(-1.0, "primary_neg")
    _primary_arm(+1.0, "primary_pos")

    def _stub(dir_x: float, dir_y: float, arm: str) -> None:
        # Edge lines start at the box so arms visually connect.
        s0x = junction_x + dir_x * primary_half
        s0y = junction_y + dir_y * primary_half
        s1x = junction_x + dir_x * (primary_half + secondary_stub_ft)
        s1y = junction_y + dir_y * (primary_half + secondary_stub_ft)
        sx1, sy1, sx2, sy2 = _first_edge_from_centerline(
            s0x, s0y, s1x, s1y, sw, side,
        )
        raw = _tag(
            build_strip_lines(secondary_road_type, sx1, sy1, sx2, sy2, **strip_kw_s),
            arm,
        )
        stop_s = _approach_stop_and_crosswalk(
            out, jx=junction_x, jy=junction_y,
            out_dx=dir_x, out_dy=dir_y,
            box_edge_ft=primary_half, travel_w=sw, arm=arm,
            crosswalk=crosswalks, stop_bar=stop_bars,
        )
        out.extend(_clip_center_lane_before_stop(
            raw, jx=junction_x, jy=junction_y,
            out_dx=dir_x, out_dy=dir_y, stop_s=stop_s,
        ))
        if turn_arrows:
            cl, cs, cr = _allowed_turns_for_arm(
                junction=junc, arm=arm, tee_side=tee_side,
            )
            _append_turn_arrow_metas(
                out, jx=junction_x, jy=junction_y,
                out_dx=dir_x, out_dy=dir_y, stop_s=stop_s,
                travel_w=sw, lane_width_ft=lane_width_ft,
                lanes_toward=s_lanes_toward, lanes_out=s_lanes_out,
                arm=arm,
                road_type=secondary_road_type,
                yellow_gap_ft=yellow_gap_ft,
                median_width_ft=secondary_median_width_ft,
                twlt_width_ft=secondary_twlt_width_ft,
                can_left=cl, can_straight=cs, can_right=cr,
            )

    if junc == "plus":
        stub_dirs = [
            (stx, sty, "secondary_left"),
            (-stx, -sty, "secondary_right"),
        ]
    else:
        ts = (tee_side or "right").strip().lower()
        if ts == "left":
            stub_dirs = [(stx, sty, "secondary_tee")]
        elif ts == "right":
            stub_dirs = [(-stx, -sty, "secondary_tee")]
        else:
            raise ValueError("tee_side must be 'left' or 'right'")

    for dx, dy, arm in stub_dirs:
        _stub(dx, dy, arm)

    if turning:
        _append_dotted_center(
            out, arm="center_extension_primary",
            x1=junction_x - ptx * secondary_half,
            y1=junction_y - pty * secondary_half,
            x2=junction_x + ptx * secondary_half,
            y2=junction_y + pty * secondary_half,
        )
        _append_dotted_center(
            out, arm="center_extension_secondary",
            x1=junction_x - stx * primary_half,
            y1=junction_y - sty * primary_half,
            x2=junction_x + stx * primary_half,
            y2=junction_y + sty * primary_half,
        )

    return out


def ramp_gore_lines(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    mainline_lanes: int,
    ramp_angle_deg: float,
    gore_station_ft: float,
    ramp_length_ft: float,
    ramp_lanes: int = 1,
    side: Literal["right", "left"] = "right",
    gore_mark_ft: float = 40.0,
    lane_width_ft: float = 12.0,
    shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0,
    gap_ft: float = 30.0,
) -> list[dict]:
    """Mainline one-way + diverging ramp one-way meeting at a gore nose.

    (x1,y1)->(x2,y2) is the mainline first travel outer edge (full length).
    Gore nose sits on the ramp-side travel outer edge at gore_station_ft
    from the start. Ramp diverges by ramp_angle_deg toward `side`.
    """
    if mainline_lanes < 1 or ramp_lanes < 1:
        raise ValueError("mainline_lanes and ramp_lanes must be >= 1")
    if ramp_angle_deg <= 0 or ramp_angle_deg >= 90:
        raise ValueError("ramp_angle_deg must be in (0, 90)")
    if gore_station_ft < 0:
        raise ValueError("gore_station_ft must be >= 0")
    if ramp_length_ft <= 0:
        raise ValueError("ramp_length_ft must be > 0")
    if gore_mark_ft < 0:
        raise ValueError("gore_mark_ft must be >= 0")

    length, tx, ty, nx, ny = _corridor_frame(x1, y1, x2, y2, side)
    if gore_station_ft > length + 1e-6:
        raise ValueError(
            f"gore_station_ft ({gore_station_ft}) exceeds mainline length ({length})"
        )

    mainline = _tag(
        lane_highway_lines(
            mainline_lanes, x1, y1, x2, y2,
            lane_width_ft=lane_width_ft,
            shoulder_width_ft=shoulder_width_ft,
            dash_ft=dash_ft, gap_ft=gap_ft, side=side,
        ),
        "mainline",
    )

    travel_w = float(mainline_lanes) * float(lane_width_ft)
    ramp_edge_off = travel_w if side == "right" else 0.0
    nose_x = x1 + nx * ramp_edge_off + tx * gore_station_ft
    nose_y = y1 + ny * ramp_edge_off + ty * gore_station_ft

    ang = -float(ramp_angle_deg) if side == "right" else float(ramp_angle_deg)
    rtx, rty = _rotate(tx, ty, ang)

    rx2 = nose_x + rtx * ramp_length_ft
    ry2 = nose_y + rty * ramp_length_ft
    ramp_side: Literal["right", "left"] = "right"
    _, _, _, rnx, rny = _corridor_frame(nose_x, nose_y, rx2, ry2, "right")
    if rnx * nx + rny * ny < 0:
        ramp_side = "left"

    ramp = _tag(
        lane_highway_lines(
            ramp_lanes, nose_x, nose_y, rx2, ry2,
            lane_width_ft=lane_width_ft,
            shoulder_width_ft=shoulder_width_ft,
            dash_ft=dash_ft, gap_ft=gap_ft, side=ramp_side,
        ),
        "ramp",
    )

    out: list[dict] = list(mainline) + list(ramp)
    if gore_mark_ft > 0:
        for i, (bx, by) in enumerate(((tx, ty), (rtx, rty))):
            out.append({
                "style": "solid", "kind": "gore", "row": 900 + i,
                "arm": "gore", "seg": i,
                "x1": nose_x, "y1": nose_y,
                "x2": nose_x - bx * gore_mark_ft,
                "y2": nose_y - by * gore_mark_ft,
            })
    out.append({
        "style": "meta", "kind": "gore_nose", "row": -1, "arm": "gore",
        "x1": nose_x, "y1": nose_y, "x2": nose_x, "y2": nose_y,
        "goreStationFt": gore_station_ft,
        "rampAngleDeg": ramp_angle_deg,
        "side": side,
    })
    return out


def strip_placeable_segments(segs: list[dict]) -> list[dict]:
    """Drop meta-only markers before placing in CAD."""
    return [s for s in segs if s.get("style") != "meta"]
