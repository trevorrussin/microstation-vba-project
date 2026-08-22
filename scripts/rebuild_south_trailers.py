"""Rebuild south lot trailers per engineer layout — do NOT touch boundary/hatch.

Counts: 10 west (E-W), 19 south (N-S), 5+5 center back-to-back (E-W).
Remove NE entry row. Scale vs upper-lot blocks: 1.25x wider, 2x longer.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session, entity_by_handle  # noqa: E402

Y_NORTH = -4978.936756757757
BOUNDARY = "6B7F6"
EW_TEMPLATE = "6A741"
NS_TEMPLATE = "6A78D"
WIDER = 1.25
LONGER = 3.0  # 2.0 original * 1.5 engineer request 2026-08-20
INSET = 8.0  # ft inside curb


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="south trailer rebuild")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="south trailer rebuild")


def clear_south_parking() -> tuple[int, int]:
    tr = st = 0
    with session() as s:
        ids: list[tuple[str, str]] = []
        for ent in s.space:
            try:
                kind = ent.ObjectName
                layer = ent.Layer
                mn, mx = ent.GetBoundingBox()
                cy = (mn[1] + mx[1]) / 2.0
            except Exception:
                continue
            if cy > Y_NORTH:
                continue
            hid = str(ent.Handle)
            if kind == "AcDbBlockReference" and layer == "C-VMF PARKING":
                ids.append(("t", hid))
            elif kind == "AcDbLine" and layer == "C-PAVEMENT MARKING":
                ids.append(("s", hid))
        for kind, hid in ids:
            _delete(hid)
            tr += kind == "t"
            st += kind == "s"
    return tr, st


def boundary_pts() -> list[tuple[float, float]]:
    with session() as s:
        ent = entity_by_handle(s.doc, BOUNDARY)
        c = list(ent.Coordinates)
    return [(c[i], c[i + 1]) for i in range(0, len(c), 2)]


def _lerp_seg(a: tuple[float, float], b: tuple[float, float], t: float) -> tuple[float, float]:
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)


def west_x_at_y(pts: list[tuple[float, float]], y: float) -> float:
    """West property edge: NW down through SW notch (before south fence runs east)."""
    chain = [pts[13], pts[14], pts[15], pts[0], pts[1]]
    for a, b in zip(chain, chain[1:]):
        ya, yb = sorted((a[1], b[1]))
        if ya - 1 <= y <= yb + 1:
            if abs(b[1] - a[1]) < 1e-6:
                return min(a[0], b[0])
            t = (y - a[1]) / (b[1] - a[1])
            return _lerp_seg(a, b, t)[0]
    return pts[14][0]


def south_chain(pts: list[tuple[float, float]]) -> list[tuple[float, float]]:
    return [pts[1], pts[2], pts[3], pts[4], pts[5], pts[6]]


def chain_lengths(chain: list[tuple[float, float]]) -> tuple[list[float], float]:
    lens: list[float] = []
    total = 0.0
    for a, b in zip(chain, chain[1:]):
        L = math.hypot(b[0] - a[0], b[1] - a[1])
        lens.append(L)
        total += L
    return lens, total


def point_on_chain(chain: list[tuple[float, float]], dist: float) -> tuple[float, float]:
    lens, total = chain_lengths(chain)
    d = max(0.0, min(dist, total))
    acc = 0.0
    for (a, b), L in zip(zip(chain, chain[1:]), lens):
        if acc + L >= d - 1e-6:
            t = (d - acc) / L if L else 0.0
            return _lerp_seg(a, b, t)
        acc += L
    return chain[-1]


EW_X, EW_Y = 14155.02424640845, -4585.566667744555
NS_X, NS_Y = 13732.353514847744, -4071.8642288933092


def place_scaled(template: str, tx: float, ty: float, ox: float, oy: float,
                 wider: float, longer: float) -> str:
    r = acad_ops.copy_element(template, tx - ox, ty - oy, own_element_only=False,
                              reason="south trailer")
    hid = str(r.get("elementId") or "")
    with session() as s:
        ent = entity_by_handle(s.doc, hid)
        ent.XScaleFactor = float(ent.XScaleFactor) * wider
        ent.YScaleFactor = float(ent.YScaleFactor) * longer
    return hid


def stall_line(x1: float, y: float, x2: float) -> None:
    acad_ops.place_line(x1, y, x2, y, layer="C-PAVEMENT MARKING", reason="south stall")


def main() -> int:
    pts = boundary_pts()
    chain = south_chain(pts)
    _, south_len = chain_lengths(chain)

    # Scaled trailer dims (EW: long=X, depth=Y)
    ew_len, ew_dep = 236.21452930715532 * LONGER, 92.34840715461269 * WIDER
    ns_w, ns_len = ew_dep, ew_len
    row_pitch = 120.0 * WIDER  # match upper-lot spacing scaled (~150 ft)

    tr, st = clear_south_parking()
    print(f"Cleared {tr} trailers, {st} stall lines", flush=True)

    west_centers: list[tuple[float, float]] = []
    y0 = Y_NORTH - 100.0 - ew_dep / 2
    for i in range(10):
        cy = y0 - i * row_pitch
        wx = west_x_at_y(pts, cy)
        cx = wx + INSET + ew_len / 2
        west_centers.append((cx, cy))
        place_scaled(EW_TEMPLATE, cx, cy, EW_X, EW_Y, WIDER, LONGER)
    print(f"West column: 10 @ x~{west_centers[0][0]:.0f}", flush=True)

    # Center: 2 columns of 5 E-W, back-to-back (not 2 rows of 5)
    aisle = 280.0
    west_max = max(c[0] + ew_len / 2 for c in west_centers)
    col1_cx = west_max + aisle + ew_len / 2
    col2_cx = col1_cx + ew_len  # abutting backs
    y0 = Y_NORTH - 100.0 - ew_dep / 2
    center_centers: list[tuple[float, float]] = []
    for col_cx in (col1_cx, col2_cx):
        for i in range(5):
            cy = y0 - i * row_pitch
            place_scaled(EW_TEMPLATE, col_cx, cy, EW_X, EW_Y, WIDER, LONGER)
            center_centers.append((col_cx, cy))
    print(f"Center: 2 cols x 5 @ cx={col1_cx:.0f}/{col2_cx:.0f}", flush=True)

    # South: 19 N-S along south fence (inside)
    south_centers: list[tuple[float, float]] = []
    margin = 60.0
    usable = south_len - 2 * margin
    for i in range(19):
        d = margin + usable * (i + 0.5) / 19
        fx, fy = point_on_chain(chain, d)
        cx, cy = fx, fy - INSET - ns_len / 2
        south_centers.append((cx, cy))
        place_scaled(NS_TEMPLATE, cx, cy, NS_X, NS_Y, WIDER, LONGER)
    print(f"South row: 19 along fence len={south_len:.0f}", flush=True)

    # Stall lines — same pattern as west column (yellow row separators)
    wx_min = west_centers[0][0] - ew_len / 2
    wx_max = west_centers[0][0] + ew_len / 2
    for i in range(len(west_centers) - 1):
        y = (west_centers[i][1] + west_centers[i + 1][1]) / 2
        stall_line(wx_min, y, wx_max)

    # Center stall lines between consecutive rows (span both columns)
    cx_min = col1_cx - ew_len / 2
    cx_max = col2_cx + ew_len / 2
    for i in range(4):
        y = (center_centers[i][1] + center_centers[i + 1][1]) / 2
        stall_line(cx_min, y, cx_max)
    # Vertical divider between the two back-to-back columns
    mid_x = (col1_cx + col2_cx) / 2
    y_top = center_centers[0][1] + ew_dep / 2
    y_bot = center_centers[4][1] - ew_dep / 2
    acad_ops.place_line(mid_x, y_bot, mid_x, y_top, layer="C-PAVEMENT MARKING",
                        reason="center aisle")

    # South stall line along fence (north of trailers)
    sx_min = min(c[0] - ns_w / 2 for c in south_centers)
    sx_max = max(c[0] + ns_w / 2 for c in south_centers)
    sy_line = max(c[1] + ns_len / 2 for c in south_centers) + 15
    stall_line(sx_min, sy_line, sx_max)

    total = 10 + 10 + 19
    print(f"Done: {total} trailers (472x115 ft EW, 115x472 ft NS). No NE row.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
