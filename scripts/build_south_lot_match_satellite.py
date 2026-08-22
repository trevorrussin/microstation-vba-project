"""Rebuild south trailer lot to match annotated satellite exactly.

Root cause of prior misses:
  - Red outline lives in the BOTTOM of Screenshot_2026-08-20_021119 (y=400..582),
    not the middle ROI used before.
  - Isotropic scale from north curb width (~17.78 ft/px) puts south fence near
    y=-8180 (~3200 ft deep), not the guessed y=-6300 (~1413 ft).
  - Outer perimeter = red L/R silhouette (closed), not red-stroke ribbon.

Places: closed curb + GRASS hatch + satellite trailer clusters + stall lines.
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import cv2
import numpy as np
from PIL import Image

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session, entity_by_handle  # noqa: E402

ANNOTATED = Path(
    r"C:\Users\RussinT\.cursor\projects\c-repos-microstation-vba-project\assets"
    r"\c__Users_RussinT_AppData_Roaming_Cursor_User_workspaceStorage_"
    r"680ae203d5859150138f6b5c22b7ce01_images_Screenshot_2026-08-20_021119-"
    r"45a530fe-f1f7-41b3-95fe-91b42a65f54b.png"
)
HATCH_REF = "6926C"
Y_NORTH = -4978.936756757757
# North-edge CAD anchors (surveyed curb / access road)
CAD_NW = (11753.781157245516, Y_NORTH)  # 69204 at cutoff
CAD_NE = (17318.73337825888, Y_NORTH)   # 691EE east
# Pixel north-edge of filled red lot (full-width band starts ~400)
PX_NW = (80.0, 400.0)
PX_NE = (393.0, 406.0)

EW_TEMPLATE = "6A741"
NS_TEMPLATE = "6A78D"
EW_X, EW_Y = 14155.02424640845, -4585.566667744555
NS_X, NS_Y = 13732.353514847744, -4071.8642288933092
COL, ROW = 240.0, 120.0

OLD_BOUNDARY = (
    "6B748", "6B749", "6B74A", "6B74B", "6B74C", "6B74D",
    "6B74E", "6B74F", "6B750", "6B751", "6B752", "6B753",
)


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="south lot satellite rebuild")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="south lot satellite rebuild")


def property_x_at_y(y: float) -> float:
    x1, y1 = 12648.780205027346, 46.98007728980156
    x2, y2 = CAD_NW
    return x2 + (y - y2) * (x2 - x1) / (y2 - y1)


def load_red_lr() -> tuple[dict[int, int], dict[int, int]]:
    a = np.array(Image.open(ANNOTATED).convert("RGB"))
    r, g, b = a[:, :, 0].astype(int), a[:, :, 1].astype(int), a[:, :, 2].astype(int)
    red = (r > 180) & (g < 100) & (b < 100)
    left: dict[int, int] = {}
    right: dict[int, int] = {}
    for y in range(400, 583):
        row = np.where(red[y])[0]
        if len(row) == 0:
            continue
        left[y] = int(row.min())
        right[y] = int(row.max())
    return left, right


def pixel_to_cad(px: float, py: float) -> tuple[float, float]:
    """Isotropic map: north curb width sets ft/px; image +Y -> CAD -Y."""
    sx = (CAD_NE[0] - CAD_NW[0]) / (PX_NE[0] - PX_NW[0])
    x = CAD_NW[0] + (px - PX_NW[0]) * sx
    y = Y_NORTH - (py - PX_NW[1]) * sx
    return x, y


def build_perimeter_px(left: dict[int, int], right: dict[int, int]) -> np.ndarray:
    """Closed outer silhouette of red lot (clockwise from NW)."""
    ys = sorted(left.keys())
    # South fence starts where leftmost begins marching east
    y_south = 538
    for y in ys:
        if y >= 520 and left[y] > left.get(y - 6, left[y]) + 8:
            y_south = y
            break

    pts: list[list[float]] = []
    # West edge north -> south (tree line + SW notch)
    for y in ys:
        if y > y_south:
            break
        pts.append([left[y], y])
    # South fence west -> east (leftmost while it climbs east)
    for y in ys:
        if y < y_south:
            continue
        pts.append([left[y], y])
    # SE tip
    y_max = max(ys)
    pts.append([right[y_max], y_max])
    # East edge south -> north
    for y in reversed(ys):
        pts.append([right[y], y])
    # Close to NW
    pts.append([left[ys[0]], ys[0]])

    arr = np.array(pts, dtype=np.float32).reshape(-1, 1, 2)
    simp = cv2.approxPolyDP(arr, 2.0, True).reshape(-1, 2)
    return simp


def perimeter_to_cad(simp: np.ndarray) -> list[list[float]]:
    out: list[list[float]] = []
    for px, py in simp:
        x, y = pixel_to_cad(float(px), float(py))
        # Snap north band to surveyed cutoff
        if abs(y - Y_NORTH) < 40:
            y = Y_NORTH
        # Snap west tree-line (north of SW notch) to property bearing
        if y > -6600 and x < property_x_at_y(y) + 80:
            x = property_x_at_y(y)
        # Snap NE to access road
        if abs(y - Y_NORTH) < 1 and x > 17000:
            x = CAD_NE[0]
        out.append([round(x, 3), round(y, 3), 0.0])

    # Ensure closed without duplicate last==first for AutoCAD Closed=True
    if len(out) > 1 and abs(out[0][0] - out[-1][0]) < 0.5 and abs(out[0][1] - out[-1][1]) < 0.5:
        out = out[:-1]
    return out


def clear_south_lot() -> tuple[int, int]:
    tr = st = 0
    with session() as s:
        ids: list[tuple[str, str]] = []
        for ent in s.space:
            try:
                kind = ent.ObjectName
                layer = ent.Layer
                min_pt, max_pt = ent.GetBoundingBox()
                cy = (min_pt[1] + max_pt[1]) / 2.0
            except Exception:
                continue
            if cy > Y_NORTH:
                continue
            hid = str(ent.Handle)
            if kind == "AcDbBlockReference" and layer == "C-VMF PARKING":
                ids.append(("t", hid))
            elif kind == "AcDbLine" and layer == "C-PAVEMENT MARKING":
                ids.append(("s", hid))
            elif kind == "AcDbPolyline" and layer == "C-CURB-EXIST" and hid in OLD_BOUNDARY:
                ids.append(("b", hid))
            elif kind == "AcDbHatch" and layer == "C-CURB-EXIST" and hid in OLD_BOUNDARY:
                ids.append(("b", hid))
        for kind, hid in ids:
            _delete(hid)
            tr += kind == "t"
            st += kind == "s"
    for hid in OLD_BOUNDARY:
        try:
            _delete(hid)
        except AcadError:
            pass
    return tr, st


def place_trailer(tx: float, ty: float, orient: str, label: str) -> str:
    if orient == "ns":
        r = acad_ops.copy_element(NS_TEMPLATE, tx - NS_X, ty - NS_Y,
                                  own_element_only=False, reason=label)
    else:
        r = acad_ops.copy_element(EW_TEMPLATE, tx - EW_X, ty - EW_Y,
                                  own_element_only=False, reason=label)
    return str(r.get("elementId") or "")


def trailer_clusters(south_y: float) -> dict[str, list[tuple[float, float, str]]]:
    """Satellite clusters mapped into the rebuilt lot (drive aisles left open)."""
    # West edge column — top-left only (~7 E-W)
    west_x = CAD_NW[0] + 160.0
    west = [(west_x, Y_NORTH - 110 - i * ROW, "ew") for i in range(7)]

    # Central back-to-back island — upper middle (7+7 E-W)
    cx0 = CAD_NW[0] + 1100.0
    c_y1 = Y_NORTH - 280.0
    c_y2 = c_y1 - ROW
    center_n = [(cx0 + i * COL, c_y1, "ew") for i in range(7)]
    center_s = [(cx0 + i * COL, c_y2, "ew") for i in range(7)]

    # NE entry — short N-S row just below north cutoff
    ne_x0 = CAD_NE[0] - 2200.0
    ne = [(ne_x0 + i * 100.0, Y_NORTH - 160.0, "ns") for i in range(5)]

    # South fence — two N-S blocks with center drive gap
    sy = south_y + 160.0
    sw = [(CAD_NW[0] + 450 + i * 100.0, sy, "ns") for i in range(10)]
    se = [(CAD_NE[0] - 2200 + i * 100.0, sy, "ns") for i in range(9)]

    return {
        "west_edge": west,
        "center_north": center_n,
        "center_south": center_s,
        "ne_entry": ne,
        "south_west": sw,
        "south_east": se,
    }


def stall_lines(clusters: dict[str, list[tuple[float, float, str]]]) -> list[tuple[float, float, float]]:
    lines: list[tuple[float, float, float]] = []
    west = clusters["west_edge"]
    if west:
        x = west[0][0]
        for _, y, _ in west[:-1]:
            lines.append((x - 80, x + 280, y - ROW / 2))
    cn = clusters["center_north"]
    if cn:
        xs = [p[0] for p in cn]
        lines.append((min(xs) - 80, max(xs) + 80, cn[0][1] - ROW / 2))
    return lines


def main() -> int:
    left, right = load_red_lr()
    simp = build_perimeter_px(left, right)
    verts = perimeter_to_cad(simp)
    south_y = min(v[1] for v in verts)
    north_y = max(v[1] for v in verts)
    xs = [v[0] for v in verts]
    print(f"perimeter verts={len(verts)} closed=True", flush=True)
    print(f"bbox x=[{min(xs):.0f},{max(xs):.0f}] y=[{south_y:.0f},{north_y:.0f}]", flush=True)
    print(f"depth_ft={north_y - south_y:.0f} width_ft={max(xs) - min(xs):.0f}", flush=True)

    out_json = ROOT / "scripts" / "south_lot_boundary_vertices.json"
    out_json.write_text(json.dumps(verts, indent=2), encoding="utf-8")

    print("=== Clear old south lot ===", flush=True)
    tr, st = clear_south_lot()
    print(f"  removed trailers={tr} stall_lines={st}", flush=True)

    print("=== Boundary + hatch ===", flush=True)
    r = acad_ops.place_polyline(verts, closed=True, layer="C-CURB-EXIST",
                                reason="south lot perimeter from satellite red silhouette")
    boundary_id = str(r.get("elementId") or "")
    print(f"  boundary={boundary_id}", flush=True)

    pattern, scale = "GRASS", 30.0
    with session() as s:
        ref = entity_by_handle(s.doc, HATCH_REF)
        try:
            pattern = str(ref.PatternName or pattern)
            scale = float(ref.PatternScale or scale)
        except Exception:
            pass
    hr = acad_ops.hatch_element(boundary_id, pattern=pattern, own_element_only=True,
                                reason="south lot hatch satellite perimeter")
    hatch_id = str(hr.get("elementId") or "")
    if hatch_id:
        acad_ops.change_element_layer(hatch_id, "C-CURB-EXIST", own_element_only=True,
                                      reason="match 6926C")
        with session() as s:
            ht = entity_by_handle(s.doc, hatch_id)
            ht.PatternScale = scale
            ht.Evaluate()
    print(f"  hatch={hatch_id} {pattern}@{scale}", flush=True)

    clusters = trailer_clusters(south_y)
    total = sum(len(v) for v in clusters.values())
    print(f"=== Place {total} trailers ===", flush=True)
    n = 0
    for name, pts in clusters.items():
        print(f"  {name}: {len(pts)}", flush=True)
        for i, (tx, ty, orient) in enumerate(pts):
            place_trailer(tx, ty, orient, f"{name}[{i}]")
            n += 1

    print("=== Stall lines ===", flush=True)
    for x1, x2, y in stall_lines(clusters):
        acad_ops.place_line(x1, y, x2, y, layer="C-PAVEMENT MARKING",
                            reason=f"cluster stall y={y}")

    print(f"Done: boundary={boundary_id} hatch={hatch_id} trailers={n}", flush=True)
    print(f"Pan to y≈{south_y:.0f}..{north_y:.0f} (lot is deeper than viewport crop).", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
