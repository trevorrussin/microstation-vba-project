"""Rebuild south lot perimeter so south fence is near-horizontal (satellite),
then keep trailers inside. Replaces shallow/diagonal south edge on 6B755.
"""
from __future__ import annotations

import json
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
CAD_NW = (11753.781157245516, Y_NORTH)
CAD_NE = (17318.73337825888, Y_NORTH)
PX_NW = (80.0, 400.0)
PX_NE = (393.0, 406.0)
OLD = ("6B755", "6B756", "6B757")

EW_TEMPLATE = "6A741"
NS_TEMPLATE = "6A78D"
EW_X, EW_Y = 14155.02424640845, -4585.566667744555
NS_X, NS_Y = 13732.353514847744, -4071.8642288933092
COL, ROW = 240.0, 120.0


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="fix south perimeter")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="fix south perimeter")


def property_x_at_y(y: float) -> float:
    x1, y1 = 12648.780205027346, 46.98007728980156
    x2, y2 = CAD_NW
    return x2 + (y - y2) * (x2 - x1) / (y2 - y1)


def sx() -> float:
    return (CAD_NE[0] - CAD_NW[0]) / (PX_NE[0] - PX_NW[0])


def p2c(px: float, py: float) -> tuple[float, float]:
    s = sx()
    return CAD_NW[0] + (px - PX_NW[0]) * s, Y_NORTH - (py - PX_NW[1]) * s


def red_mask() -> np.ndarray:
    a = np.array(Image.open(ANNOTATED).convert("RGB"))
    r, g, b = a[:, :, 0].astype(int), a[:, :, 1].astype(int), a[:, :, 2].astype(int)
    return (r > 180) & (g < 100) & (b < 100)


def build_perimeter_px(red: np.ndarray) -> np.ndarray:
    left: dict[int, int] = {}
    right: dict[int, int] = {}
    bottom: dict[int, int] = {}
    for y in range(400, 583):
        xs = np.where(red[y])[0]
        if len(xs) == 0:
            continue
        left[y] = int(xs.min())
        right[y] = int(xs.max())
        for x in xs:
            xi = int(x)
            bottom[xi] = max(bottom.get(xi, 0), y)

    ys = sorted(left)
    # SW corner = where bottommost is still on south fence (x<=324) and west joins
    south_xs = sorted(x for x, y in bottom.items() if 108 <= x <= 324 and y >= 530)

    pts: list[list[float]] = []
    # West edge north -> south (includes SW notch where L jumps east)
    for y in ys:
        if left[y] >= 108 and y >= 538:
            break
        pts.append([float(left[y]), float(y)])
    # South fence west -> east along bottommost red
    for x in south_xs:
        pts.append([float(x), float(bottom[x])])
    # East edge south -> north
    for y in reversed(ys):
        pts.append([float(right[y]), float(y)])
    # Close
    pts.append(pts[0][:])

    arr = np.array(pts, dtype=np.float32).reshape(-1, 1, 2)
    return cv2.approxPolyDP(arr, 2.5, True).reshape(-1, 2)


def perimeter_to_cad(simp: np.ndarray) -> list[list[float]]:
    out: list[list[float]] = []
    for px, py in simp:
        x, y = p2c(float(px), float(py))
        if abs(y - Y_NORTH) < 50:
            y = Y_NORTH
        if y > -6600 and x < property_x_at_y(y) + 100:
            x = property_x_at_y(y)
        if abs(y - Y_NORTH) < 1 and x > 17000:
            x = CAD_NE[0]
        if out and abs(out[-1][0] - x) < 1 and abs(out[-1][1] - y) < 1:
            continue
        out.append([round(x, 3), round(y, 3), 0.0])
    if len(out) > 1 and abs(out[0][0] - out[-1][0]) < 1 and abs(out[0][1] - out[-1][1]) < 1:
        out = out[:-1]
    return out


def clear_trailers_stalls() -> tuple[int, int]:
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


def place_trailer(tx: float, ty: float, orient: str, label: str) -> None:
    if orient == "ns":
        acad_ops.copy_element(NS_TEMPLATE, tx - NS_X, ty - NS_Y,
                              own_element_only=False, reason=label)
    else:
        acad_ops.copy_element(EW_TEMPLATE, tx - EW_X, ty - EW_Y,
                              own_element_only=False, reason=label)


def main() -> int:
    red = red_mask()
    simp = build_perimeter_px(red)
    verts = perimeter_to_cad(simp)
    south_y = min(v[1] for v in verts)
    xs = [v[0] for v in verts]
    print(f"verts={len(verts)} depth={Y_NORTH - south_y:.0f} "
          f"x=[{min(xs):.0f},{max(xs):.0f}] y=[{south_y:.0f},{Y_NORTH:.0f}]", flush=True)
    for i, v in enumerate(verts):
        print(f"  {i}: {v[0]:.1f},{v[1]:.1f}", flush=True)

    (ROOT / "scripts" / "south_lot_boundary_vertices.json").write_text(
        json.dumps(verts, indent=2), encoding="utf-8")

    for hid in OLD:
        try:
            _delete(hid)
            print(f"deleted {hid}", flush=True)
        except AcadError as e:
            print(f"skip {hid}: {e}", flush=True)

    r = acad_ops.place_polyline(verts, closed=True, layer="C-CURB-EXIST",
                                reason="south lot satellite silhouette (fixed south fence)")
    bid = str(r.get("elementId") or "")
    print(f"boundary={bid}", flush=True)

    pattern, scale = "GRASS", 30.0
    with session() as s:
        ref = entity_by_handle(s.doc, HATCH_REF)
        try:
            pattern = str(ref.PatternName or pattern)
            scale = float(ref.PatternScale or scale)
        except Exception:
            pass
    hr = acad_ops.hatch_element(bid, pattern=pattern, own_element_only=True,
                                reason="south lot hatch")
    hid = str(hr.get("elementId") or "")
    acad_ops.change_element_layer(hid, "C-CURB-EXIST", own_element_only=True, reason="match 6926C")
    with session() as s:
        ht = entity_by_handle(s.doc, hid)
        ht.PatternScale = scale
        ht.Evaluate()
    print(f"hatch={hid} {pattern}@{scale}", flush=True)

    tr, st = clear_trailers_stalls()
    print(f"cleared trailers={tr} stalls={st}", flush=True)

    # Clusters from satellite pixel map (all inside new perimeter)
    wx, wy0 = p2c(100, 412)
    west = [(wx, wy0 - i * ROW, "ew") for i in range(7)]
    cx0, cy1 = p2c(145, 428)
    cy2 = cy1 - ROW
    cn = [(cx0 + i * COL, cy1, "ew") for i in range(7)]
    cs = [(cx0 + i * COL, cy2, "ew") for i in range(7)]
    nex, ney = p2c(310, 418)
    ne = [(nex + i * 100.0, ney, "ns") for i in range(5)]
    # South blocks just north of south fence (~y_px 560 -> CAD)
    _, sy = p2c(200, 555)
    swx, _ = p2c(125, 555)
    sex, _ = p2c(245, 555)
    sw = [(swx + i * 100.0, sy, "ns") for i in range(10)]
    se = [(sex + i * 100.0, sy, "ns") for i in range(9)]
    print(f"south trailer y={sy:.0f} (fence={south_y:.0f}) gap={se[0][0]-sw[-1][0]:.0f}", flush=True)

    n = 0
    for name, pts in [("w", west), ("cn", cn), ("cs", cs), ("ne", ne), ("sw", sw), ("se", se)]:
        for i, (tx, ty, o) in enumerate(pts):
            place_trailer(tx, ty, o, f"{name}{i}")
            n += 1
        print(f"  {name}:{len(pts)}", flush=True)

    for _, y, _ in west[:-1]:
        acad_ops.place_line(wx - 80, y - ROW / 2, wx + 280, y - ROW / 2,
                            layer="C-PAVEMENT MARKING", reason="west stall")
    xs2 = [p[0] for p in cn]
    acad_ops.place_line(min(xs2) - 80, cy1 - ROW / 2, max(xs2) + 80, cy1 - ROW / 2,
                        layer="C-PAVEMENT MARKING", reason="center stall")

    print(f"Done boundary={bid} hatch={hid} trailers={n}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
