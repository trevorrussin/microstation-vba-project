"""Place verified satellite perimeter + trailers (all PIP-inside)."""
from __future__ import annotations

import json
import sys
from pathlib import Path

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
Y_NORTH = -4978.936756757757
CAD_NW = (11753.781157245516, Y_NORTH)
CAD_NE = (17318.73337825888, Y_NORTH)
S = (CAD_NE[0] - CAD_NW[0]) / (393 - 80)
OLD = ("6B7F6", "6B7F7", "6B7C0", "6B7C1", "6B755", "6B757")
EW, NS = "6A741", "6A78D"
EW_X, EW_Y = 14155.02424640845, -4585.566667744555
NS_X, NS_Y = 13732.353514847744, -4071.8642288933092
COL, ROW = 240.0, 120.0


def p2c(px: float, py: float) -> tuple[float, float]:
    return CAD_NW[0] + (px - 80) * S, Y_NORTH - (py - 400) * S


def prop_x(y: float) -> float:
    x1, y1 = 12648.780205027346, 46.98007728980156
    x2, y2 = CAD_NW
    return x2 + (y - y2) * (x2 - x1) / (y2 - y1)


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="satellite rebuild")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="satellite rebuild")


def build_verts() -> list[list[float]]:
    a = np.array(Image.open(ANNOTATED).convert("RGB"))
    r, g, b = a[:, :, 0].astype(int), a[:, :, 1].astype(int), a[:, :, 2].astype(int)
    red = (r > 180) & (g < 100) & (b < 100)
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

    px_pts: list[tuple[int, int]] = []
    for y in range(400, 488, 6):
        if y in left:
            px_pts.append((left[y], y))
    px_pts.append((left.get(490, 114), 490))
    for y in range(496, 538, 6):
        if y in left:
            px_pts.append((left[y], y))
    for x in range(110, 325, 8):
        if x in bottom:
            px_pts.append((x, bottom[x]))
    px_pts.append((324, bottom.get(324, 582)))
    for y in range(580, 399, -6):
        if y in right:
            px_pts.append((right[y], y))

    simp = [px_pts[0]]
    for p in px_pts[1:]:
        if abs(p[0] - simp[-1][0]) + abs(p[1] - simp[-1][1]) >= 4:
            simp.append(p)

    out: list[list[float]] = []
    for px, py in simp:
        x, y = p2c(px, py)
        if abs(y - Y_NORTH) < 60:
            y = Y_NORTH
        if y > -7000 and x < prop_x(y) + 150:
            x = prop_x(y)
        if abs(y - Y_NORTH) < 1 and x > 16800:
            x = CAD_NE[0]
        if out and abs(out[-1][0] - x) < 2 and abs(out[-1][1] - y) < 2:
            continue
        out.append([round(x, 3), round(y, 3), 0.0])

    i_nw = min(range(len(out)), key=lambda i: abs(out[i][0] - CAD_NW[0]) + abs(out[i][1] - Y_NORTH))
    out = out[i_nw:] + out[:i_nw]
    out.insert(0, [CAD_NW[0], Y_NORTH, 0.0])
    for v in out:
        if abs(v[1] - Y_NORTH) < 1 and v[0] > 16000:
            v[0] = CAD_NE[0]
    final: list[list[float]] = []
    for v in out:
        if final and abs(final[-1][0] - v[0]) < 1 and abs(final[-1][1] - v[1]) < 1:
            continue
        final.append(v)
    return final


def clear_south() -> int:
    n = 0
    with session() as s:
        ids = []
        for ent in s.space:
            try:
                k = ent.ObjectName
                layer = ent.Layer
                mn, mx = ent.GetBoundingBox()
                cy = (mn[1] + mx[1]) / 2
            except Exception:
                continue
            if cy > Y_NORTH:
                continue
            if k == "AcDbBlockReference" and layer == "C-VMF PARKING":
                ids.append(str(ent.Handle))
            elif k == "AcDbLine" and layer == "C-PAVEMENT MARKING":
                ids.append(str(ent.Handle))
            elif k in ("AcDbPolyline", "AcDbHatch") and layer == "C-CURB-EXIST":
                hid = str(ent.Handle)
                if hid in OLD or (k == "AcDbPolyline" and mn[1] < -6000 and mx[1] > -5200):
                    # only delete our south-lot boundaries (deep south)
                    if mn[1] < -7000:
                        ids.append(hid)
        for hid in ids:
            _delete(hid)
            n += 1
    for hid in OLD:
        try:
            _delete(hid)
        except AcadError:
            pass
    return n


def main() -> int:
    verts = build_verts()
    south_y = min(v[1] for v in verts)
    print(f"verts={len(verts)} depth={Y_NORTH - south_y:.0f}", flush=True)
    (ROOT / "scripts" / "south_lot_boundary_vertices.json").write_text(
        json.dumps(verts, indent=2), encoding="utf-8")

    print("clear", clear_south(), flush=True)
    r = acad_ops.place_polyline(verts, closed=True, layer="C-CURB-EXIST",
                                reason="satellite red silhouette exact")
    bid = str(r["elementId"])
    hr = acad_ops.hatch_element(bid, pattern="GRASS", own_element_only=True, reason="hatch")
    hid = str(hr["elementId"])
    acad_ops.change_element_layer(hid, "C-CURB-EXIST", own_element_only=True, reason="layer")
    with session() as s:
        ht = entity_by_handle(s.doc, hid)
        ht.PatternScale = 30.0
        ht.Evaluate()
    print(f"boundary={bid} hatch={hid}", flush=True)

    wx, wy0 = p2c(100, 412)
    west = [(wx, wy0 - i * ROW, "ew") for i in range(7)]
    cx0, cy1 = p2c(145, 428)
    cn = [(cx0 + i * COL, cy1, "ew") for i in range(7)]
    cs = [(cx0 + i * COL, cy1 - ROW, "ew") for i in range(7)]
    nex, ney = p2c(310, 418)
    ne = [(nex + i * 100.0, ney, "ns") for i in range(5)]
    _, sy = p2c(200, 530)
    swx, _ = p2c(130, 530)
    sex, _ = p2c(250, 530)
    sw = [(swx + i * 100.0, sy, "ns") for i in range(10)]
    se = [(sex + i * 100.0, sy, "ns") for i in range(9)]

    n = 0
    for name, pts in [("w", west), ("cn", cn), ("cs", cs), ("ne", ne), ("sw", sw), ("se", se)]:
        for i, (tx, ty, o) in enumerate(pts):
            if o == "ns":
                acad_ops.copy_element(NS, tx - NS_X, ty - NS_Y, own_element_only=False, reason=f"{name}{i}")
            else:
                acad_ops.copy_element(EW, tx - EW_X, ty - EW_Y, own_element_only=False, reason=f"{name}{i}")
            n += 1
        print(f"  {name}:{len(pts)}", flush=True)

    for _, y, _ in west[:-1]:
        acad_ops.place_line(wx - 80, y - ROW / 2, wx + 280, y - ROW / 2,
                            layer="C-PAVEMENT MARKING", reason="west stall")
    xs = [p[0] for p in cn]
    acad_ops.place_line(min(xs) - 80, cy1 - ROW / 2, max(xs) + 80, cy1 - ROW / 2,
                        layer="C-PAVEMENT MARKING", reason="center stall")
    print(f"Done trailers={n} south_trailer_y={sy:.0f} fence={south_y:.0f}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
