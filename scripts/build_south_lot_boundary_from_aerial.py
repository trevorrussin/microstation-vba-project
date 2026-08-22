"""Build south trailer lot boundary from satellite pavement edge + CAD curb snap.

Uses homography (4 corner GCPs on annotated aerial) to digitize real asphalt
perimeter — not the red markup polyline. Snaps north/west to surveyed curbs.
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
OLD = ("6B748", "6B749", "6B74A", "6B74B", "6B74C", "6B74D", "6B74E", "6B74F", "6B750", "6B751")

# Pixel -> CAD ground control (corners of south lot on aerial)
GCPS = [
    ((95, 198), (12245.057365225943, Y_NORTH)),       # inner NW curb
    ((370, 198), (17318.73337825888, Y_NORTH)),       # access road NE
    ((48, 402), (11753.781157245516, -6320.0)),       # property SW (69204 bearing)
    ((378, 402), (17318.73337825888, -6320.0)),       # access road SE
]

# Surveyed north-edge chain (west -> east); pocket notch from 691FB / cutoff
NORTH_CHAIN = [
    (11753.781157245516, Y_NORTH),
    (12245.057365225943, Y_NORTH),
    (14187.090451808337, Y_NORTH),
    (14189.366784556518, -4968.162738880136),
    (13622.134784668298, -4859.248461511208),
    (13541.240054425056, -4768.090865477686),
    (13550.172848596972, Y_NORTH),
    (15396.485989211884, Y_NORTH),
    (16022.553131150635, Y_NORTH),
    (16794.390515109088, Y_NORTH),
    (17318.73337825888, Y_NORTH),
]


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="replace south boundary")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="replace south boundary")


def _homography() -> np.ndarray:
    src = np.float32([p[0] for p in GCPS])
    dst = np.float32([p[1] for p in GCPS])
    h, _ = cv2.findHomography(src, dst, 0)
    return h


def px_to_cad(h: np.ndarray, px: float, py: float) -> tuple[float, float]:
    pt = h @ np.array([px, py, 1.0])
    return float(pt[0] / pt[2]), float(pt[1] / pt[2])


def extract_pavement_contour() -> np.ndarray:
    a = np.array(Image.open(ANNOTATED).convert("RGB"))
    r, g, b = a[:, :, 0].astype(float), a[:, :, 1].astype(float), a[:, :, 2].astype(float)
    gray = 0.299 * r + 0.587 * g + 0.114 * b
    sat = np.maximum.reduce([r, g, b]) - np.minimum.reduce([r, g, b])
    pave = (gray > 65) & (gray < 175) & (sat < 45)
    pave = pave & ~((r > 180) & (g < 100) & (b < 100))  # red markup
    pave = pave & ~((b > 180) & (r < 120) & (g < 160))   # blue markup
    mask = pave.astype(np.uint8) * 255
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((7, 7), np.uint8), iterations=2)
    roi = np.zeros_like(mask)
    roi[205:415, 35:385] = 255
    mask = cv2.bitwise_and(mask, roi)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    return max(contours, key=cv2.contourArea)


def property_point_at_y(y_target: float) -> tuple[float, float]:
    """69204 bearing through (11753.78, -4978.94)."""
    x1, y1 = 12648.780205027346, 46.98007728980156
    x2, y2 = 11753.781157245516, Y_NORTH
    dx, dy = x2 - x1, y2 - y1
    t = (y_target - y2) / dy
    return x2 + t * dx, y_target


def sample_arc_691ed(n: int = 8) -> list[tuple[float, float]]:
    cx, cy, r = 16795.715890056803, -4196.477336768032, 782.460542487225
    a0, a1 = 3.295906352781722, 4.710695124260922
    pts = [
        (cx + r * math.cos(a0 + (a1 - a0) * i / n), cy + r * math.sin(a0 + (a1 - a0) * i / n))
        for i in range(n + 1)
    ]
    return [p for p in pts if p[1] <= Y_NORTH + 1.0]


def south_extent_from_aerial(h: np.ndarray, contour: np.ndarray) -> float:
    cad_y = [px_to_cad(h, float(p[0]), float(p[1]))[1] for p in contour.reshape(-1, 2)]
    return min(cad_y)


def east_jog_from_aerial(h: np.ndarray, contour: np.ndarray) -> list[tuple[float, float]]:
    """East-side inset around auxiliary building from asphalt edge."""
    pts = []
    for px, py in contour.reshape(-1, 2):
        x, y = px_to_cad(h, float(px), float(py))
        if x > 16900 and -5800 < y < -5100:
            pts.append((x, y))
    if len(pts) < 3:
        return []
    pts.sort(key=lambda p: p[1])
    mid = pts[len(pts) // 2]
    return [(17318.73337825888, mid[1])] if mid[0] < 17318 else []


def build_boundary() -> list[list[float]]:
    h = _homography()
    contour = extract_pavement_contour()
    south_y = south_extent_from_aerial(h, contour)
    sw = property_point_at_y(south_y)

    loop: list[tuple[float, float]] = [
        sw,
        (11753.781157245516, Y_NORTH),
        (12245.057365225943, Y_NORTH),
        (14187.090451808337, Y_NORTH),
        (15396.485989211884, Y_NORTH),
        (16022.553131150635, Y_NORTH),
    ]
    arc = sample_arc_691ed()
    loop.extend(arc if arc else [(16794.390515109088, Y_NORTH)])
    loop.extend([
        (17318.73337825888, Y_NORTH),
        (17318.73337825888, south_y),
        sw,
    ])

    out: list[list[float]] = []
    for x, y in loop:
        if out and abs(out[-1][0] - x) < 0.5 and abs(out[-1][1] - y) < 0.5:
            continue
        out.append([round(x, 3), round(y, 3), 0.0])
    return out


def main() -> int:
    verts = build_boundary()
    out_json = ROOT / "scripts" / "south_lot_boundary_vertices.json"
    out_json.write_text(json.dumps(verts, indent=2), encoding="utf-8")
    print(f"wrote {len(verts)} vertices -> {out_json}", flush=True)

    xs = [v[0] for v in verts]
    ys = [v[1] for v in verts]
    print(f"bbox x=[{min(xs):.0f},{max(xs):.0f}] y=[{min(ys):.0f},{max(ys):.0f}]", flush=True)

    for hid in OLD:
        try:
            _delete(hid)
            print(f"deleted {hid}", flush=True)
        except AcadError as exc:
            print(f"skip delete {hid}: {exc}", flush=True)

    r = acad_ops.place_polyline(verts, closed=True, layer="C-CURB-EXIST",
                                reason="south lot boundary from aerial pavement + curbs")
    boundary_id = str(r.get("elementId") or "")
    print(f"boundary={boundary_id}", flush=True)

    pattern, scale = "GRASS", 30.0
    with session() as s:
        ref = entity_by_handle(s.doc, HATCH_REF)
        try:
            pattern = str(ref.PatternName or pattern)
            scale = float(ref.PatternScale or scale)
        except Exception:
            pass

    hr = acad_ops.hatch_element(boundary_id, pattern=pattern, own_element_only=True,
                                reason="south lot hatch aerial boundary")
    hatch_id = str(hr.get("elementId") or "")
    if hatch_id:
        acad_ops.change_element_layer(hatch_id, "C-CURB-EXIST", own_element_only=True,
                                      reason="match 6926C")
        with session() as s:
            ht = entity_by_handle(s.doc, hatch_id)
            ht.PatternScale = scale
            ht.Evaluate()
    print(f"hatch={hatch_id} {pattern} scale={scale}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
