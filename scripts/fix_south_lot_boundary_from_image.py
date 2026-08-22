"""Trace south lot boundary from engineer's red outline on annotated satellite screenshot.

Extracts red polyline pixels -> CAD coords via bilinear quad calibrated to existing curbs.
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

# CAD anchors for red-outline quad corners (from existing curbs / 69204 bearing)
NW = (12245.057365225943, -4978.936756757757)   # inner SW curb north — red NW
NE = (17318.73337825888, -4978.936756757757)    # 691EE east
SE = (17318.73337825888, -6300.0)               # access road south
SW = (11518.53047885154, -6300.0)               # 69204 property bearing @ y=-6300

OLD_BOUNDARY = ("6B744", "6B745", "6B746", "6B747", "6B748", "6B749")


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="replace boundary")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="replace boundary")


def bilinear(nx: float, ny: float) -> tuple[float, float]:
    """Map normalised image coords (0-1) to CAD using quad corners."""
    corners = (SW, SE, NW, NE)  # (0,0), (1,0), (0,1), (1,1)
    x = sum(c[0] * w for c, w in [
        (corners[0], (1 - nx) * (1 - ny)),
        (corners[1], nx * (1 - ny)),
        (corners[2], (1 - nx) * ny),
        (corners[3], nx * ny),
    ])
    y = sum(c[1] * w for c, w in [
        (corners[0], (1 - nx) * (1 - ny)),
        (corners[1], nx * (1 - ny)),
        (corners[2], (1 - nx) * ny),
        (corners[3], nx * ny),
    ])
    return x, y


def extract_red_contour(eps: float = 1.0) -> np.ndarray:
    a = np.array(Image.open(ANNOTATED).convert("RGB"))
    r, g, b = a[:, :, 0].astype(int), a[:, :, 1].astype(int), a[:, :, 2].astype(int)
    mask = (r > 180) & (g < 100) & (b < 100)
    mask_u8 = mask.astype(np.uint8) * 255
    contours, _ = cv2.findContours(mask_u8, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    c = max(contours, key=cv2.contourArea)
    simp = cv2.approxPolyDP(c, eps, True).reshape(-1, 2)
    return simp.astype(float)


def contour_to_cad(pts_px: np.ndarray) -> list[list[float]]:
    px_min = pts_px.min(axis=0)
    px_max = pts_px.max(axis=0)
    span = px_max - px_min
    span[span == 0] = 1.0
    cad: list[list[float]] = []
    for px, py in pts_px:
        nx = (px - px_min[0]) / span[0]
        ny = (py - px_min[1]) / span[1]
        x, y = bilinear(nx, ny)
        cad.append([round(x, 3), round(y, 3), 0.0])
    return cad


def snap_north_edge(verts: list[list[float]], y_ref: float = -4978.937, tol: float = 12.0) -> None:
    """Snap only vertices already on the north cutoff (do not flatten NW diagonals)."""
    for v in verts:
        if abs(v[1] - y_ref) < tol:
            v[1] = y_ref


def main() -> int:
    for eps in (0.6, 0.8, 1.0):
        pts = extract_red_contour(eps)
        if 60 <= len(pts) <= 150:
            break
    print(f"contour vertices: {len(pts)} (eps={eps})", flush=True)

    cad_verts = contour_to_cad(pts)
    snap_north_edge(cad_verts)

    out_json = ROOT / "scripts" / "south_lot_boundary_vertices.json"
    out_json.write_text(json.dumps(cad_verts, indent=2), encoding="utf-8")
    print(f"wrote {out_json}", flush=True)

    for hid in OLD_BOUNDARY:
        try:
            _delete(hid)
            print(f"deleted {hid}", flush=True)
        except AcadError as e:
            print(f"skip delete {hid}: {e}", flush=True)

    r = acad_ops.place_polyline(cad_verts, closed=True, layer="C-CURB-EXIST",
                                reason="south lot boundary from satellite red trace")
    boundary_id = str(r.get("elementId") or "")
    print(f"boundary={boundary_id} verts={len(cad_verts)}", flush=True)

    pattern, scale = "GRASS", 30.0
    with session() as s:
        ref = entity_by_handle(s.doc, HATCH_REF)
        try:
            pattern = str(ref.PatternName or pattern)
            scale = float(ref.PatternScale or scale)
        except Exception:
            pass

    hr = acad_ops.hatch_element(boundary_id, pattern=pattern, own_element_only=True,
                                reason="south lot hatch from traced boundary")
    hatch_id = str(hr.get("elementId") or "")
    if hatch_id:
        acad_ops.change_element_layer(hatch_id, "C-CURB-EXIST", own_element_only=True,
                                      reason="match 6926C")
        with session() as s:
            h = entity_by_handle(s.doc, hatch_id)
            h.PatternScale = scale
            h.Evaluate()
    print(f"hatch={hatch_id} pattern={pattern} scale={scale}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
