"""Fix south lot pavement boundary to match satellite + existing curbs.

Deletes wrong rectangle 6B62B/6B62C. Replaces with L-shaped boundary:
  - North: existing south cutoff y=-4979 from inner curb (12245) east to access road
  - West: 69204 property bearing from (11753,-4979) to SW corner
  - South: horizontal at y=-6300
  - East: access road x=17319
  - NW jog: (11753,-4979) to (12245,-4979) closes VMF pocket / trailer lot gap
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session, entity_by_handle  # noqa: E402

HATCH_REF = "6926C"
SOUTH_Y = -6300.0

# Anchors from existing C-CURB-EXIST geometry
PROP_NW = (11753.781157245516, -4978.9367567577565)  # 69204 south tip
INNER_NW = (12245.057365225943, -4978.936756757757)   # 69207 / 6B0A0
POCKET_E = (13541.240054425056, -4978.936756757757)   # 691FB junction
E_INNER = (15396.485989211884, -4978.936756757757)     # 69208 south end
ARC_S = (16022.553131150635, -4978.936756757757)        # 691ED arc tie
E_ARC = (16794.390515109088, -4978.936756757757)      # 691EE west end
E_ROAD = (17318.73337825888, -4978.936756757757)       # 691EE east / access road


def sw_corner(y_s: float = SOUTH_Y) -> tuple[float, float]:
    """Extend 69204 property bearing south from PROP_NW."""
    x1, y1 = 12648.780205027346, 46.98007728980156
    x2, y2 = PROP_NW
    dx, dy = x2 - x1, y2 - y1
    ux, uy = dx / math.hypot(dx, dy), dy / math.hypot(dx, dy)
    t = (y_s - y2) / uy
    return x2 + t * ux, y_s


def boundary_vertices() -> list[list[float]]:
    sx, sy = sw_corner()
    return [
        [INNER_NW[0], INNER_NW[1], 0],
        [POCKET_E[0], POCKET_E[1], 0],
        [E_INNER[0], E_INNER[1], 0],
        [ARC_S[0], ARC_S[1], 0],
        [E_ARC[0], E_ARC[1], 0],
        [E_ROAD[0], E_ROAD[1], 0],
        [E_ROAD[0], SOUTH_Y, 0],
        [sx, sy, 0],
        [PROP_NW[0], PROP_NW[1], 0],
        [INNER_NW[0], INNER_NW[1], 0],
    ]


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="replace wrong boundary")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="replace wrong boundary")


def main() -> int:
    for hid in ("6B62B", "6B62C"):
        try:
            _delete(hid)
            print(f"deleted {hid}", flush=True)
        except AcadError:
            print(f"{hid} not found", flush=True)

    verts = boundary_vertices()
    r = acad_ops.place_polyline(verts, closed=True, layer="C-CURB-EXIST",
                                reason="south lot L-shaped pavement boundary")
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
                                reason="south lot hatch match 6926C")
    hatch_id = str(hr.get("elementId") or "")
    if hatch_id:
        acad_ops.change_element_layer(hatch_id, "C-CURB-EXIST",
                                      own_element_only=True, reason="match upper lot")
        with session() as s:
            h = entity_by_handle(s.doc, hatch_id)
            h.PatternScale = scale
            h.Evaluate()
    print(f"hatch={hatch_id} pattern={pattern} scale={scale}", flush=True)

    sx, _ = sw_corner()
    print(f"SW corner ({sx:.1f}, {SOUTH_Y})", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
