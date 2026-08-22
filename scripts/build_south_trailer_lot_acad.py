"""South trailer lot — satellite-accurate sparse clusters with drive aisles.

680 US-130 Trenton VMF. Red zone on annotated screenshot:
  - West edge column (E-W trailers, backs to west)
  - Central back-to-back island (2 short E-W rows) — east of west column
  - South fence: TWO separate N-S blocks with gap = drive aisle
  - NE corner: short row near entry
  - CENTER OF LOT = OPEN DRIVE AISLE (no trailers)

E-W template: 6A741 (rotation 270).  N-S template: 6A78D (rotation ~180).
"""
from __future__ import annotations

import sys
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session  # noqa: E402

EW_TEMPLATE = "6A741"
NS_TEMPLATE = "6A78D"
EW_X = 14155.02424640845
EW_Y = -4585.566667744555
NS_X = 13732.353514847744
NS_Y = -4071.8642288933092

SOUTH_CUTOFF = -4979.0
COL = 240
ROW = 120

# --- Satellite-matched clusters (centers).  DO NOT fill drive aisles. ---

def _line_x(x0: float, n: int) -> list[float]:
    return [x0 + i * COL for i in range(n)]


CLUSTERS: dict[str, list[tuple[float, float, str]]] = {
    # West edge: ~9 E-W trailers — TOP of red zone only (not full lot height)
    "west_edge": [
        (11900, y, "ew") for y in [-5080, -5200, -5320, -5440, -5560, -5680, -5800, -5920, -6040]
    ],
    # Central island — 2 back-to-back E-W rows (5 stalls each), east of west column
    "center_row_south": [(x, -5100, "ew") for x in _line_x(12800, 5)],
    "center_row_north": [(x, -5220, "ew") for x in _line_x(12800, 5)],
    # NE corner near entry — short row
    "ne_entry": [(x, -5080, "ew") for x in _line_x(15000, 4)],
    # South fence — TWO N-S blocks separated by ~1200 ft drive gap (satellite gap)
    "south_west_block": [(x, -6220, "ns") for x in _line_x(11900, 13)],
    "south_east_block": [(x, -6220, "ns") for x in _line_x(15600, 8)],
}

# Stall lines only at cluster edges — not across drive aisles
STALL_LINES: list[tuple[float, float, float]] = [
    # west edge row separators
    (11820, 12180, -5140),
    (11820, 12180, -5260),
    (11820, 12180, -5380),
    (11820, 12180, -5500),
    (11820, 12180, -5620),
    (11820, 12180, -5740),
    (11820, 12180, -5860),
    (11820, 12180, -5980),
    (11820, 12180, -6100),
    # center island
    (12720, 14080, -5160),
    (12720, 14080, -5280),
    # south blocks
    (11820, 14380, -6160),
    (15520, 17500, -6160),
]


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="south lot rebuild")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="south lot rebuild")


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
            if cy > SOUTH_CUTOFF:
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


def place_trailer(tx: float, ty: float, orient: str, label: str) -> str:
    if orient == "ns":
        r = acad_ops.copy_element(NS_TEMPLATE, tx - NS_X, ty - NS_Y,
                                  own_element_only=False, reason=label)
    else:
        r = acad_ops.copy_element(EW_TEMPLATE, tx - EW_X, ty - EW_Y,
                                  own_element_only=False, reason=label)
    return str(r.get("elementId") or "")


def main() -> int:
    total = sum(len(v) for v in CLUSTERS.values())
    print("=== Clear ===", flush=True)
    tr, st = clear_south_lot()
    print(f"  removed {tr} trailers, {st} stall lines", flush=True)

    print(f"=== Place {total} trailers (satellite clusters) ===", flush=True)
    n = 0
    for name, pts in CLUSTERS.items():
        print(f"  {name}: {len(pts)}", flush=True)
        for i, (tx, ty, orient) in enumerate(pts):
            place_trailer(tx, ty, orient, f"{name}[{i}]")
            n += 1

    print("=== Stall lines ===", flush=True)
    for x1, x2, y in STALL_LINES:
        acad_ops.place_line(x1, y, x2, y, layer="C-PAVEMENT MARKING",
                            reason=f"cluster stall y={y}")
    print(f"Done: {n} trailers. Center drive aisle left OPEN.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
