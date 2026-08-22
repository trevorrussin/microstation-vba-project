"""Rebuild center fleet as 2 columns of 5 (E-W, back-to-back), not 2 rows of 5."""
from __future__ import annotations

import sys
import time
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session, entity_by_handle  # noqa: E402

Y_NORTH = -4978.936756757757
EW_TEMPLATE = "6A741"
# Template bbox center (copy shifts center-to-center before scale)
EW_X, EW_Y = 14155.02424640845, -4585.566667744555
WIDER = 1.25
LONGER = 2.0
STALL_LAYER = "C-PAVEMENT MARKING"


def _delete(hid: str) -> None:
    try:
        acad_ops.delete_element(hid, own_element_only=True, reason="center cols")
    except AcadError:
        acad_ops.delete_element(hid, own_element_only=False, reason="center cols")


def main() -> int:
    # Scaled EW size (matches live west trailers: 472 x 115)
    ew_len = 236.21452930715532 * LONGER
    ew_dep = 92.34840715461269 * WIDER
    row_pitch = 120.0 * WIDER  # ~150 ft

    # Collect center EW trailers + the long center stall line
    del_ids: list[str] = []
    west_max_x = 0.0
    with session() as s:
        for ent in s.space:
            try:
                oname = str(ent.ObjectName)
                layer = str(ent.Layer)
                mn_r, mx_r = ent.GetBoundingBox()
                mn = (float(mn_r[0]), float(mn_r[1]))
                mx = (float(mx_r[0]), float(mx_r[1]))
                cx = (mn[0] + mx[0]) / 2
                cy = (mn[1] + mx[1]) / 2
            except Exception:
                continue
            if cy > Y_NORTH:
                continue
            hid = str(ent.Handle)
            if oname == "AcDbBlockReference" and layer == "C-VMF PARKING":
                w, h = mx[0] - mn[0], mx[1] - mn[1]
                if w > h and cx > 12300:
                    del_ids.append(hid)
                elif w > h and cx < 12300:
                    west_max_x = max(west_max_x, mx[0])
            elif oname == "AcDbLine" and layer == STALL_LAYER:
                # long horizontal center separator (~2500+ ft wide)
                if abs(mx[1] - mn[1]) < 5 and (mx[0] - mn[0]) > 1000:
                    del_ids.append(hid)

    print(f"Deleting {len(del_ids)} center elems; west_max_x={west_max_x:.0f}", flush=True)
    for hid in del_ids:
        _delete(hid)
    time.sleep(0.5)

    # Two columns of 5, back-to-back (abutting ends), east of west fleet + aisle
    aisle = 280.0
    col1_cx = west_max_x + aisle + ew_len / 2
    col2_cx = col1_cx + ew_len  # back-to-back
    y0 = Y_NORTH - 100.0 - ew_dep / 2  # align with west top row

    new_xs = 12.0 * WIDER
    new_ys = 12.0 * LONGER
    handles: list[str] = []
    centers: list[tuple[float, float]] = []

    for col_cx in (col1_cx, col2_cx):
        for i in range(5):
            cy = y0 - i * row_pitch
            r = acad_ops.copy_element(
                EW_TEMPLATE, col_cx - EW_X, cy - EW_Y,
                own_element_only=False, reason="center col",
            )
            hid = str(r.get("elementId") or "")
            handles.append(hid)
            centers.append((col_cx, cy))
            print(f"  queued ({col_cx:.0f}, {cy:.0f}) -> {hid}", flush=True)

    time.sleep(0.5)
    with session() as s:
        for hid in handles:
            ent = entity_by_handle(s.doc, hid)
            ent.XScaleFactor = new_xs
            ent.YScaleFactor = new_ys
    print(f"Scaled {len(handles)} center trailers", flush=True)

    # Stall lines between consecutive rows spanning both columns
    x1 = col1_cx - ew_len / 2
    x2 = col2_cx + ew_len / 2
    for i in range(4):
        y = (centers[i][1] + centers[i + 1][1]) / 2
        acad_ops.place_line(x1, y, x2, y, layer=STALL_LAYER, reason="center stall")
        print(f"  stall y={y:.0f}", flush=True)

    # Mid divider between the two back-to-back columns
    mid_x = (col1_cx + col2_cx) / 2
    y_top = centers[0][1] + ew_dep / 2
    y_bot = centers[4][1] - ew_dep / 2
    acad_ops.place_line(mid_x, y_bot, mid_x, y_top, layer=STALL_LAYER, reason="center aisle")
    print(f"Done: 2 cols x 5 @ cx={col1_cx:.0f}/{col2_cx:.0f}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
