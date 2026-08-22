"""Fix NS (south-wall) trailers: flush to fence + add stall lines between them."""
from __future__ import annotations
import sys
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops
from acad_com import session, entity_by_handle, AcadError

BOUNDARY = "6B7F6"
NS_TEMPLATE = "6A78D"
Y_NORTH = -4978.936756757757
INSET = 8.0   # ft: south edge of trailer sits this far north of the fence
WIDER = 1.25
LONGER = 2.0
STALL_LAYER = "C-PAVEMENT MARKING"


def south_chain():
    with session() as s:
        b = s.doc.HandleToObject(BOUNDARY)
        coords = list(b.Coordinates)
        pts = [(coords[i], coords[i + 1]) for i in range(0, len(coords), 2)]
        return [pts[1], pts[2], pts[3], pts[4], pts[5], pts[6]]


def fence_y_at_x(chain, x):
    for a, b in zip(chain, chain[1:]):
        xmin, xmax = min(a[0], b[0]), max(a[0], b[0])
        if xmin - 1 <= x <= xmax + 1:
            if abs(b[0] - a[0]) < 1:
                return a[1]
            t = (x - a[0]) / (b[0] - a[0])
            return a[1] + (b[1] - a[1]) * t
    return None


def main():
    # --- gather template geometry ---
    with session() as s:
        tmpl = entity_by_handle(s.doc, NS_TEMPLATE)
        mn_raw, mx_raw = tmpl.GetBoundingBox()
        mn = (float(mn_raw[0]), float(mn_raw[1]))
        mx = (float(mx_raw[0]), float(mx_raw[1]))
        ins = [float(tmpl.InsertionPoint[0]), float(tmpl.InsertionPoint[1])]
        xs0 = float(tmpl.XScaleFactor)
        ys0 = float(tmpl.YScaleFactor)

    print(f"mn={mn} mx={mx} ins={ins} xs0={xs0} ys0={ys0}", flush=True)
    # internal block units (invariant across scale changes)
    min_y_unit = (mn[1] - ins[1]) / ys0
    max_y_unit = (mx[1] - ins[1]) / ys0
    min_x_unit = (mn[0] - ins[0]) / xs0
    max_x_unit = (mx[0] - ins[0]) / xs0

    new_xs = xs0 * WIDER
    new_ys = ys0 * LONGER
    min_y_off = min_y_unit * new_ys   # offset from ins to south edge after scale
    cx_off    = (min_x_unit + max_x_unit) / 2 * new_xs  # ins → cx after scale
    trailer_h = (max_y_unit - min_y_unit) * new_ys
    trailer_w = (max_x_unit - min_x_unit) * new_xs
    print(f"Trailer scaled: h={trailer_h:.1f} w={trailer_w:.1f}", flush=True)

    import time; time.sleep(1)
    # NS trailer cx values measured earlier (19 trailers along south fence)
    # These were sorted by cx when originally placed
    ns_cx_list = [
        12260.1, 12287.3, 12313.7, 12340.9,  # leftmost cluster (4, near SW corner)
        12430.9, 12648.0, 12865.5, 13082.6, 13302.6,  # main south row pt1
        13522.6, 13742.6, 13962.6, 14182.6, 14401.6,  # main south row pt2
        14621.6, 14841.6, 15061.6, 15281.6, 15500.0,  # main south row pt3
    ]
    ns_trailers = [{"cx": cx} for cx in ns_cx_list]
    print(f"Using {len(ns_trailers)} hardcoded NS cx values", flush=True)

    print("NS trailers already cleared. Skipping delete step.", flush=True)

    # --- place corrected NS trailers ---
    chain = south_chain()
    cx_list = sorted(t["cx"] for t in ns_trailers)

    placed = []  # list of dicts
    placed_handles = []  # (hid, cx, min_y, fence_y)
    for cx in cx_list:
        fy = fence_y_at_x(chain, cx)
        if fy is None:
            print(f"  skip cx={cx:.0f}: outside fence chain")
            continue

        desired_min_y = fy + INSET               # south edge of trailer
        ins_y = desired_min_y - min_y_off        # insertion point y
        ins_x = ins[0] + (cx - cx_off - ins[0])  # insertion point x

        dx = ins_x - ins[0]
        dy = ins_y - ins[1]

        r = acad_ops.copy_element(NS_TEMPLATE, dx, dy, own_element_only=False, reason="ns fix")
        hid = str(r.get("elementId") or "")
        placed_handles.append((hid, cx, desired_min_y, fy))

        placed.append({
            "cx": cx,
            "min_x": cx - trailer_w / 2,
            "max_x": cx + trailer_w / 2,
            "min_y": desired_min_y,
            "max_y": desired_min_y + trailer_h,
            "fence_y": fy,
        })
        print(f"  queued cx={cx:.0f}  fence={fy:.0f}  minY={desired_min_y:.0f}")

    # apply scale to all placed handles in one session
    with session() as s:
        for hid, cx, min_y, fy in placed_handles:
            ent = entity_by_handle(s.doc, hid)
            ent.XScaleFactor = new_xs
            ent.YScaleFactor = new_ys
    print(f"\nPlaced+scaled {len(placed)} NS trailers")

    # --- add vertical stall lines between adjacent trailers ---
    chain = south_chain()
    stall_count = 0
    for i in range(len(placed) - 1):
        a = placed[i]
        b = placed[i + 1]
        gap_x = (a["max_x"] + b["min_x"]) / 2
        fy = fence_y_at_x(chain, gap_x)
        if fy is None:
            continue
        y_bottom = fy + INSET
        y_top = y_bottom + trailer_h
        acad_ops.place_line(gap_x, y_bottom, gap_x, y_top,
                            layer=STALL_LAYER, reason="ns stall")
        stall_count += 1

    print(f"Placed {stall_count} stall lines")
    return 0


if __name__ == "__main__":
    import traceback
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
