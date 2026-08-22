"""Scale remaining south trailers 1.5x longer, one COM session each, then snap to curb."""
from __future__ import annotations

import time
import sys
from pathlib import Path

BRIDGE = Path(r"c:\repos\autocad-bridge\mcp-server")
sys.path.insert(0, str(BRIDGE))

import acad_ops  # noqa: E402
from acad_com import AcadError, session  # noqa: E402

Y_NORTH = -4978.936756757757
BOUNDARY = "6B7F6"
INSET = 8.0
TARGET_YS = 36.0  # 24 * 1.5
STALL = "C-PAVEMENT MARKING"


def bbox(ent):
    mn, mx = ent.GetBoundingBox()
    return (float(mn[0]), float(mn[1]), float(mx[0]), float(mx[1]))


def west_x_at_y(pts, y: float) -> float:
    chain = [pts[13], pts[14], pts[15], pts[0], pts[1]]
    for a, b in zip(chain, chain[1:]):
        ya, yb = sorted((a[1], b[1]))
        if ya - 1 <= y <= yb + 1:
            if abs(b[1] - a[1]) < 1e-6:
                return min(a[0], b[0])
            t = (y - a[1]) / (b[1] - a[1])
            return a[0] + (b[0] - a[0]) * t
    return pts[14][0]


def fence_y_at_x(chain, x: float):
    for a, b in zip(chain, chain[1:]):
        xmin, xmax = min(a[0], b[0]), max(a[0], b[0])
        if xmin - 1 <= x <= xmax + 1:
            if abs(b[0] - a[0]) < 1:
                return a[1]
            t = (x - a[0]) / (b[0] - a[0])
            return a[1] + (b[1] - a[1]) * t
    return None


def collect_ids():
    ids = []
    with session() as s:
        for ent in s.space:
            try:
                if str(ent.ObjectName) != "AcDbBlockReference":
                    continue
                if str(ent.Layer) != "C-VMF PARKING":
                    continue
                mn, mx = ent.GetBoundingBox()
                cy = (float(mn[1]) + float(mx[1])) / 2
                if cy > Y_NORTH:
                    continue
                ids.append(str(ent.Handle))
            except Exception:
                continue
    return ids


def scale_one(hid: str) -> None:
    with session() as s:
        ent = s.doc.HandleToObject(hid)
        ys = float(ent.YScaleFactor)
        if abs(ys - TARGET_YS) < 0.1:
            return
        ent.YScaleFactor = TARGET_YS
    time.sleep(0.08)


def main() -> int:
    ids = collect_ids()
    print(f"trailers={len(ids)}", flush=True)
    for i, hid in enumerate(ids):
        scale_one(hid)
        print(f"  scaled {i+1}/{len(ids)} {hid}", flush=True)

    with session() as s:
        b = s.doc.HandleToObject(BOUNDARY)
        coords = list(b.Coordinates)
        pts = [(float(coords[i]), float(coords[i + 1])) for i in range(0, len(coords), 2)]
        south = [pts[1], pts[2], pts[3], pts[4], pts[5], pts[6]]

        recs = []
        for hid in ids:
            ent = s.doc.HandleToObject(hid)
            minx, miny, maxx, maxy = bbox(ent)
            recs.append({
                "h": hid, "minx": minx, "miny": miny, "maxx": maxx, "maxy": maxy,
                "cx": (minx + maxx) / 2, "cy": (miny + maxy) / 2,
                "w": maxx - minx, "ht": maxy - miny,
            })

    west = [t for t in recs if t["ht"] < t["w"] and t["cx"] < 12350]
    center = [t for t in recs if t["ht"] < t["w"] and t["cx"] >= 12350]
    ns = [t for t in recs if t["ht"] > t["w"]]
    print(f"west={len(west)} center={len(center)} ns={len(ns)}", flush=True)

    # snap west
    with session() as s:
        for t in west:
            ent = s.doc.HandleToObject(t["h"])
            minx, miny, maxx, maxy = bbox(ent)
            cy = (miny + maxy) / 2
            wx = west_x_at_y(pts, cy)
            dx = (wx + INSET) - minx
            if abs(dx) > 0.05:
                ent.Move((0, 0, 0), (dx, 0, 0))
            minx, miny, maxx, maxy = bbox(ent)
            t.update(minx=minx, miny=miny, maxx=maxx, maxy=maxy,
                     cx=(minx + maxx) / 2, cy=(miny + maxy) / 2)
            print(f"  west snap {t['h']} minx={minx:.0f} curb={wx:.0f}", flush=True)

    if center:
        cxs = sorted({round(t["cx"], 0) for t in center})
        mid_cx = sum(cxs) / 2
        with session() as s:
            for t in center:
                ent = s.doc.HandleToObject(t["h"])
                minx, miny, maxx, maxy = bbox(ent)
                if t["cx"] < mid_cx:
                    dx = mid_cx - maxx
                else:
                    dx = mid_cx - minx
                if abs(dx) > 0.05:
                    ent.Move((0, 0, 0), (dx, 0, 0))
                minx, miny, maxx, maxy = bbox(ent)
                t.update(minx=minx, miny=miny, maxx=maxx, maxy=maxy,
                         cx=(minx + maxx) / 2, cy=(miny + maxy) / 2)
                print(f"  center snap {t['h']} {minx:.0f}->{maxx:.0f}", flush=True)
    else:
        mid_cx = 0.0

    with session() as s:
        for t in ns:
            ent = s.doc.HandleToObject(t["h"])
            minx, miny, maxx, maxy = bbox(ent)
            cx = (minx + maxx) / 2
            fy = fence_y_at_x(south, cx)
            if fy is None:
                continue
            dy = (fy + INSET) - miny
            if abs(dy) > 0.05:
                ent.Move((0, 0, 0), (0, dy, 0))
            minx, miny, maxx, maxy = bbox(ent)
            t.update(minx=minx, miny=miny, maxx=maxx, maxy=maxy,
                     cx=(minx + maxx) / 2, cy=(miny + maxy) / 2)
            print(f"  ns snap {t['h']} miny={miny:.0f} fence={fy:.0f}", flush=True)

    # stall lines
    west_sorted = sorted(west, key=lambda t: -t["cy"])
    for i in range(len(west_sorted) - 1):
        a, b = west_sorted[i], west_sorted[i + 1]
        y = (a["cy"] + b["cy"]) / 2
        acad_ops.place_line(min(a["minx"], b["minx"]), y, max(a["maxx"], b["maxx"]), y,
                            layer=STALL, reason="west stall")
        time.sleep(0.05)

    if center:
        left = sorted([t for t in center if t["cx"] < mid_cx], key=lambda t: -t["cy"])
        right = sorted([t for t in center if t["cx"] >= mid_cx], key=lambda t: -t["cy"])
        if left and right:
            x1 = min(t["minx"] for t in left)
            x2 = max(t["maxx"] for t in right)
            for i in range(len(left) - 1):
                y = (left[i]["cy"] + left[i + 1]["cy"]) / 2
                acad_ops.place_line(x1, y, x2, y, layer=STALL, reason="center stall")
                time.sleep(0.05)
            mid_x = (max(t["maxx"] for t in left) + min(t["minx"] for t in right)) / 2
            acad_ops.place_line(mid_x, min(t["miny"] for t in center),
                                mid_x, max(t["maxy"] for t in center),
                                layer=STALL, reason="center mid")

    ns_sorted = sorted(ns, key=lambda t: t["cx"])
    for i in range(len(ns_sorted) - 1):
        a, b = ns_sorted[i], ns_sorted[i + 1]
        gap_x = (a["maxx"] + b["minx"]) / 2
        acad_ops.place_line(gap_x, min(a["miny"], b["miny"]), gap_x, max(a["maxy"], b["maxy"]),
                            layer=STALL, reason="ns stall")
        time.sleep(0.05)

    with session() as s:
        s.doc.Regen(1)
    print("done", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
