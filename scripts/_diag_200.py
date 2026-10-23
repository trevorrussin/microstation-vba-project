"""Diagnose 200+ station offset measurement points and nearest dashed hits."""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))
sys.path.insert(0, str(ROOT / "scripts"))

import pythoncom
import compare_median_alignment_offset as cm
from alignment_geometry import nearest_station, segments_from_polyline, station_to_xy, total_length

CHECK = [
    20050, 20450, 20850, 21050, 21250, 22050, 22650, 24050, 24150, 24650, 24900,
]


def hits_at(ax, ay, tx, ty, segs, max_dist=8.0):
    ln = math.hypot(tx, ty)
    tx, ty = tx / ln, ty / ln
    nx, ny = -ty, tx
    left, right = [], []
    for p0, p1 in segs:
        if not cm._segment_parallel_to_tangent(p0, p1, tx, ty):
            continue
        dx, dy = p1[0] - p0[0], p1[1] - p0[1]
        denom = dx * dx + dy * dy
        if denom < 1e-18:
            qx, qy = p0
        else:
            u = max(0.0, min(1.0, ((ax - p0[0]) * dx + (ay - p0[1]) * dy) / denom))
            qx, qy = p0[0] + u * dx, p0[1] + u * dy
        along = (qx - ax) * nx + (qy - ay) * ny
        d = math.hypot(qx - ax, qy - ay)
        if d < 0.05 or d > max_dist:
            continue
        if along > 0.05:
            left.append(d)
        elif along < -0.05:
            right.append(d)
    left.sort()
    right.sort()
    return left[:5], right[:5]


pythoncom.CoInitialize()
try:
    labels = cm._read_station_labels()
    events = cm._read_acsc_events()
finally:
    pythoncom.CoUninitialize()

labels_f = [(lab, sta, x, y) for lab, sta, x, y in labels if sta >= cm.FIRST_STATION_FT]
label_rows = [(lab, sta) for lab, sta, _x, _y in labels_f]
output = cm._build_output_stations(label_rows, events)
ac_segs = segments_from_polyline(cm._ordered_ac_vertices())
physical = cm._physical_ac_segments()
ac_len = total_length(ac_segs)
label_xy = cm._label_xy_map(labels_f, events, ac_segs)

pythoncom.CoInitialize()
try:
    mr = cm._app().ActiveDesignFile.DefaultModelReference
    refs = cm._working_attachments(mr)
finally:
    pythoncom.CoUninitialize()

primary_all = []
fallback_all = []
for ref_label, att, ref_df in refs:
    pythoncom.CoInitialize()
    try:
        primary, fallback, _names = cm._collect_dashed_segments(att, ref_df, ac_segs, ref_label)
    finally:
        pythoncom.CoUninitialize()
    primary_all.extend(primary)
    fallback_all.extend(fallback)

print(f"primary={len(primary_all)} fallback={len(fallback_all)}")
sta_map = {sta: (lab, pt) for lab, pt, sta in output}

for want in CHECK:
    lab, pt = sta_map[want]
    lx, ly = label_xy.get(want, (None, None))
    ch, method = cm._chainage_for_station(want, lx, ly, ac_segs, ac_len)
    formula = want - cm.STATION_CHAINAGE_OFFSET
    sx, sy, tx, ty = station_to_xy(ac_segs, ch)
    if lx is not None:
        ch_lab, perp = nearest_station(ac_segs, lx, ly)
    else:
        ch_lab, perp = None, None
    ax, ay, ptx, pty = cm._nearest_on_physical_ac(sx, sy, physical)
    phys_dist = math.hypot(ax - sx, ay - sy)

    Lp, Rp = hits_at(sx, sy, tx, ty, primary_all)
    Lf, Rf = hits_at(sx, sy, tx, ty, fallback_all)
    Lpp, Rpp = hits_at(ax, ay, ptx, pty, primary_all)
    print(
        f"\n{lab} ch={ch:.1f} ({method}) formula2732={formula:.1f} "
        f"label_ch={ch_lab} perp={perp:.1f} phys_snap={phys_dist:.1f}ft"
    )
    print(f"  at station_xy: primary L={Lp[:3]} R={Rp[:3]}  fallback L={Lf[:3]} R={Rf[:3]}")
    print(f"  at physical:   primary L={Lpp[:3]} R={Rpp[:3]}")
    n, s = cm._offsets_at_chainage(ch, ac_segs, physical, primary_all, fallback_all)
    print(f"  reported N={n} S={s}")
