
import sys, math, csv
from pathlib import Path
sys.path.insert(0, r'c:\repos\microstation-vba-project\mcp-server')
sys.path.insert(0, r'c:\repos\microstation-vba-project\scripts')
import pythoncom
import compare_median_alignment_offset as cm
from alignment_geometry import station_to_xy, segments_from_polyline, total_length, nearest_station

CHECK_STATIONS = [4450, 4550, 4650, 7450, 7550, 7650]

pythoncom.CoInitialize()
try:
    labels = cm._read_station_labels()
    events = cm._read_acsc_events()
    labels_f = [(lab, sta, x, y) for lab, sta, x, y in labels if sta >= cm.FIRST_STATION_FT]
    label_rows = [(lab, sta) for lab, sta, _x, _y in labels_f]
    output = cm._build_output_stations(label_rows, events)
    ac_segs = segments_from_polyline(cm._ordered_ac_vertices())
    physical = cm._physical_ac_segments()
    ac_len = total_length(ac_segs)
    label_xy = cm._label_xy_map(labels_f, events, ac_segs)
    mr = cm._app().ActiveDesignFile.DefaultModelReference
    refs = cm._working_attachments(mr)
    
    primary_with_info = []
    fallback_with_info = []
    
    for ref_label, att, ref_df in refs:
        xf = None
        try: xf = att.GetReferenceToMasterTransform()
        except: pass
        ref_mr = ref_df.DefaultModelReference
        levels = cm._dashed_levels_for_ref(ref_df, ref_label)
        from_basemap = 'basemap' in ref_label.lower()
        for lv in levels:
            segs = cm._scan_level_segments(ref_mr, ref_df, lv, xf)
            for p0, p1 in segs:
                if cm._dist_point_to_ac((p0[0]+p1[0])/2, (p0[1]+p1[1])/2, ac_segs) > cm.MAX_CORRIDOR_LATERAL_FT:
                    continue
                if not cm._segment_parallel_to_ac(p0, p1, ac_segs):
                    continue
                info = (p0, p1, lv, ref_label)
                if from_basemap and lv in cm.BASEMAP_EXTRA_LEVELS:
                    fallback_with_info.append(info)
                else:
                    primary_with_info.append(info)

    sta_map = {sta: (lab, pt) for lab, pt, sta in output}

    def get_hits(ax, ay, tx, ty, info_list, max_dist=15.0):
        ln = math.hypot(tx, ty)
        if ln < 1e-12: return []
        tx /= ln
        ty /= ln
        nx, ny = -ty, tx
        hits = []
        for p0, p1, lv, ref in info_list:
            dx, dy = p1[0]-p0[0], p1[1]-p0[1]
            denom = dx*dx + dy*dy
            if denom < 1e-18: continue
            u_raw = ((ax - p0[0]) * dx + (ay - p0[1]) * dy) / denom
            u = max(0.0, min(1.0, u_raw))
            qx, qy = p0[0]+u*dx, p0[1]+u*dy
            dist = math.hypot(qx-ax, qy-ay)
            if dist > max_dist: continue
            along = (qx-ax)*nx + (qy-ay)*ny
            gap = 0
            seg_len = math.sqrt(denom)
            if u_raw < 0: gap = abs(u_raw) * seg_len
            elif u_raw > 1: gap = (u_raw-1) * seg_len
            hits.append({'dist': dist, 'along': along, 'lv': lv, 'ref': ref, 'gap': gap})
        return hits

    for want in CHECK_STATIONS:
        if want not in sta_map: continue
        lab, pt = sta_map[want]
        lx, ly = label_xy.get(want, (None, None))
        ch, method = cm._chainage_for_station(want, lx, ly, ac_segs, ac_len)
        sx, sy, tx, ty = station_to_xy(ac_segs, ch)
        print(f'\n--- {lab} ch={ch:.1f} ({method}) ---')
        hits = get_hits(sx, sy, tx, ty, primary_with_info + fallback_with_info)
        hits.sort(key=lambda x: x['dist'])
        for h in hits[:12]:
            side = 'L (N)' if h['along'] > 0 else 'R (S)'
            print(f'    {side} dist={h["dist"]:.3f} gap={h["gap"]:.1f} lv={h["lv"]} ref={h["ref"]}')
finally:
    pythoncom.CoUninitialize()
