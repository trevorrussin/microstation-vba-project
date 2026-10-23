
import sys, math
from pathlib import Path
sys.path.insert(0, r'c:\repos\microstation-vba-project\mcp-server')
sys.path.insert(0, r'c:\repos\microstation-vba-project\scripts')
import pythoncom
import compare_median_alignment_offset as cm
from alignment_geometry import station_to_xy, segments_from_polyline, total_length
from win32com.client import Dispatch

WANT = 7550

pythoncom.CoInitialize()
try:
    ac_segs = segments_from_polyline(cm._ordered_ac_vertices())
    ac_len = total_length(ac_segs)
    labels = cm._read_station_labels()
    events = cm._read_acsc_events()
    labels_f = [(lab, sta, x, y) for lab, sta, x, y in labels if sta >= cm.FIRST_STATION_FT]
    label_xy = cm._label_xy_map(labels_f, events, ac_segs)
    lx, ly = label_xy.get(WANT, (None, None))
    ch, method = cm._chainage_for_station(WANT, lx, ly, ac_segs, ac_len)
    sx, sy, tx, ty = station_to_xy(ac_segs, ch)
    
    ln = math.hypot(tx, ty)
    tx_unit, ty_unit = tx/ln, ty/ln
    nx, ny = -ty_unit, tx_unit
    
    app = cm._app()
    mr = app.ActiveDesignFile.DefaultModelReference
    
    def scan_level(model, level_name, x, y, xf=None, range_ft=30.0):
        found = []
        try:
            o = Dispatch('MicroStationDGN.ElementScanCriteria')
            o.ExcludeNonGraphical()
            # find level
            lv = None
            try: lv = model.DesignFile.Levels(level_name)
            except: 
                try: lv = app.ActiveDesignFile.Levels(level_name)
                except: pass
            if lv: o.IncludeLevel(lv)
            
            ee = model.Scan(o)
            while ee.MoveNext():
                el = ee.Current
                if el.Type not in [2, 3, 4, 12, 16]: continue
                try:
                    pts = el.AsVertexList.GetVertices()
                    for i in range(len(pts)-1):
                        p0 = (pts[i].X, pts[i].Y)
                        p1 = (pts[i+1].X, pts[i+1].Y)
                        if xf:
                            p0 = cm._transform_point(pts[i], xf)
                            p1 = cm._transform_point(pts[i+1], xf)
                        
                        dx, dy = p1[0]-p0[0], p1[1]-p0[1]
                        d2 = dx*dx + dy*dy
                        if d2 < 1e-12: u = 0
                        else: u = max(0, min(1, ((x-p0[0])*dx + (y-p0[1])*dy)/d2))
                        px, py = p0[0]+u*dx, p0[1]+u*dy
                        d = math.hypot(px-x, py-y)
                        if d < range_ft:
                            along = (px-x)*nx + (py-y)*ny
                            seg_ln = math.sqrt(d2)
                            dot = abs((dx/seg_ln)*tx_unit + (dy/seg_ln)*ty_unit)
                            found.append({'dist': d, 'along': along, 'dot': dot, 'id': el.ID.Low, 'lv': el.Level.Name if el.Level else '?'})
                except: pass
        except Exception as e:
            print(f'Error scanning {level_name}: {e}')
        return found

    print(f'Target station {WANT} at ({sx:.1f}, {sy:.1f})')
    
    levels_to_check = ['PET', 'PELB', 'BELB', 'PE', 'RGCB', 'RGB', 'RG', 'Level 2', 'LLP', 'LST', 'LTS']
    
    hits = []
    for lv in levels_to_check:
        hits += scan_level(mr, lv, sx, sy)
        
    for att in mr.Attachments:
        xf = None
        try: xf = att.GetReferenceToMasterTransform()
        except: pass
        for lv in levels_to_check:
            hits += scan_level(att.DesignFile.DefaultModelReference, lv, sx, sy, xf)
            
    hits.sort(key=lambda x: x['dist'])
    for h in hits[:40]:
        side = 'L (N)' if h['along'] > 0 else 'R (S)'
        print(f'  {side} dist={h["dist"]:.3f} lv={h["lv"]} dot={h["dot"]:.3f} id={h["id"]}')

finally:
    pythoncom.CoUninitialize()
