"""Copy station-guided AC path into host as one linestring for Create Alignment from Elements.

Places a single orange-ish complex path on level WZTC_AC_CIVIL_SOURCE spanning
19+50..249+00 (west) so the engineer/ORD tool can civilize it. Does not create
ruled civil geometry (CifNET required).
"""
from __future__ import annotations

import math
import re
import sys
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))
sys.path.insert(0, str(ROOT / "scripts"))

import com_cache  # noqa: F401
import pythoncom
from win32com.client import Dispatch

import compare_median_alignment_offset as cm
from alignment_geometry import PathSegment, station_to_xy, total_length

LEVEL = "WZTC_AC_CIVIL_SOURCE"
NAME_TEXT = "CL_CBC_Mainline_SOURCE"


def _unwrap(created):
    return created[0] if isinstance(created, tuple) else created


def _ensure_level(adf, name: str):
    try:
        return adf.Levels(name)
    except Exception:
        pass
    adf.AddNewLevel(name)
    try:
        adf.Levels.Rewrite()
    except Exception:
        pass
    return adf.Levels(name)


def _read_chained_ac(app):
    mr = app.ActiveModelReference
    map_att = map_df = None
    for label, att, ref_df in cm._working_attachments(mr):
        if "alg" in label.lower():
            map_att, map_df = att, ref_df
            break
    if map_att is None:
        raise RuntimeError("map_alg not attached")

    ac_level = cm._resolve_level(map_att, map_df, cm.AC_LEVEL_CANDIDATES, "AC")
    sta_level = cm._resolve_level(map_att, map_df, cm.STATION_LEVEL_CANDIDATES, "station")

    ac_cloud: list[PathSegment] = []
    o_scan = Dispatch("MicroStationDGN.ElementScanCriteria")
    o_scan.ExcludeNonGraphical()
    o_scan.ExcludeAllLevels()
    o_scan.IncludeLevel(map_df.Levels(ac_level))
    ee = map_att.Scan(o_scan)
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.Type == 12:
                groups = []
                sub = el.AsComplexStringElement.GetSubElements()
                while sub.MoveNext():
                    try:
                        groups.append(sub.Current.AsVertexList.GetVertices())
                    except Exception:
                        pass
            elif el.Type in (3, 4, 16):
                groups = [el.AsVertexList.GetVertices()]
            else:
                continue
        except Exception:
            continue
        for raw in groups:
            pts = [(float(p.X), float(p.Y)) for p in raw]
            for i in range(len(pts) - 1):
                ln = math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
                if ln < 1e-4:
                    continue
                ac_cloud.append(
                    PathSegment(
                        False,
                        pts[i][0],
                        pts[i][1],
                        0.0,
                        pts[i + 1][0],
                        pts[i + 1][1],
                        0.0,
                        ln,
                    )
                )

    labels = []
    ee = cm._level_scan(map_att, map_df, sta_level)
    while ee.MoveNext():
        el = ee.Current
        try:
            if not (el.IsTextElement or el.IsTextNodeElement):
                continue
            content = cm._text_node_lines(el)[0]
        except Exception:
            continue
        if re.match(r"^\s*(PC|PT)\b", content, re.I):
            continue
        sta = cm._parse_station_label(content)
        origin = cm._label_origin(el)
        if sta is None or origin is None:
            continue
        # Origins already host master units via attachment Scan
        labels.append((content, sta, origin[0], origin[1]))
    label_xy = {sta: (x, y) for _c, sta, x, y in labels}
    ac_segs = cm._chain_ac_by_station_labels(ac_cloud, label_xy, max_snap_ft=900.0)
    return ac_segs, labels


def main() -> int:
    pythoncom.CoInitialize()
    app = cm._app()
    adf = app.ActiveDesignFile
    mr = app.ActiveModelReference
    print("host", adf.FullName)

    ac_segs, labels = _read_chained_ac(app)
    if not ac_segs:
        print("ERROR: no AC chain")
        return 1
    ac_len = total_length(ac_segs)
    print(f"AC chain {ac_len:.1f} ft labels={len(labels)}")

    # Sample west mainline 19+50..249+00 at 50 ft
    pts = []
    ch0, ch1 = 0.0, min(ac_len, 24900.0 - 1950.0)
    ch = ch0
    while ch <= ch1 + 1e-6:
        x, y, _, _ = station_to_xy(ac_segs, ch)
        pts.append((x, y))
        if ch >= ch1:
            break
        ch = min(ch + 50.0, ch1)
    print(f"sample pts={len(pts)} from ch {ch0:.0f}..{ch1:.0f}")

    level = _ensure_level(adf, LEVEL)
    # Place connected segments (one visual chain)
    last = None
    for i in range(len(pts) - 1):
        a = app.Point3dFromXYZ(pts[i][0], pts[i][1], 0.0)
        b = app.Point3dFromXYZ(pts[i + 1][0], pts[i + 1][1], 0.0)
        seg = _unwrap(app.CreateLineElement2(None, a, b))
        seg.Level = level
        seg.Color = 4  # orange-ish index often used for AC-like
        seg.LineWeight = 4
        mr.AddElement(seg)
        seg.Rewrite()
        last = seg

    # Label at start
    te = app.CreateTextElement1(
        None,
        NAME_TEXT,
        app.Point3dFromXYZ(pts[0][0] + 20, pts[0][1] + 20, 0.0),
        app.Matrix3dIdentity(),
    )
    if isinstance(te, tuple):
        te = te[0]
    te.Level = level
    te.Color = 4
    mr.AddElement(te)
    te.Rewrite()

    # Select last segment so Create Alignment can pick nearby
    try:
        app.CadInputQueue.SendKeyin("CHOOSE NONE")
    except Exception:
        pass
    if last is not None:
        try:
            mr.SelectElement(last)
            print("selected last segment id", getattr(last.ID, "Low", last.ID))
        except Exception as e:
            print("select failed", e)

    print(f"DONE: host source on level {LEVEL} — run Create Alignment from Elements on this chain")
    print(f"start XY=({pts[0][0]:.2f},{pts[0][1]:.2f}) end XY=({pts[-1][0]:.2f},{pts[-1][1]:.2f})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
