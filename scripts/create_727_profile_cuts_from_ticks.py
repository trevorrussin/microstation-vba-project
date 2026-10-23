"""Place live 1000-ft AC profile-cut guides from map_alg AC + AC_C ticks.

Uses the orange AC control line and station tick marks in the attached
``727.21.101_map_alg_091026.dgn``. Writes cut anchors CSV and places, in the
active host model:

- one polyline per 1000-ft cut along AC (level WZTC_EG_PROFILE_CUTS)
- a short tick mark + station text at each cut start/end

Does NOT invent ORD ruled profiles (Quick Profile From Surface has no known
headless keyin). These guides are the locate targets for that ribbon tool /
Named Boundary Civil Profile station limits.
"""
from __future__ import annotations

import csv
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
from alignment_geometry import PathSegment, nearest_station, station_to_xy, total_length

OUT = ROOT / "outputs" / "727-profile-cuts"
OUT.mkdir(parents=True, exist_ok=True)
LEVEL = "WZTC_EG_PROFILE_CUTS"
CUT_FT = 1000.0
WEST_END = 24900.0
EAST_START = 29634.0
EAST_END = 29975.0


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


def _point3d(app, x: float, y: float, z: float = 0.0):
    p = app.Point3dFromXYZ(x, y, z)
    return p


def _read_map_alg(app):
    mr = app.ActiveModelReference
    map_att = map_df = None
    for label, att, ref_df in cm._working_attachments(mr):
        if "alg" in label.lower():
            map_att, map_df = att, ref_df
            print("map_alg", label)
            break
    if map_att is None:
        raise RuntimeError("map_alg attachment not found")

    ac_level = cm._resolve_level(map_att, map_df, cm.AC_LEVEL_CANDIDATES, "AC")
    sta_level = cm._resolve_level(map_att, map_df, cm.STATION_LEVEL_CANDIDATES, "station")

    # AC cloud — Scan returns host master units already (do not ×10).
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

    # Tick labels — Origins already host master units.
    labels: list[tuple[str, float, float, float]] = []
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
        labels.append((content, sta, origin[0], origin[1]))
    labels.sort(key=lambda r: r[1])

    label_xy = {sta: (x, y) for _c, sta, x, y in labels}
    # Ticks can sit >250 ft from sparse AC graphics near forks — allow 900 ft.
    ac_segs = cm._chain_ac_by_station_labels(ac_cloud, label_xy, max_snap_ft=900.0)
    return ac_cloud, ac_segs, labels, label_xy


def _cut_stations() -> list[tuple[str, float, float, str]]:
    """(cut_id, from_sta, to_sta, chain)."""
    cuts: list[tuple[str, float, float, str]] = []
    s = 1950.0
    i = 1
    while s < WEST_END - 1e-6:
        e = min(s + CUT_FT, WEST_END)
        cuts.append((f"W{i:02d}", s, e, "west"))
        s = e
        i += 1
    cuts.append(("E01", EAST_START, EAST_END, "east"))
    return cuts


def _sample_polyline(ac_segs: list[PathSegment], sta0: float, sta1: float, step: float = 25.0):
    """Sample host XY along AC between station ft (chainage = sta - 1950)."""
    ch0 = max(0.0, sta0 - 1950.0)
    ch1 = max(ch0, sta1 - 1950.0)
    ac_len = total_length(ac_segs)
    ch0 = min(ch0, ac_len)
    ch1 = min(ch1, ac_len)
    pts = []
    ch = ch0
    while ch <= ch1 + 1e-6:
        x, y, _tx, _ty = station_to_xy(ac_segs, ch)
        pts.append((x, y))
        if ch >= ch1:
            break
        ch = min(ch + step, ch1)
    return pts


def _place_linestring(app, mr, level, pts: list[tuple[float, float]], color: int, weight: int):
    if len(pts) < 2:
        return None
    # CreateLineElement1 needs a SAFEARRAY of Point3d — build via successive
    # CreateLineElement2 segments (reliable on this COM install).
    def _unwrap(created):
        return created[0] if isinstance(created, tuple) else created

    last = None
    for i in range(len(pts) - 1):
        a = _point3d(app, pts[i][0], pts[i][1])
        b = _point3d(app, pts[i + 1][0], pts[i + 1][1])
        seg = _unwrap(app.CreateLineElement2(None, a, b))
        seg.Level = level
        seg.Color = color
        seg.LineWeight = weight
        mr.AddElement(seg)
        seg.Rewrite()
        last = seg
    return last


def _place_tick_and_label(app, mr, level, x: float, y: float, text: str, color: int):
    def _unwrap(created):
        return created[0] if isinstance(created, tuple) else created

    # 20-ft cross tick
    d = 10.0
    h0 = _point3d(app, x - d, y)
    h1 = _point3d(app, x + d, y)
    v0 = _point3d(app, x, y - d)
    v1 = _point3d(app, x, y + d)
    for a, b in ((h0, h1), (v0, v1)):
        el = _unwrap(app.CreateLineElement2(None, a, b))
        el.Level = level
        el.Color = color
        el.LineWeight = 2
        mr.AddElement(el)
        el.Rewrite()
    rot = app.Matrix3dIdentity()
    te = app.CreateTextElement1(None, text, _point3d(app, x + 12, y + 12), rot)
    if isinstance(te, tuple):
        te = te[0]
    te.Level = level
    te.Color = color
    try:
        te.TextStyle.Height = 8.0
        te.TextStyle.Width = 8.0
    except Exception:
        pass
    mr.AddElement(te)
    te.Rewrite()


def main() -> int:
    pythoncom.CoInitialize()
    app = cm._app()
    adf = app.ActiveDesignFile
    mr = app.ActiveModelReference
    print("host", adf.FullName)

    _ac_cloud, ac_segs, labels, label_xy = _read_map_alg(app)
    if not ac_segs:
        print("ERROR: failed to chain AC along ticks")
        return 1
    print(f"AC chained length={total_length(ac_segs):.1f} ft  labels={len(labels)}")

    cuts = _cut_stations()
    csv_path = OUT / "cut_anchors_from_ticks.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "cut_id",
                "from_sta",
                "to_sta",
                "from_tick",
                "to_tick",
                "from_x",
                "from_y",
                "to_x",
                "to_y",
                "chain",
            ]
        )
        for cid, a, b, chain in cuts:
            # nearest tick labels
            fa = min(labels, key=lambda r: abs(r[1] - a))
            fb = min(labels, key=lambda r: abs(r[1] - b))
            # AC snap XY at chainage
            xa, ya, _, _ = station_to_xy(ac_segs, max(0.0, a - 1950.0))
            xb, yb, _, _ = station_to_xy(ac_segs, max(0.0, b - 1950.0))
            w.writerow(
                [
                    cid,
                    a,
                    b,
                    fa[0],
                    fb[0],
                    f"{xa:.4f}",
                    f"{ya:.4f}",
                    f"{xb:.4f}",
                    f"{yb:.4f}",
                    chain,
                ]
            )
    print("wrote", csv_path)

    level = _ensure_level(adf, LEVEL)
    app.ActiveSettings.Level = level
    placed = 0
    for cid, a, b, chain in cuts:
        pts = _sample_polyline(ac_segs, a, b, step=25.0)
        if len(pts) < 2:
            print(f"skip {cid}: no path")
            continue
        color = 4 if chain == "west" else 5  # blue / purple-ish
        _place_linestring(app, mr, level, pts, color=color, weight=3)
        _place_tick_and_label(
            app, mr, level, pts[0][0], pts[0][1], f"{cid} {cm._format_station(a)}", color
        )
        _place_tick_and_label(
            app, mr, level, pts[-1][0], pts[-1][1], cm._format_station(b), color
        )
        placed += 1
        print(f"placed {cid} {cm._format_station(a)}..{cm._format_station(b)} pts={len(pts)}")

    print(f"DONE placed={placed}/{len(cuts)} on level {LEVEL}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
