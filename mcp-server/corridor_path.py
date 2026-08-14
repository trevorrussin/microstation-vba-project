"""Polyline helpers for corridor pick / work-bay snap (no MicroStation)."""
from __future__ import annotations

import math
from typing import Optional


MAX_PATH_VERTS = 80


def _xy(p) -> tuple[float, float]:
    return float(p[0]), float(p[1])


def polyline_length(pts: list) -> float:
    if not pts or len(pts) < 2:
        return 0.0
    total = 0.0
    for i in range(1, len(pts)):
        x0, y0 = _xy(pts[i - 1])
        x1, y1 = _xy(pts[i])
        total += math.hypot(x1 - x0, y1 - y0)
    return total


def downsample_polyline(pts: list, max_n: int = MAX_PATH_VERTS) -> list[list[float]]:
    clean = []
    for p in pts or []:
        x, y = _xy(p)
        z = float(p[2]) if len(p) > 2 else 0.0
        if clean:
            px, py = clean[-1][0], clean[-1][1]
            if math.hypot(x - px, y - py) < 1e-6:
                continue
        clean.append([x, y, z])
    if len(clean) <= max_n:
        return clean
    n = len(clean)
    out = []
    for i in range(max_n):
        idx = int(round(i * (n - 1) / (max_n - 1)))
        if not out or out[-1] is not clean[idx]:
            out.append(clean[idx])
    if out[-1] is not clean[-1]:
        out.append(clean[-1])
    return [[p[0], p[1], p[2]] for p in out]


def point_at_station(pts: list, station: float) -> list[float]:
    if not pts:
        raise ValueError("empty path")
    if len(pts) == 1:
        x, y = _xy(pts[0])
        return [x, y, 0.0]
    if station <= 0:
        x, y = _xy(pts[0])
        return [x, y, 0.0]
    remaining = station
    for i in range(1, len(pts)):
        x0, y0 = _xy(pts[i - 1])
        x1, y1 = _xy(pts[i])
        seg = math.hypot(x1 - x0, y1 - y0)
        if seg < 1e-12:
            continue
        if remaining <= seg:
            t = remaining / seg
            return [x0 + t * (x1 - x0), y0 + t * (y1 - y0), 0.0]
        remaining -= seg
    x, y = _xy(pts[-1])
    return [x, y, 0.0]


def nearest_station(pts: list, x: float, y: float) -> dict:
    """Closest point on the polyline. Returns stationFt, x, y, distFt."""
    if not pts:
        raise ValueError("empty path")
    best = {"stationFt": 0.0, "x": _xy(pts[0])[0], "y": _xy(pts[0])[1],
            "distFt": math.hypot(x - _xy(pts[0])[0], y - _xy(pts[0])[1])}
    sta_at = 0.0
    for i in range(1, len(pts)):
        x0, y0 = _xy(pts[i - 1])
        x1, y1 = _xy(pts[i])
        dx, dy = x1 - x0, y1 - y0
        seg = math.hypot(dx, dy)
        if seg < 1e-12:
            continue
        t = ((x - x0) * dx + (y - y0) * dy) / (seg * seg)
        t = 0.0 if t < 0 else (1.0 if t > 1 else t)
        px, py = x0 + t * dx, y0 + t * dy
        d = math.hypot(x - px, y - py)
        if d < best["distFt"]:
            best = {
                "stationFt": sta_at + t * seg,
                "x": px, "y": py, "distFt": d,
            }
        sta_at += seg
    return best


def sample_span(pts: list, sta0: float, sta1: float, step_ft: float = 25.0) -> list[list[float]]:
    if step_ft <= 0:
        raise ValueError("step_ft must be > 0")
    span = abs(sta1 - sta0)
    if span < 1e-9:
        return [point_at_station(pts, sta0)]
    n = max(1, int(math.ceil(span / step_ft)))
    out = []
    for i in range(n + 1):
        sta = sta0 + (sta1 - sta0) * (i / n)
        p = point_at_station(pts, sta)
        if out and math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) < 1e-6:
            continue
        out.append(p)
    return out


def offset_polyline(pts: list, dist_ft: float) -> list[list[float]]:
    """Offset to the right of travel (vertex order) by dist_ft."""
    if dist_ft == 0 or len(pts) < 2:
        return downsample_polyline(pts)
    out = []
    n = len(pts)
    for i in range(n):
        x, y = _xy(pts[i])
        if i == 0:
            x1, y1 = _xy(pts[1])
            tx, ty = x1 - x, y1 - y
        elif i == n - 1:
            x0, y0 = _xy(pts[i - 1])
            tx, ty = x - x0, y - y0
        else:
            x0, y0 = _xy(pts[i - 1])
            x1, y1 = _xy(pts[i + 1])
            tx, ty = x1 - x0, y1 - y0
        mag = math.hypot(tx, ty)
        if mag < 1e-12:
            nx, ny = 0.0, 0.0
        else:
            nx, ny = ty / mag, -tx / mag  # right of travel
        out.append([x + nx * dist_ft, y + ny * dist_ft, 0.0])
    return downsample_polyline(out)


def reverse_polyline(pts: list) -> list[list[float]]:
    return [[float(p[0]), float(p[1]), float(p[2]) if len(p) > 2 else 0.0]
            for p in reversed(pts)]


def endpoint_label(pt: list) -> str:
    x, y = _xy(pt)
    return f"({x:.0f}, {y:.0f})"


def travel_choice_options(pts: list) -> list[dict]:
    a, b = pts[0], pts[-1]
    return [
        {
            "label": f"Toward {endpoint_label(b)}",
            "description": f"Travel from {endpoint_label(a)} to {endpoint_label(b)} (path vertex order)",
            "value": "as_drawn",
        },
        {
            "label": f"Toward {endpoint_label(a)}",
            "description": f"Travel from {endpoint_label(b)} to {endpoint_label(a)} (reverse path)",
            "value": "reverse",
        },
    ]


def sheet_approach_ft(spec: dict, resolved: dict) -> dict:
    """Upstream + downstream scheme length from the station walk (no work bay)."""
    import sheet_spec
    walk = sheet_spec.station_walk(spec, resolved)
    up = [w["stationFt"] for w in walk if int(w.get("alignIdx") or 0) == 1
          and w.get("rowNum") is not None]
    dn = [w["stationFt"] for w in walk if int(w.get("alignIdx") or 0) == 2
          and w.get("rowNum") is not None]
    upstream = max(up) if up else 0.0
    downstream = max(dn) if dn else 0.0
    return {
        "upstreamFt": upstream,
        "downstreamFt": downstream,
        "bothSidesFt": upstream + downstream,
    }


def length_check(path_len: float, approach: dict, work_len: Optional[float] = None) -> dict:
    need = float(approach["bothSidesFt"])
    if work_len is not None:
        need += float(work_len)
    ok = path_len + 1e-6 >= need
    short = 0.0 if ok else need - path_len
    return {
        "ok": ok,
        "pathLengthFt": path_len,
        "neededFt": need,
        "shortfallFt": short,
        "upstreamFt": approach["upstreamFt"],
        "downstreamFt": approach["downstreamFt"],
        "workLenFt": work_len,
        "note": (
            None if ok else (
                f"Path is {path_len:.0f} ft; this sheet needs "
                f"{approach['upstreamFt']:.0f} ft upstream + "
                f"{approach['downstreamFt']:.0f} ft downstream"
                + (f" + {work_len:.0f} ft work bay" if work_len else "")
                + f" = {need:.0f} ft. Extend the road by {short:.0f} ft, "
                "or pick a longer chain."
            )
        ),
    }
