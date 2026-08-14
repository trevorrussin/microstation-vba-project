"""Build overlap: ledger classify + Tier 1 stacks + Tier 2 station/offset.

AABB is a fetch prefilter only — never the verdict.
"""
from __future__ import annotations

import math
from collections import defaultdict
from typing import Any, Optional

import alignment_geometry as ag
import build_ledger

ORIGIN_TOL_FT = 80.0
OFFSET_PAD_FT = 20.0
PATH_STEP_FT = 50.0


def origin_xy(origin: list | tuple | None, path: list | None) -> tuple[float, float]:
    if origin and len(origin) >= 2:
        return float(origin[0]), float(origin[1])
    if path and len(path) >= 1 and len(path[0]) >= 2:
        return float(path[0][0]), float(path[0][1])
    return (0.0, 0.0)


def same_origin(a: tuple[float, float], b: tuple[float, float],
                tol: float = ORIGIN_TOL_FT) -> bool:
    return math.hypot(a[0] - b[0], a[1] - b[1]) <= tol


def stack_key(row: dict) -> tuple:
    cx = round(float(row.get("cx") or 0), 1)
    cy = round(float(row.get("cy") or 0), 1)
    w = round(float(row.get("w") or 0), 1)
    h = round(float(row.get("h") or 0), 1)
    t = str(row.get("type") or "OTHER").upper()
    text = str(row.get("text") or "").strip()
    return (t, cx, cy, w, h, text)


def tier1_duplicates(model_rows: list[dict]) -> list[dict]:
    """Exact-duplicate hash. count>1 at the same rounded center/size/text."""
    buckets: dict[tuple, list[dict]] = defaultdict(list)
    for r in model_rows or []:
        buckets[stack_key(r)].append(r)
    out = []
    for key, group in buckets.items():
        if len(group) < 2:
            continue
        ids = [str(g.get("elementId") or "") for g in group if g.get("elementId")]
        out.append({
            "key": list(key),
            "count": len(group),
            "elementIds": ids[:12],
        })
    out.sort(key=lambda d: -int(d["count"]))
    return out


def _path_length(pts: list) -> float:
    n = 0.0
    for i in range(1, len(pts)):
        n += math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1])
    return n


def corridor_bbox(path: list, pad: float = 80.0) -> dict:
    xs = [float(p[0]) for p in path if p and len(p) >= 2]
    ys = [float(p[1]) for p in path if p and len(p) >= 2]
    if not xs:
        return {}
    return {
        "lowX": min(xs) - pad, "lowY": min(ys) - pad,
        "highX": max(xs) + pad, "highY": max(ys) + pad,
    }


def aabb_overlap(a: dict, b: dict) -> bool:
    """Prefetch only. Not a collision verdict."""
    if not a or not b:
        return False
    return not (
        float(a["highX"]) < float(b["lowX"])
        or float(a["lowX"]) > float(b["highX"])
        or float(a["highY"]) < float(b["lowY"])
        or float(a["lowY"]) > float(b["highY"])
    )


def tier2_path_conflict(
    our_path: list,
    our_half: float,
    other: dict,
) -> Optional[dict]:
    """Station/offset overlap vs another ledger build. No COM."""
    other_path = other.get("path_vertices") or []
    if len(our_path) < 2 or len(other_path) < 2:
        return None
    segs = ag.segments_from_polyline(our_path)
    other_half = float(other.get("lateralHalfWidth") or 40.0)
    band = our_half + other_half + OFFSET_PAD_FT
    total = _path_length(other_path)
    if total < 1.0:
        return None
    n = max(2, int(total / PATH_STEP_FT) + 1)
    hits = []
    for i in range(n):
        t = i / (n - 1)
        # sample along other polyline by cumulative length
        target = t * total
        acc = 0.0
        x, y = other_path[0][0], other_path[0][1]
        for k in range(1, len(other_path)):
            dx = other_path[k][0] - other_path[k - 1][0]
            dy = other_path[k][1] - other_path[k - 1][1]
            seg = math.hypot(dx, dy)
            if acc + seg >= target:
                u = 0.0 if seg < 1e-9 else (target - acc) / seg
                x = other_path[k - 1][0] + u * dx
                y = other_path[k - 1][1] + u * dy
                break
            acc += seg
            x, y = other_path[k][0], other_path[k][1]
        sta, dist = ag.nearest_station(segs, x, y)
        if dist <= band:
            hits.append({"sta": round(sta, 1), "offsetFt": round(dist, 1)})
    if not hits:
        return None
    stas = [h["sta"] for h in hits]
    return {
        "otherBuildId": other.get("buildId"),
        "otherSheet": other.get("sheetNum"),
        "staMin": min(stas),
        "staMax": max(stas),
        "hitCount": len(hits),
        "sample": hits[:4],
    }


def tier2_model_conflicts(
    our_path: list,
    our_half: float,
    sta0: float,
    sta1: float,
    model_rows: list[dict],
    ignore_ids: set[str] | None = None,
) -> list[dict]:
    """Live leftovers: element center projects into the plan band."""
    if len(our_path) < 2:
        return []
    segs = ag.segments_from_polyline(our_path)
    ignore = ignore_ids or set()
    lo, hi = (sta0, sta1) if sta1 >= sta0 else (sta1, sta0)
    if hi - lo < 1.0:
        hi = lo + ag.total_length(segs)
    band = our_half + OFFSET_PAD_FT
    conflicts = []
    for r in model_rows or []:
        eid = str(r.get("elementId") or "")
        if eid in ignore:
            continue
        try:
            cx = float(r.get("cx"))
            cy = float(r.get("cy"))
        except (TypeError, ValueError):
            continue
        sta, dist = ag.nearest_station(segs, cx, cy)
        if dist <= band and lo - 1.0 <= sta <= hi + 1.0:
            conflicts.append({
                "elementId": eid,
                "type": r.get("type"),
                "sta": round(sta, 1),
                "offsetFt": round(dist, 1),
                "text": (r.get("text") or "")[:40],
            })
    return conflicts[:40]


def classify(
    *,
    sheet_num: str,
    origin: list | None,
    path_vertices: list | None,
    lateral_half_width: float = 40.0,
    sta0: float = 0.0,
    sta1: float = 0.0,
    model_rows: list | None = None,
    ignore_ids: set[str] | None = None,
    ledger_rows: list | None = None,
) -> dict[str, Any]:
    ox, oy = origin_xy(origin, path_vertices)
    path = list(path_vertices or [])
    half = float(lateral_half_width or 40.0)
    ledger = ledger_rows if ledger_rows is not None else build_ledger.load_builds()
    bbox = corridor_bbox(path) if path else {}

    rebuild = []
    same_sheet_hits = []
    other_sheet_hits = []
    for rec in ledger:
        o = rec.get("origin") or [0, 0]
        other_xy = (float(o[0]), float(o[1]))
        same_s = str(rec.get("sheetNum") or "") == str(sheet_num or "")
        if same_s and same_origin((ox, oy), other_xy):
            rebuild.append(rec.get("buildId"))
            continue
        conf = None
        if path and rec.get("path_vertices"):
            if bbox and rec.get("bbox") and not aabb_overlap(bbox, rec["bbox"]):
                conf = None
            else:
                conf = tier2_path_conflict(path, half, rec)
        if not conf:
            continue
        if same_s:
            same_sheet_hits.append(conf)
        else:
            other_sheet_hits.append(conf)

    dups = tier1_duplicates(model_rows or [])
    live = tier2_model_conflicts(
        path, half, sta0, sta1, model_rows or [], ignore_ids)

    if other_sheet_hits:
        verdict = "collision_other_sheet"
        next_tool = "ask_user_choice"
        next_step = "Ask the engineer — another sheet's build overlaps this corridor."
        msg = "PLAN_OVERLAP: different sheet overlapping this corridor."
    elif same_sheet_hits:
        verdict = "collision_same_sheet"
        next_tool = "clear_plan_elements"
        next_step = "Wipe the other origin or pick a clear band, then rebuild."
        msg = "PLAN_OVERLAP: same sheet, different origin, corridors intersect."
    elif rebuild:
        verdict = "rebuild_same_origin"
        next_tool = "clear_plan_elements"
        next_step = "Same sheet at this origin — call clear_plan_elements first (routine)."
        msg = "PLAN_OVERLAP: rebuilding the same sheet at the same origin."
    elif dups:
        verdict = "stacked_duplicates"
        next_tool = "clear_plan_elements"
        next_step = "Tier 1 stacks in the model — wipe leftovers then rebuild."
        msg = "PLAN_OVERLAP: stacked duplicates in the model (same center/size/text)."
    else:
        verdict = "ok"
        next_tool = ""
        next_step = ""
        msg = ""

    return {
        "code": "PLAN_OVERLAP" if verdict != "ok" else "",
        "verdict": verdict,
        "blocking": False,
        "message": msg,
        "nextTool": next_tool,
        "nextStep": next_step,
        "currentStep": "corridor_ready",
        "duplicates": dups[:20],
        "conflicts": (other_sheet_hits + same_sheet_hits)[:12],
        "liveConflicts": live[:20],
        "priorBuildIds": rebuild[-5:],
        "hint": "call get_plan_status() — do not invent a workaround",
    }
