"""Live 619-311 on a sustained constant-radius C-curve two-way real road.

Sibling of build_619311_on_real_road.py (straight) and
build_619311_on_curved_road.py (L-bend). This one uses a long, steady
highway curve so the whole work bay — hatch, cones, dim spans — sits on
real curvature rather than a single filleted corner.

Urban 55 needs ~2400 ft of approach each side, so the arc is long and
gentle: R = 3000 ft swept ~110 deg is ~5760 ft of centerline.

Set ARC_DIMS=1 in the environment to place bowed spans as REAL annotative
Arc Size DimensionElements (msdDimTypeArcSize) instead of the constructed
ArcElement graphics — see scripts/diag_arc_size_root_cause.py.
"""
from __future__ import annotations

import math
import os
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import alignment_geometry as ag  # noqa: E402
import lane_highway as lh  # noqa: E402
import view_capture  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)

OUT = ROOT / "Bridge" / "captures"
OUT.mkdir(parents=True, exist_ok=True)

# Fresh band. 84000 was used for the 2026-08-13 iteration runs and still holds
# orphaned dim text from the constructed-dim era (journal rotation broke
# clear_plan_elements' ownership proof), so this moved clear of it.
ORIGIN_X = float(os.environ.get("ORIGIN_X", "100000"))
ORIGIN_Y = 288000.0
LANE = 12.0
YELLOW_GAP = 2.0
SHOULDER = 8.0
SIDE = "right"
# Left edge of closed outer lane on undivided 4-lane, same as the straight
# and L-bend builds. Do not use CHAN_OFF=LANE — that reads as a center closure.
CHAN_OFF = 2 * LANE + YELLOW_GAP + LANE  # 38
WA_LEN = 100.0

CURVE_R = 3000.0
CURVE_SWEEP_DEG = 110.0


def _c_curve(bx: float, by: float) -> list[list[float]]:
    """Constant-radius arc, sampled fine enough to read as a true curve."""
    sweep = math.radians(CURVE_SWEEP_DEG)
    cx, cy = bx, by + CURVE_R
    n = 72
    pts = []
    for i in range(n + 1):
        a = -math.pi / 2 + sweep * i / n
        pts.append([cx + CURVE_R * math.cos(a), cy + CURVE_R * math.sin(a)])
    return pts


def _offset_path(verts: list[list[float]], off: float) -> list[list[float]]:
    segs = lh._prepare_path_segments(
        lh.resolve_edge_path(0, 0, 0, 0, verts),
        fillet_radius_ft=150.0,
    )
    return lh._continuous_offset_polyline(segs, off, SIDE, step_ft=25.0)


def _pt_on_path(segs, sta: float) -> list[float]:
    x, y, _, _ = ag.point_at_extended(segs, sta)
    return [x, y, 0.0]


def main() -> int:
    # Honor the module default unless ARC_DIMS is explicitly set (ARC_DIMS=0
    # forces the constructed fallback for an A/B).
    env = os.environ.get("ARC_DIMS", "")
    if env != "":
        ops.ARC_SIZE_BEND_DIMS = env not in ("0", "false")
    arc_dims = ops.ARC_SIZE_BEND_DIMS
    print(f"bend dims = {'REAL Arc Size (annotative)' if arc_dims else 'constructed ArcElement'}")

    outer = _c_curve(ORIGIN_X, ORIGIN_Y)
    align_verts = _offset_path(outer, CHAN_OFF)
    align_segs = ag.segments_from_polyline(align_verts)
    total = ag.total_length(align_segs)
    print(f"outer verts={len(outer)} align verts={len(align_verts)} "
          f"alignLen={total:.1f} ft R={CURVE_R:.0f} chanOff={CHAN_OFF}")

    # Work bay at mid-curve, where curvature is unambiguous.
    sta_up = max(0.0, total * 0.5 - WA_LEN * 0.35)
    sta_dn = min(total, sta_up + WA_LEN)
    up = _pt_on_path(align_segs, sta_up)
    dn = _pt_on_path(align_segs, sta_dn)
    print(f"WA up={[round(v,1) for v in up[:2]]} "
          f"dn={[round(v,1) for v in dn[:2]]} sta={sta_up:.1f}..{sta_dn:.1f}")

    print("clear", ops.clear_plan_elements(keep_alignments=False).get("deleted"))

    # Wipe AFTER clear so non-journal leftovers (old chords, prior curved-plan
    # arcs/tips) are gone before rebuild — engineer: remove bad dims.
    pad = CURVE_R + 400
    wipe = ops.delete_dimension_elements_in_range(
        ORIGIN_X - 400, ORIGIN_Y - 400,
        ORIGIN_X + pad, ORIGIN_Y + pad,
        reason="wipe leftover bad dims before C-curve rebuild")
    print("wipe_dims", wipe.get("deleted"), wipe.get("status"))

    ot = ops.build_wztc_order_table(
        speed=55, road_type="Non-Freeway", lane_width=12,
        shoulder_width=">= 8 ft", sheet_num="619-311", area_type="URBAN",
    )
    print("order_table", ot.get("status"),
          "signs", len(ops._PLAN_SESSION.locked_sign_rows))

    lat = ops.resolve_sheet_lateral(
        up, dn, closed_side="right", real_road_edge=True,
        path_vertices=align_verts,
    )
    print("lateral", lat.get("outward_sign"), lat.get("half_len"),
          "curved", lat.get("curved"))

    t0 = time.time()
    result = ops.run_sheet_build(
        upstream_edge=up, downstream_edge=dn,
        path_vertices=align_verts,
        arrow_panel_choice="trailer",
        include_visual_qa=True,
        force=True,
    )
    print(f"run_sheet_build in {time.time() - t0:.1f}s "
          f"status={result.get('status')}")
    for p in result.get("phases") or []:
        detail = p.get("result") or p.get("note") or p.get("error") or ""
        if isinstance(detail, dict):
            detail = {k: detail.get(k) for k in (
                "status", "workAreaLengthFt", "curved", "placedCount",
                "passed", "note") if k in detail}
        print(f"  phase {p.get('phase')}: {detail}")

    road = ops.place_two_way_highway(
        lanes=4, vertices=outer,
        lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP,
        shoulder_width_ft=SHOULDER, side=SIDE,
        reason="619-311 C-curve real road",
    )
    print("corridor", road.get("status"),
          "placed", road.get("placedCount") or len(road.get("placed") or []),
          "errors", (road.get("errors") or [])[:3])

    guides = ops.delete_construction_guides()
    print("guides", guides.get("status"), guides.get("deleted"))

    sc = ops.get_geometry_scorecard("619-311")
    print("scorecard", sc.get("passed"), sc.get("failures") or sc.get("note"))

    mx = 0.5 * (up[0] + dn[0])
    my = 0.5 * (up[1] + dn[1])
    tag = "arcsize" if arc_dims else "constructed"
    captures = [
        (f"qa_311_ccurve_{tag}_overview", ORIGIN_X + 1200, ORIGIN_Y + 1800, 7000, 5200),
        (f"qa_311_ccurve_{tag}_work", mx, my, 520, 240),
        (f"qa_311_ccurve_{tag}_dims", mx, my, 1600, 700),
    ]
    for name, cx, cy, w, h in captures:
        view_capture.navigate_view(cx, cy, w, h, view_num=1)
        time.sleep(0.6)
        src = Path(view_capture.capture_microstation())
        dest = OUT / f"{name}.png"
        shutil.copy2(src, dest)
        print("saved", dest.name)

    print(f"LOOK HERE: work bay ~({mx:.0f}, {my:.0f})  curve origin "
          f"({ORIGIN_X:.0f}, {ORIGIN_Y:.0f}) R={CURVE_R:.0f}")
    return 0 if result.get("status") == "OK" and sc.get("passed") else 1


if __name__ == "__main__":
    raise SystemExit(main())
