"""Live 619-311 on a long curved two-way (reverse-S) real-road strip.

Urban 55 needs ~2400 ft approach each side — curve-matrix paths are too
short, so this builds an elongated reverse-S (~5600 ft) with the bend at
the work bay. Signs stay view-horizontal; corridor/hatch follow path_vertices.
"""
from __future__ import annotations

import math
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import bridge  # noqa: E402
import alignment_geometry as ag  # noqa: E402
import lane_highway as lh  # noqa: E402
import view_capture  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(bridge)

OUT = ROOT / "Bridge" / "captures"
OUT.mkdir(parents=True, exist_ok=True)

# Fresh band for L-bend rebuild (match straight real-road lateral).
ORIGIN_X = 76000.0
ORIGIN_Y = 288000.0
LANE = 12.0
YELLOW_GAP = 2.0
SHOULDER = 8.0
SIDE = "right"
# Same as build_619311_on_real_road.py: left edge of closed outer lane on
# undivided 4-lane = Y_north_outer − (2·L + gap + L). CHAN_OFF=LANE put
# hatch on the near-yellow lane and read as a center closure (QA 2026-08-13).
CHAN_OFF = 2 * LANE + YELLOW_GAP + LANE  # 38
WA_LEN = 100.0


def _long_l_bend(bx: float, by: float) -> list[list[float]]:
    """~5000+ ft first-travel-outer with a sharp L corner mid-corridor.

    Reverse-S with only +90 ft offset was nearly a straight diagonal through
    the work bay — hatch looked rectangular. L-bend puts real curvature
    (filleted) under the WA.
    """
    return [
        [bx + 0.0, by],
        [bx + 2600.0, by],
        [bx + 2600.0, by + 2600.0],
    ]


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
    outer = _long_l_bend(ORIGIN_X, ORIGIN_Y)
    align_verts = _offset_path(outer, CHAN_OFF)
    align_segs = ag.segments_from_polyline(align_verts)
    total = ag.total_length(align_segs)
    print(f"outer verts={len(outer)} align verts={len(align_verts)} "
          f"alignLen={total:.1f} ft chanOff={CHAN_OFF}")

    # Work bay straddles the L-corner (filleted) so hatch must bend.
    corner_sta = 2600.0  # near outer corner before fillet shrinks path
    # Snap to mid-path corner region via nearest station to corner XY.
    corner_xy = [ORIGIN_X + 2600.0, ORIGIN_Y + 0.0]  # approach to corner
    # Prefer station on align path closest to the geometric corner.
    c_sta, _ = ag.nearest_station(align_segs, ORIGIN_X + 2600.0, ORIGIN_Y)
    sta_up = max(0.0, c_sta - WA_LEN * 0.35)
    sta_dn = min(total, sta_up + WA_LEN)
    up = _pt_on_path(align_segs, sta_up)
    dn = _pt_on_path(align_segs, sta_dn)
    print(f"WA up={up[:2]} dn={dn[:2]} pathSta={sta_up:.1f}..{sta_dn:.1f} "
          f"(cornerSta={c_sta:.1f})")

    print("clear", ops.clear_plan_elements(keep_alignments=False).get("deleted"))

    # Journal clear misses prior-band leftovers (engineer: old build still
    # visible). Fence-delete the whole L-bend box, then dim wipe.
    ops.place_fence_block(
        ORIGIN_X - 200, ORIGIN_Y - 400,
        ORIGIN_X + 3200, ORIGIN_Y + 3200,
        reason="wipe old L-bend 619-311 before rebuild")
    fd = ops.fence_delete_contents(reason="wipe old L-bend 619-311")
    print("fence_wipe", fd.get("deleted"), fd.get("status"))
    ops.fence_undefine(reason="clear wipe fence")
    wipe = ops.delete_dimension_elements_in_range(
        ORIGIN_X - 200, ORIGIN_Y - 400,
        ORIGIN_X + 3200, ORIGIN_Y + 3200,
        reason="wipe leftover bad dims before curved rebuild")
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
          "curved", lat.get("curved"),
          "travel", lat.get("travelUnit"))

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
        phase = p.get("phase")
        detail = p.get("result") or p.get("note") or p.get("error") or ""
        if isinstance(detail, dict):
            detail = {k: detail.get(k) for k in (
                "status", "workAreaLengthFt", "curved", "placedCount",
                "passed", "note") if k in detail}
        print(f"  phase {phase}: {detail}")

    road = ops.place_two_way_highway(
        lanes=4, vertices=outer,
        lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP,
        shoulder_width_ft=SHOULDER, side=SIDE,
        reason="619-311 curved reverse-S real road",
    )
    print("corridor", road.get("status"),
          "placed", road.get("placedCount") or len(road.get("placed") or []),
          "errors", (road.get("errors") or [])[:3])
    if road.get("status") != "OK":
        # Long curved solids can still fail; place solids alone as fallback.
        import lane_highway as lh
        solids = [s for s in lh.two_way_highway_lines(
            4, vertices=outer, lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP,
            shoulder_width_ft=SHOULDER, side=SIDE,
        ) if s.get("style") == "solid"]
        sok = 0
        for s in solids:
            verts = s.get("vertices") or [[s["x1"], s["y1"]], [s["x2"], s["y2"]]]
            rr = ops.place_polyline(
                verts, reason=f"curved solid fallback {s.get('kind')}")
            if rr.get("status") == "OK":
                sok += 1
                color = 4 if s.get("kind") == "yellow" else 0
                ids = str(rr.get("createdElementIds") or rr.get("elementId") or "")
                for eid in ids.replace(",", " ").split():
                    if eid.strip().isdigit():
                        ops.change_element_symbology(
                            eid.strip(), color=color, weight=0)
        print(f"solid fallback placed {sok}/{len(solids)}")

    guides = ops.delete_construction_guides()
    print("guides", guides.get("status"), guides.get("deleted"))

    sc = ops.get_geometry_scorecard("619-311")
    print("scorecard", sc.get("passed"), sc.get("failures") or sc.get("note"))

    # Frame work-bay bend + overview
    mx = 0.5 * (up[0] + dn[0])
    my = 0.5 * (up[1] + dn[1])
    captures = [
        ("qa_311_curve_overview", ORIGIN_X + 2800, ORIGIN_Y + 40, 6000, 800),
        ("qa_311_curve_work", mx, my, 500, 220),
        ("qa_311_curve_upstream", up[0] - 400, up[1], 1200, 280),
        ("qa_311_curve_downstream", dn[0] + 200, dn[1], 900, 250),
        ("qa_311_curve_g20_post", dn[0], dn[1] + 400, 350, 350),
        ("qa_311_curve_w20_post", up[0] - 800, up[1], 350, 220),
    ]
    for name, cx, cy, w, h in captures:
        view_capture.navigate_view(cx, cy, w, h, view_num=1)
        time.sleep(0.4)
        src = Path(view_capture.capture_microstation())
        dest = OUT / f"{name}.png"
        shutil.copy2(src, dest)
        print("saved", dest.name)

    print(f"LOOK HERE: work bay ~({mx:.0f}, {my:.0f})  "
          f"outer path origin ({ORIGIN_X:.0f}, {ORIGIN_Y:.0f})")
    return 0 if result.get("status") == "OK" and sc.get("passed") else 1


if __name__ == "__main__":
    raise SystemExit(main())
