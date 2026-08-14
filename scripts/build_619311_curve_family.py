"""Fresh-band 619-311 L / C / S builds, 1000 ft clear of each other.

Session reset only — no clear_plan_elements — so siblings stay in the DGN.
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
import sheet_compile as scmp  # noqa: E402
import view_capture  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(bridge)

OUT = ROOT / "Bridge" / "captures"
OUT.mkdir(parents=True, exist_ok=True)

LANE = 12.0
YELLOW_GAP = 2.0
SHOULDER = 8.0
SIDE = "right"
CHAN_OFF = 2 * LANE + YELLOW_GAP + LANE  # 38
WA_LEN = 100.0
CURVE_R = 3000.0
CURVE_SWEEP_DEG = 110.0

# Fresh band north of the 76k/288k family (2026-08-13).
L_OX, L_OY = 90000.0, 300000.0
L_LEG = 2600.0
C_OX = L_OX + L_LEG + 1000.0  # 93600
C_OY = L_OY
STRAIGHT_OX, STRAIGHT_OY = L_OX, L_OY - 1000.0
STRAIGHT_LEN = 5600.0


def _c_curve(bx: float, by: float) -> list[list[float]]:
    sweep = math.radians(CURVE_SWEEP_DEG)
    cx, cy = bx, by + CURVE_R
    n = 72
    pts = []
    for i in range(n + 1):
        a = -math.pi / 2 + sweep * i / n
        pts.append([cx + CURVE_R * math.cos(a), cy + CURVE_R * math.sin(a)])
    return pts


def _c_east_extent(bx: float, by: float) -> float:
    return max(p[0] for p in _c_curve(bx, by))


def _s_curve(bx: float, by: float) -> list[list[float]]:
    return [
        [bx, by],
        [bx + 2200.0, by],
        [bx + 2800.0, by + 600.0],
        [bx + 3400.0, by],
        [bx + 5600.0, by],
    ]


def _straight(bx: float, by: float) -> list[list[float]]:
    return [[bx, by], [bx + STRAIGHT_LEN, by]]


def _offset_path(verts: list[list[float]], off: float) -> list[list[float]]:
    segs = lh._prepare_path_segments(
        lh.resolve_edge_path(0, 0, 0, 0, verts),
        fillet_radius_ft=150.0,
    )
    return lh._continuous_offset_polyline(segs, off, SIDE, step_ft=25.0)


def _pt_on_path(segs, sta: float) -> list[float]:
    x, y, _, _ = ag.point_at_extended(segs, sta)
    return [x, y, 0.0]


def _aabb(verts: list[list[float]], pad: float) -> tuple[float, float, float, float]:
    xs = [p[0] for p in verts]
    ys = [p[1] for p in verts]
    return min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad


def _session_reset_keep_dgn() -> None:
    """Drop Python/VBA plan memory without deleting prior DGN builds."""
    ops._PLAN_SESSION.reset()
    ops._clear_sheet_plan_file()
    ops.placement_registry.clear_registry()


def _place_highway(outer: list[list[float]], reason: str) -> None:
    road = ops.place_two_way_highway(
        lanes=4, vertices=outer,
        lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP,
        shoulder_width_ft=SHOULDER, side=SIDE,
        reason=reason,
    )
    print("corridor", road.get("status"),
          "placed", road.get("placedCount") or len(road.get("placed") or []),
          "errors", (road.get("errors") or [])[:3])
    if road.get("status") == "OK":
        return
    solids = [s for s in lh.two_way_highway_lines(
        4, vertices=outer, lane_width_ft=LANE, yellow_gap_ft=YELLOW_GAP,
        shoulder_width_ft=SHOULDER, side=SIDE,
    ) if s.get("style") == "solid"]
    sok = 0
    for s in solids:
        verts = s.get("vertices") or [[s["x1"], s["y1"]], [s["x2"], s["y2"]]]
        rr = ops.place_polyline(verts, reason=f"solid fallback {s.get('kind')}")
        if rr.get("status") == "OK":
            sok += 1
            color = 4 if s.get("kind") == "yellow" else 0
            ids = str(rr.get("createdElementIds") or rr.get("elementId") or "")
            for eid in ids.replace(",", " ").split():
                if eid.strip().isdigit():
                    ops.change_element_symbology(
                        eid.strip(), color=color, weight=0)
    print(f"solid fallback placed {sok}/{len(solids)}")


def _build_one(name: str, outer: list[list[float]],
               wa_frac: float, captures: list) -> dict:
    print(f"\n======== {name} ========")
    _session_reset_keep_dgn()
    align_verts = _offset_path(outer, CHAN_OFF)
    align_segs = ag.segments_from_polyline(align_verts)
    total = ag.total_length(align_segs)
    print(f"{name} outer={len(outer)} align={len(align_verts)} "
          f"len={total:.1f} ft")

    sta_up = max(0.0, total * wa_frac - WA_LEN * 0.35)
    sta_dn = min(total, sta_up + WA_LEN)
    up = _pt_on_path(align_segs, sta_up)
    dn = _pt_on_path(align_segs, sta_dn)
    print(f"WA up={[round(v, 1) for v in up[:2]]} "
          f"dn={[round(v, 1) for v in dn[:2]]} sta={sta_up:.1f}..{sta_dn:.1f}")
    # Lane taper is 615–1295 ft upstream of WA on Align1 (roll 120 + buffer 495).
    lt0 = max(0.0, sta_up - 1295.0)
    lt1 = max(0.0, sta_up - 615.0)
    if lt1 - lt0 > 50.0:
        tip = scmp._dim_tip_path(
            align_segs, lt0, lt1, 1.0, 20.0, step_ft=10.0)
        kind = scmp.classify_dim_path(tip)
        runs = scmp.split_dim_path_runs(tip)
        print(f"laneTaper classify={kind} runs="
              f"{[(k, round(scmp._path_length(p), 1)) for k, p in runs]}")
    dn_tip = scmp._dim_tip_path(
        align_segs, sta_dn, min(total, sta_dn + 50.0), 1.0, 20.0,
        step_ft=5.0, align_idx=2)
    print(f"downstream50 classify={scmp.classify_dim_path(dn_tip)} "
          f"headingDeg={math.degrees(scmp._path_heading_delta(dn_tip)):.2f}")

    x0, y0, x1, y1 = _aabb(outer, 450.0)
    ops.place_fence_block(x0, y0, x1, y1, reason=f"wipe {name} band")
    fd = ops.fence_delete_contents(reason=f"wipe {name} band")
    print("fence_wipe", fd.get("deleted"), fd.get("status"))
    ops.fence_undefine(reason="clear wipe fence")
    wipe = ops.delete_dimension_elements_in_range(
        x0, y0, x1, y1, reason=f"wipe leftover dims {name}")
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
        include_visual_qa=False,
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

    _place_highway(outer, f"619-311 {name} real road")
    guides = ops.delete_construction_guides()
    print("guides", guides.get("status"), guides.get("deleted"))
    sc = ops.get_geometry_scorecard("619-311")
    print("scorecard", sc.get("passed"), sc.get("failures") or sc.get("note"))
    compiled = ops._PLAN_SESSION.last_compiled or {}
    inner = compiled.get("plan") or compiled
    by_zone: dict = {}
    for a_str, prims in (inner.get("planByAlign") or {}).items():
        for p in prims:
            if p.get("kind") != "dimension":
                continue
            ref = p.get("specRef") or {}
            if int(ref.get("partCount") or 1) <= 1:
                continue
            by_zone.setdefault((a_str, ref.get("zone")), []).append(p)
    for (a_str, zone), parts in by_zone.items():
        lens = [float((p.get("specRef") or {}).get("partLengthFt") or 0) for p in parts]
        sheet = (parts[0].get("specRef") or {}).get("sheetLengthFt")
        kinds = [(p.get("specRef") or {}).get("partKind") for p in parts]
        print(f"  split {zone} align{a_str}: {kinds} {lens} "
              f"sum={sum(lens):.1f} sheet={sheet}")

    mx = 0.5 * (up[0] + dn[0])
    my = 0.5 * (up[1] + dn[1])
    for cap_name, cx, cy, w, h in captures(mx, my, up, dn):
        view_capture.navigate_view(cx, cy, w, h, view_num=1)
        time.sleep(0.8)
        src = Path(view_capture.capture_microstation())
        dest = OUT / f"{cap_name}.png"
        shutil.copy2(src, dest)
        print("saved", dest.name)

    return {
        "name": name,
        "status": result.get("status"),
        "scorecard": sc.get("passed"),
        "failures": sc.get("failures"),
        "work": (mx, my),
        "up": up[:2],
        "dn": dn[:2],
        "curved": lat.get("curved"),
    }


def main() -> int:
    want = (sys.argv[1] if len(sys.argv) > 1 else "LCS").upper()
    print(f"builds={want}  bend dims = "
          f"{'REAL Arc Size' if ops.ARC_SIZE_BEND_DIMS else 'constructed'}")
    c_east = _c_east_extent(C_OX, C_OY)
    s_ox = c_east + 1000.0
    s_oy = C_OY
    print(f"C origin ({C_OX:.0f},{C_OY:.0f}) east={c_east:.0f}")
    print(f"S origin ({s_ox:.0f},{s_oy:.0f})")
    print(f"L origin ({L_OX:.0f},{L_OY:.0f})")

    results = []

    def _l_bend():
        return [[L_OX, L_OY], [L_OX + L_LEG, L_OY],
                [L_OX + L_LEG, L_OY + L_LEG]]

    def caps_l(mx, my, up, dn):
        return [
            ("qa_311_fresh_l_overview", L_OX + 1400, L_OY + 200, 5000, 2800),
            ("qa_311_fresh_l_work", mx, my, 500, 220),
            ("qa_311_fresh_l_dims", mx, my, 1600, 700),
            ("qa_311_fresh_l_g20", dn[0], dn[1] + 80, 280, 220),
        ]

    if "L" in want:
        results.append(_build_one("L-bend", _l_bend(), 0.48, caps_l))

    def caps_c(mx, my, up, dn):
        return [
            ("qa_311_fresh_c_overview", C_OX + 1400, C_OY + 2000, 7200, 5400),
            ("qa_311_fresh_c_work", mx, my, 520, 240),
            ("qa_311_fresh_c_dims", mx, my, 1800, 800),
            ("qa_311_fresh_c_g20", dn[0] + 40, dn[1] + 40, 280, 220),
        ]

    if "C" in want:
        results.append(_build_one("C-curve", _c_curve(C_OX, C_OY), 0.5, caps_c))

    def caps_s(mx, my, up, dn):
        return [
            ("qa_311_fresh_s_overview", s_ox + 2800, s_oy + 200, 7000, 1800),
            ("qa_311_fresh_s_work", mx, my, 520, 240),
            ("qa_311_fresh_s_dims", mx, my, 2200, 1000),
            ("qa_311_fresh_s_lanetaper", up[0] - 800, up[1], 2500, 900),
            ("qa_311_fresh_s_g20", dn[0], dn[1], 280, 220),
        ]

    if "S" in want:
        results.append(_build_one("S-curve", _s_curve(s_ox, s_oy), 0.5, caps_s))

    print("\n======== SUMMARY ========")
    ok = True
    for r in results:
        print(r)
        if r["status"] != "OK" or not r["scorecard"]:
            ok = False
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
