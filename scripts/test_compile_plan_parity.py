"""Golden-file parity test for the placement-plan compiler, Stage 1 (see
Data/sheet-specs/STATUS.md). Two layers:

1. Geometry engine parity (live, requires MicroStation): for a committed
   test alignment, diff alignment_geometry.station_to_xy() against the
   live GetPointAndTangent()-backed STATION_TO_POINT bridge op at a range
   of stations, including segment-boundary and out-of-range (clamped)
   cases. This is the part of the compiler doing genuinely new work
   (Python replicating VBA's arc-length walk instead of calling it per
   point) and is worth checking against live ground truth every time,
   not just trusted from a one-off manual check.

2. compile_plan() smoke test (no MicroStation needed once segments are
   captured): confirms compile_plan() runs against a real sheet spec
   (619-311) and produces the expected primitive kinds/order/text without
   crashing. The dimension/label offset FORMULA itself (PERP_HALF_LEN=40,
   offsetDist=15, textExtraAlong=20, outwardSign=-1.0 default) is a direct
   line-for-line transcription of Modules/PerpPlacement.bas's
   PlaceOrderTableDimensions/PlaceOrderTableLabels -- verified by reading,
   not by an independent rendered-plan pixel diff. A full diff against
   619-311's actually-drawn dimension/label elements (via find_elements_near)
   is the natural next step if this compiler goes beyond a proof of
   architecture, but isn't required to validate the geometry engine itself.

Requires a committed test alignment (see AlignmentGeometryError below for
how to make one) and MicroStation open with WZTCBridge polling for layer 1.
Layer 2 only needs a committed alignment's already-fetched vertices, not a
live session, if you pass --vertices-only with a saved GET_ALIGNMENT_VERTICES
dump -- not implemented here since layer 1 already requires live access.

Usage:
    python scripts/test_compile_plan_parity.py --align-idx 9
"""
from __future__ import annotations

import argparse
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))
import alignment_geometry as ag  # noqa: E402
import bridge_client  # noqa: E402
import sheet_spec  # noqa: E402
import wztc_ops  # noqa: E402
from sheet_spec import check_taper_continuity  # noqa: E402

TOL = 0.001


def check_geometry_parity(align_idx: int) -> list[str]:
    """Layer 1. Returns a list of failure strings (empty = all passed)."""
    fails = []
    rows = wztc_ops.get_alignment_vertices(align_idx)
    segs = ag.parse_vertices(rows)
    total = ag.total_length(segs)

    # Sample every segment boundary plus midpoints, plus out-of-range on
    # both ends to exercise the clamp path.
    stations = {0.0, total, -10.0, total + 50.0}
    cum = 0.0
    for seg in segs:
        stations.add(cum)
        stations.add(cum + seg.seg_len / 2.0)
        cum += seg.seg_len
        stations.add(cum)

    for sta in sorted(stations):
        vba = wztc_ops.station_to_point(align_idx, sta)
        vx, vy = float(vba["ptX"]), float(vba["ptY"])
        vtx, vty = float(vba["tanX"]), float(vba["tanY"])
        px, py, ptx, pty = ag.station_to_xy(segs, sta)
        if not (abs(vx - px) < TOL and abs(vy - py) < TOL
                and abs(vtx - ptx) < TOL and abs(vty - pty) < TOL):
            fails.append(
                f"station={sta}: VBA=({vx:.4f},{vy:.4f},{vtx:.3f},{vty:.3f}) "
                f"Python=({px:.4f},{py:.4f},{ptx:.3f},{pty:.3f})")
    print(f"geometry parity: {len(stations)} stations checked, {len(fails)} mismatches")
    return fails


def check_compile_plan_smoke(align_idx: int) -> list[str]:
    """Layer 2. Returns a list of failure strings (empty = all passed)."""
    fails = []
    rows = wztc_ops.get_alignment_vertices(align_idx)
    segs = ag.parse_vertices(rows)

    spec = sheet_spec.load("619-311")
    if spec is None:
        fails.append("Data/sheet-specs/619-311.json not found")
        return fails

    resolved = sheet_spec.resolve(spec, 45, 12, "12 ft", "URBAN",
                                   "LANE CLOSURE OR ENCROACHMENT",
                                   "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC")
    prims = sheet_spec.compile_plan(
        spec, resolved, align_idx=1, segments=segs,
        sheet_elements="MergingTaper|ShoulderTaper|DownstreamTaper|"
                       "ProtectiveVehicle|ArrowPanel|ChannelizingDevices")

    kinds = [p["kind"] for p in prims]
    if not kinds:
        fails.append("compile_plan produced no primitives at all")
    if "station" not in kinds:
        fails.append("no 'station' primitives produced")
    if "dimension" not in kinds:
        fails.append("no 'dimension' primitives produced")

    expected_texts = {"80", "360", "560"}  # roll ahead / buffer / lane taper, 45mph/12ft/urban
    dim_texts = {p["text"] for p in prims if p["kind"] == "dimension"}
    missing = expected_texts - dim_texts
    if missing:
        fails.append(f"expected dimension texts {missing} not found in {dim_texts}")

    print(f"compile_plan smoke: {len(prims)} primitives "
          f"({kinds.count('station')} station, {kinds.count('dimension')} dimension, "
          f"{kinds.count('label')} label)")
    return fails


def check_channelizing_smoke(align_idx: int) -> list[str]:
    """Layer 3 (Stage 2): device counts match the resolved spec exactly, and
    taper-continuity holds at every run junction (the actual bug this stage
    fixes -- PerpPlacement.PlaceOrderTableChannelizing's laneWidthFt*0.35
    fudge factor left a lateral jog at exactly this junction)."""
    fails = []
    rows = wztc_ops.get_alignment_vertices(align_idx)
    segs = ag.parse_vertices(rows)

    spec = sheet_spec.load("619-311")
    resolved = sheet_spec.resolve(spec, 45, 12, "12 ft", "URBAN",
                                   "LANE CLOSURE OR ENCROACHMENT",
                                   "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC")
    prims = sheet_spec.compile_channelizing(spec, resolved, align_idx=1, segments=segs,
                                             lane_width_ft=12, shoulder_width_ft=8)

    from collections import Counter
    counts = Counter(p["run"] for p in prims)
    expect = {"laneTaperRun": resolved["laneTaper"]["devices"],
              "shoulderTaperRun": resolved["shoulderTaper"]["devices"]}
    for run_id, expected_count in expect.items():
        actual = counts.get(run_id, 0)
        if actual != expected_count:
            fails.append(f"{run_id}: expected {expected_count} devices (from resolved spec), got {actual}")

    fails += check_taper_continuity(prims)

    print(f"channelizing smoke: {len(prims)} cones across {len(counts)} runs {dict(counts)}")
    return fails


def check_symbols_smoke(align_idx: int) -> list[str]:
    """Layer 4 (Stage 3): every symbol item with a stationAnchor compiles,
    the arrow-panel/VEH alternative is cross-linked (not duplicated), and
    vehicle-mounted signs attach to their vehicle's own position. Checked
    against both 619-311 (2 PVs, one conditional) and 619-302 (3 PVs) so
    this isn't just proven on the sheet it was written against."""
    fails = []
    rows = wztc_ops.get_alignment_vertices(align_idx)
    segs = ag.parse_vertices(rows)

    for sheet_num in ("619-311", "619-302"):
        spec = sheet_spec.load(sheet_num)
        resolved = sheet_spec.resolve(spec, 45, 12, "12 ft", "URBAN",
                                       "LANE CLOSURE OR ENCROACHMENT",
                                       "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC")
        prims = sheet_spec.compile_symbols(spec, resolved, align_idx=1, segments=segs)

        expected_pvs = sum(1 for s in spec["symbols"]["items"]
                            if s.get("cellHint") and s.get("stationAnchor"))
        actual_pvs = sum(1 for p in prims if p["kind"] == "protectiveVehicle")
        if actual_pvs != expected_pvs:
            fails.append(f"{sheet_num}: expected {expected_pvs} protective vehicles, got {actual_pvs}")

        alt_groups: dict[str, int] = {}
        for p in prims:
            if p.get("altGroup"):
                alt_groups[p["altGroup"]] = alt_groups.get(p["altGroup"], 0) + 1
        for group, count in alt_groups.items():
            if count != 2:
                fails.append(f"{sheet_num}: altGroup {group} has {count} members, expected exactly 2 (cross-linked pair)")

        mounted = [p for p in prims if p["kind"] == "vehicleMountedSign"]
        for m in mounted:
            host = next((p for p in prims if p.get("id") == m["mountedOn"]), None)
            if host is None:
                fails.append(f"{sheet_num}: vehicleMountedSign {m['signCode']} mountedOn "
                             f"{m['mountedOn']!r} which has no compiled primitive")
            elif abs(host["x"] - m["x"]) > 0.001 or abs(host["y"] - m["y"]) > 0.001:
                fails.append(f"{sheet_num}: vehicleMountedSign {m['signCode']} position doesn't "
                             f"match its host vehicle's position")

        print(f"symbols smoke [{sheet_num}]: {len(prims)} primitives, "
              f"{actual_pvs} PVs, {len(alt_groups)} alt group(s), {len(mounted)} mounted sign(s)")

    return fails


def check_hatch_smoke(align1_idx: int, align2_idx: int) -> list[str]:
    """Layer 5 (Stage 4): hatch boundary is a valid non-degenerate
    quadrilateral spanning align1's and align2's own committed station-0
    points (not an externally-supplied length -- see compile_hatch's
    module note for why that redesign happened), and the conditional
    transverse-run threshold (800 ft) fires only when it should.

    The live pair (align1_idx/align2_idx, expected several hundred+ ft
    apart) proves real committed-geometry correctness and the
    transverse-triggering case. The "short, no transverse" case is checked
    against synthetic segments (no MicroStation call needed) since the two
    live test alignments already committed this session are >800 ft apart
    and re-committing a third alignment just for this isn't worth a new
    bridge round trip."""
    fails = []
    segs1 = ag.parse_vertices(wztc_ops.get_alignment_vertices(align1_idx))
    segs2 = ag.parse_vertices(wztc_ops.get_alignment_vertices(align2_idx))

    spec = sheet_spec.load("619-311")
    resolved = sheet_spec.resolve(spec, 45, 12, "12 ft", "URBAN",
                                   "LANE CLOSURE OR ENCROACHMENT",
                                   "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC")

    # Synthetic short pair: two straight segments 200 ft apart, both
    # starting at station 0 -- proves the < 800 ft / no-transverse path
    # without needing a third live alignment.
    short_segs1 = [ag.PathSegment(False, 0, 0, 0, 0, 100, 0, 100.0)]
    short_segs2 = [ag.PathSegment(False, 0, 200, 0, 0, 300, 0, 100.0)]
    short_prims = sheet_spec.compile_hatch(spec, resolved, align1_segments=short_segs1,
                                            align2_segments=short_segs2,
                                            lane_width_ft=12, shoulder_width_ft=8)
    hatches = [p for p in short_prims if p["kind"] == "hatch"]
    if len(hatches) != 1:
        fails.append(f"expected exactly 1 hatch primitive, got {len(hatches)}")
    elif len(hatches[0]["boundary"]) != 4:
        fails.append(f"hatch boundary has {len(hatches[0]['boundary'])} corners, expected 4")
    elif abs(hatches[0]["workAreaLengthFt"] - 200.0) > 0.01:
        fails.append(f"synthetic 200 ft pair should give workAreaLengthFt=200, "
                     f"got {hatches[0]['workAreaLengthFt']:.1f}")
    if any(p["kind"] == "transverseRun" for p in short_prims):
        fails.append("200 ft work area (< 800 ft threshold) should not produce transverse runs")

    long_prims = sheet_spec.compile_hatch(spec, resolved, align1_segments=segs1,
                                           align2_segments=segs2,
                                           lane_width_ft=12, shoulder_width_ft=8)
    transverse = [p for p in long_prims if p["kind"] == "transverseRun"]
    long_len = next(p["workAreaLengthFt"] for p in long_prims if p["kind"] == "hatch")
    expected_transverse = int(long_len // 800)
    if long_len <= 800:
        fails.append(f"live align pair is only {long_len:.0f} ft apart -- expected > 800 ft "
                     f"to exercise the transverse-run path")
    elif len(transverse) != expected_transverse:
        fails.append(f"{long_len:.0f} ft work area should produce {expected_transverse} "
                     f"transverse runs (every 800 ft), got {len(transverse)}")

    print(f"hatch smoke: synthetic(len=200ft)->{len(hatches)} hatch/0 transverse, "
          f"live(len={long_len:.0f}ft)->{len(transverse)} transverse")
    return fails


def check_rules_gate(align_idx: int) -> list[str]:
    """Layer 6 (Stage 5): the gate itself passes on a correctly-compiled
    619-311 plan (no false positives), and correctly FAILS when a known
    bad plan is fed to it (no false negatives) -- both directions matter,
    a gate that never fires is as useless as one that always does."""
    fails = []
    segs = ag.parse_vertices(wztc_ops.get_alignment_vertices(align_idx))

    spec = sheet_spec.load("619-311")
    resolved = sheet_spec.resolve(spec, 45, 12, "12 ft", "URBAN",
                                   "LANE CLOSURE OR ENCROACHMENT",
                                   "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC")

    plan = sheet_spec.compile_plan(spec, resolved, align_idx=1, segments=segs,
                                    sheet_elements="MergingTaper|ShoulderTaper|DownstreamTaper|"
                                                   "ProtectiveVehicle|ArrowPanel|ChannelizingDevices")
    chan = sheet_spec.compile_channelizing(spec, resolved, align_idx=1, segments=segs,
                                            lane_width_ft=12, shoulder_width_ft=8)
    syms = sheet_spec.compile_symbols(spec, resolved, align_idx=1, segments=segs)

    clean_fails = sheet_spec.run_rules_gate(spec, resolved, align_idx=1,
                                             plan_primitives=plan, channelizing_primitives=chan,
                                             symbol_primitives=syms)
    if clean_fails:
        fails.append(f"gate reported failures on a correctly-compiled plan (false positive): {clean_fails}")

    # Deliberately break taper-continuity (the exact bug this stage fixes)
    # and confirm the gate catches it.
    broken_chan = [dict(p) for p in chan]
    for p in broken_chan:
        if p["run"] == "shoulderTaperRun":
            p["x"] += 5.0  # introduce a fake lateral jog
    broken_fails = sheet_spec.run_rules_gate(spec, resolved, align_idx=1,
                                              plan_primitives=plan, channelizing_primitives=broken_chan,
                                              symbol_primitives=syms)
    if not any("taper-continuity" in f for f in broken_fails):
        fails.append("gate did NOT catch a deliberately-introduced taper discontinuity (false negative)")

    print(f"rules gate: clean plan -> {len(clean_fails)} failures (expect 0), "
          f"broken plan -> {len(broken_fails)} failures (expect >=1, taper-continuity)")
    return fails


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--align-idx", type=int, default=9,
                     help="committed alignment to test against (default: 9, "
                          "the Stage 1 throwaway test alignment)")
    args = ap.parse_args()

    wztc_ops.set_bridge(bridge_client.bridge)

    all_fails = []
    all_fails += [f"[geometry] {f}" for f in check_geometry_parity(args.align_idx)]
    all_fails += [f"[compile_plan] {f}" for f in check_compile_plan_smoke(args.align_idx)]
    all_fails += [f"[channelizing] {f}" for f in check_channelizing_smoke(args.align_idx)]
    all_fails += [f"[symbols] {f}" for f in check_symbols_smoke(args.align_idx)]
    all_fails += [f"[hatch] {f}" for f in check_hatch_smoke(args.align_idx, 10)]
    all_fails += [f"[rules-gate] {f}" for f in check_rules_gate(args.align_idx)]

    print()
    if all_fails:
        print(f"FAILED — {len(all_fails)} issue(s)")
        for f in all_fails:
            print("  ", f)
        return 1
    print("ALL CHECKS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
