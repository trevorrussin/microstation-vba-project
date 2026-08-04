"""Placement-plan compiler, Stage 5: a hard pre-draw gate over primitives
compiled by sheet_compile.py, checking the subset of each sheet's own
rules[] that's mechanically checkable from geometry alone (no MicroStation
needed). This is the "wrong is caught at compile time, not discovered in a
screenshot three passes later" goal from the original design discussion.

Honest scope: this does NOT cover every rules[] entry on every sheet --
some (e.g. sign-spacing-source, speed-range, shoulder-band-collapse) are
about table-lookup correctness, already covered by sheet_resolve.resolve()
and scripts/validate_sheet_spec.py's domain-invariant pass, not by
anything geometric a compiled plan could re-check. And this stage does
NOT retire Modules/PerpPlacement.bas's PlaceOrderTable* functions -- the
original plan's own bar for that ("golden test passes for at least one
reference sheet per family, 9 families") is real, substantial follow-up
work, not something this pass claims to have done. What IS delivered:
the gate mechanism itself, covering every geometric rule practical to
check today, run against two different sheet families (619-311, 619-302)
to prove it isn't 619-311-specific.

Part of the sheet_spec split (2026-08-04): sheet_resolve.py owns "what does
this sheet need", sheet_compile.py turns that into coordinates, this module
validates the result before it reaches the bridge. sheet_spec.py re-exports
all three so every existing `sheet_spec.X` call site keeps working
unchanged.
"""
from __future__ import annotations

from sheet_resolve import station_walk
from sheet_compile import check_taper_continuity, _zone_station_ranges


def check_corridor_topology(spec: dict, resolved: dict, align1_segments, align2_segments,
                             width_ft: float = 0.0) -> list[str]:
    """Standalone corridor-topology check — same math run_rules_gate's
    corridor-topology check uses (which reads it off compile_hatch's
    diagnostic fields), factored out so it can also run immediately after
    both alignments are committed/adopted, before place_sheet_geometry
    (and therefore before lane/shoulder width might even be known yet) —
    not only deep in the compile pipeline where a bad corridor could
    already have a full plan computed against it. See the live incident
    this guards against: Downstream committed +1000ft along the same line
    as Upstream, which looked like a valid work area but actually covered
    Upstream's own Roll Ahead + Buffer + taper zones.

    width_ft: lane+shoulder width if already known (tightens the
    collinearity tolerance); 0.0 falls back to a 20ft floor only."""
    import alignment_geometry as ag
    p1x, p1y, tan1x, tan1y = ag.station_to_xy(align1_segments, 0.0)
    p2x, p2y, _, _ = ag.station_to_xy(align2_segments, 0.0)
    dx, dy = p2x - p1x, p2y - p1y
    proj_sta = dx * tan1x + dy * tan1y
    perp_dist = abs(dx * (-tan1y) + dy * tan1x)

    align1_walk = [w for w in station_walk(spec, resolved) if w["alignIdx"] == 1]
    approach_max_sta = max(
        (w["stationFt"] for w in align1_walk if w["type"] == "Non-Sign"),
        default=0.0,
    )
    collinear_tol = max(width_ft, 20.0)
    fails: list[str] = []
    if 0.0 < proj_sta <= approach_max_sta and perp_dist < collinear_tol:
        fails.append(
            f"corridor-topology: align2 (Downstream) station-0 point projects onto "
            f"align1 (Upstream)'s own line at station {proj_sta:.1f} ft, inside "
            f"align1's approach-zone corridor (Roll Ahead/Buffer/taper span to "
            f"{approach_max_sta:.1f} ft) and only {perp_dist:.1f} ft off align1's "
            f"line -- align2 looks like a point further along align1's own "
            f"corridor, not a geometrically distinct work-area edge. Recommit "
            f"align2 at the actual downstream edge of the work area."
        )
    return fails


def run_rules_gate(spec: dict, resolved: dict, align_idx: int,
                    plan_primitives: list[dict], channelizing_primitives: list[dict],
                    symbol_primitives: list[dict], hatch_primitives: list[dict] | None = None) -> list[str]:
    """Returns failure strings (empty = gate passes). Call before any
    primitive list reaches the bridge."""
    fails: list[str] = []

    # taper-continuity (already built in Stage 2 -- reused, not duplicated)
    fails += [f"taper-continuity: {f}" for f in check_taper_continuity(channelizing_primitives)]

    # cone-spacing: no consecutive pair of cones in the same run may exceed
    # 40 ft of STATION separation (the sheet's own longitudinal spacing cap;
    # taper runs are exact by construction from the device count, so this
    # mainly catches a run that came out with too few points).
    by_run: dict[str, list[float]] = {}
    for p in channelizing_primitives:
        if p["kind"] == "cone":
            by_run.setdefault(p["run"], []).append(p["stationFt"])
    for run_id, stas in by_run.items():
        stas.sort()
        for a, b in zip(stas, stas[1:]):
            if b - a > 40.0 + 0.01:
                fails.append(f"cone-spacing: {run_id} has a {b - a:.1f} ft gap "
                             f"(stations {a:g}->{b:g}), exceeds the 40 ft cap")

    # sign-order: on the alignment station_walk gives us, Sign rows should
    # appear in strictly increasing station order walking upstream (this is
    # true by construction from the spec's own rowNum ordering -- checked
    # here as a real assertion, not assumed, in case a future spec authors
    # rows out of order).
    sign_stations = [(p["rowNum"], p["stationFt"]) for p in plan_primitives
                     if p["kind"] == "station"]
    for (n1, s1), (n2, s2) in zip(sign_stations, sign_stations[1:]):
        if s2 <= s1:
            fails.append(f"sign-order: row {n2} (station {s2:g}) is not further "
                         f"upstream than row {n1} (station {s1:g})")

    # arrow-panel-anchor: arrow panel and its station-matched taper zone tip
    # must be the same station (this is also what altGroup cross-linking to
    # a same-station vehicle alternative already implies, checked directly
    # here regardless of whether an alternative exists).
    zone_range = _zone_station_ranges(spec, resolved, align_idx)
    ap = next((p for p in symbol_primitives if p["kind"] == "arrowPanel"), None)
    if ap is not None:
        lane_rng = zone_range.get("laneTaper")
        if lane_rng is not None:
            expected_sta = max(lane_rng)  # upstream end = tip, per _anchor_station's convention
            if abs(ap["stationFt"] - expected_sta) > 0.01:
                fails.append(f"arrow-panel-anchor: arrow panel at station {ap['stationFt']:g}, "
                             f"expected laneTaper's upstream end at {expected_sta:g}")

    # no-occupancy-buffer-rollahead: compile_hatch's corrected design (see
    # its own module note) builds the hatch from align1/align2's own
    # station-0 points, which sit on the opposite side of station 0 from
    # rollAheadDistance/bufferSpace's positive-station ranges -- so an
    # overlap is structurally impossible by construction now, not just
    # checked for. What IS worth guarding: a degenerate (near-zero-length)
    # hatch, which usually means the two alignments were committed with
    # coincident or badly mismatched station-0 points rather than at the
    # work area's actual two edges.
    hatch = next((p for p in (hatch_primitives or []) if p.get("kind") == "hatch"), None)
    if hatch is not None and float(hatch["workAreaLengthFt"]) < 1.0:
        fails.append(f"no-occupancy-buffer-rollahead: work area length "
                     f"{hatch['workAreaLengthFt']:.2f} ft is near zero -- align1/align2 "
                     f"station-0 points may not be committed at the work area's real edges")

    # corridor-topology: a non-degenerate workAreaLengthFt (the check above)
    # is NOT sufficient -- align2's station-0 point can be far from align1's
    # AND still be wrong, if it was placed by walking further along align1's
    # own line instead of at a geometrically distinct work-area edge
    # (confirmed live 2026-08-04: Downstream offset +1000ft along the same
    # blank line as Upstream produced a "valid" 1000ft-long hatch that
    # actually covered Upstream's own Roll Ahead + Buffer + taper zones).
    # Detect it directly: if align2's station-0 point is both near-collinear
    # with align1's own line AND its projected station falls inside
    # align1's own approach-zone corridor (Roll Ahead/Buffer/tapers), the
    # two alignments are really the same line, not independent edges.
    if hatch is not None and "align2ProjectedStationOnAlign1Ft" in hatch:
        proj_sta = float(hatch["align2ProjectedStationOnAlign1Ft"])
        perp_dist = float(hatch["align2PerpDistFromAlign1Ft"])
        align1_walk = [w for w in station_walk(spec, resolved) if w["alignIdx"] == 1]
        approach_max_sta = max(
            (w["stationFt"] for w in align1_walk if w["type"] == "Non-Sign"),
            default=0.0,
        )
        collinear_tol = max(float(hatch.get("widthFt", 0.0)), 20.0)
        if 0.0 < proj_sta <= approach_max_sta and perp_dist < collinear_tol:
            fails.append(
                f"corridor-topology: align2 (Downstream) station-0 point projects onto "
                f"align1 (Upstream)'s own line at station {proj_sta:.1f} ft, inside "
                f"align1's approach-zone corridor (Roll Ahead/Buffer/taper span to "
                f"{approach_max_sta:.1f} ft) and only {perp_dist:.1f} ft off align1's "
                f"line -- align2 looks like a point further along align1's own "
                f"corridor, not a geometrically distinct work-area edge. Recommit "
                f"align2 at the actual downstream edge of the work area."
            )

    return fails
