"""Shapely-backed geometric QA over compiled placement primitives.

These checks run inside sheet_rules.run_rules_gate after the station/
topology asserts. They catch planar defects the station walk cannot see
(self-intersecting hatch, PV sitting inside the work-area hatch, AP/PV
centers stacked). Failure strings only — never dump WKT into agent tool
results.

Optional dependency: shapely. If import fails, checks are skipped with
no failures (so a half-installed env does not brick compile); the
requirements.txt pin is the real gate for production.
"""
from __future__ import annotations

from typing import Any

# Nominal plan footprints (ft) for point-symbol collision checks — not
# cell bbox reads from MicroStation. Tuned to catch stacked AP/PV, not
# to model TWZWVA_P exactly. Same-station OR-alternatives (altGroup) are
# exempt — compile_symbols intentionally co-locates those.
_SYMBOL_PAIR_MIN_SEP_FT = 20.0


def _shapely():
    try:
        from shapely.geometry import LineString, Point, Polygon
        from shapely.validation import explain_validity
        return Point, Polygon, LineString, explain_validity
    except ImportError:
        return None


def _hatch_polygon(hatch: dict, Polygon):
    boundary = hatch.get("boundary") or []
    if len(boundary) < 3:
        return None
    coords = [(float(p[0]), float(p[1])) for p in boundary]
    # Close ring if needed
    if coords[0] != coords[-1]:
        coords = coords + [coords[0]]
    return Polygon(coords)


def check_compiled_geometry(symbol_primitives: list[dict] | None = None,
                            hatch_primitives: list[dict] | None = None,
                            channelizing_primitives: list[dict] | None = None
                            ) -> list[str]:
    """Return failure strings (empty = pass). Safe no-op if shapely missing."""
    mods = _shapely()
    if mods is None:
        return []
    Point, Polygon, LineString, explain_validity = mods
    fails: list[str] = []

    hatch = next((p for p in (hatch_primitives or []) if p.get("kind") == "hatch"), None)
    poly = None
    if hatch is not None:
        poly = _hatch_polygon(hatch, Polygon)
        if poly is None:
            fails.append("geometry-qa: hatch boundary has fewer than 3 vertices")
        elif not poly.is_valid:
            fails.append(
                f"geometry-qa: hatch polygon is invalid "
                f"({explain_validity(poly)}) — self-intersection or bad ring"
            )
        elif poly.is_empty or poly.area < 1.0:
            fails.append(
                f"geometry-qa: hatch polygon area {poly.area:.2f} sq ft is "
                f"degenerate (check work-area edges / lane width)"
            )
        elif not poly.exterior.is_simple:
            fails.append("geometry-qa: hatch exterior ring self-intersects")

    # PV / AP must not sit inside the work-area hatch (live miss: PV drawn
    # inside a hatch that wrongly covered buffer/roll-ahead).
    symbols = [
        p for p in (symbol_primitives or [])
        if p.get("kind") in ("protectiveVehicle", "arrowPanel")
        and "x" in p and "y" in p
    ]
    if poly is not None and poly.is_valid and not poly.is_empty:
        for p in symbols:
            pt = Point(float(p["x"]), float(p["y"]))
            if poly.contains(pt) or poly.covers(pt):
                fails.append(
                    f"geometry-qa: {p['kind']} '{p.get('id', '')}' at "
                    f"({p['x']:.1f},{p['y']:.1f}) lies inside the work-area "
                    f"hatch — symbols belong in buffer/taper, not the hatched bay"
                )

    # Pairwise AP/PV center separation. Skip OR-alternatives that share an
    # altGroup (or near-identical station) — those are meant to co-locate;
    # place_sheet_geometry picks one via arrow_panel_choice.
    for i, a in enumerate(symbols):
        for b in symbols[i + 1:]:
            if a.get("kind") == b.get("kind"):
                continue
            if a.get("altGroup") and a.get("altGroup") == b.get("altGroup"):
                continue
            try:
                if abs(float(a.get("stationFt", 0)) - float(b.get("stationFt", 0))) < 0.5:
                    continue
            except (TypeError, ValueError):
                pass
            d = Point(float(a["x"]), float(a["y"])).distance(
                Point(float(b["x"]), float(b["y"]))
            )
            if d < _SYMBOL_PAIR_MIN_SEP_FT:
                fails.append(
                    f"geometry-qa: {a['kind']} '{a.get('id', '')}' and "
                    f"{b['kind']} '{b.get('id', '')}' centers are only "
                    f"{d:.1f} ft apart (min {_SYMBOL_PAIR_MIN_SEP_FT:g} ft) — "
                    f"likely stacked cells"
                )

    # Channelizing run as a linestring: flag self-intersecting polylines
    # (folded taper). Skip short runs.
    by_run: dict[str, list[tuple[float, float, float]]] = {}
    for p in (channelizing_primitives or []):
        if p.get("kind") != "cone":
            continue
        by_run.setdefault(str(p.get("run")), []).append(
            (float(p["stationFt"]), float(p["x"]), float(p["y"]))
        )
    for run_id, pts in by_run.items():
        if len(pts) < 4:
            continue
        pts.sort(key=lambda t: t[0])
        coords = [(x, y) for _, x, y in pts]
        try:
            ls = LineString(coords)
        except Exception:
            continue
        if not ls.is_simple:
            fails.append(
                f"geometry-qa: channelizing run '{run_id}' polyline "
                f"self-intersects ({len(coords)} cones)"
            )

    return fails
