"""
WZTC op wrappers over WZTCBridge.bas — the framework-agnostic core shared
by mcp-server/server.py (MCP tools, for Claude Code and any other MCP
client) and mcp-server/chat_driver.py (M7's in-MicroStation chat panel
agent loop, via the Anthropic tool_runner). Extracted from server.py
mechanically — same op names, same params, same docstrings; only the
`@mcp.tool()` decorator and the direct `bridge` import were removed.

Each caller must call set_bridge() once at startup before calling anything
here, naming which Bridge instance (bridge_client.bridge for the stdio MCP
server, bridge_client.chat_bridge for the chat driver) to route through —
see bridge_client.py's docstring for why these must stay separate (two
processes racing on the same request.tsv/response.tsv otherwise).

Engineering-judgment boundary (CLAUDE.md, and the plan's core design
rule): this module never computes a spacing, taper length, or sign size
itself. compute_spacing / get_sheet_requirements wrap WZTCRules.bas /
WZTCSheetRegistry.bas so those numbers stay deterministic and PE-auditable.
The calling agent decides *what* to place and *how to respond to a site
condition* (an obstruction, a driveway); it must never invent a number
that belongs in one of those two tools.
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
import json

import plan_workflow
import placement_registry
import sheet_sandbox
import sheet_scorecard
import sheet_spec
import view_capture

_bridge = None

_BRIDGE_DIR = Path(__file__).resolve().parent.parent / "Bridge"
SHEET_PLAN_PATH = _BRIDGE_DIR / "sheet-plan.json"

# list_registry_commands' hard cap on rows returned -- see that function's
# docstring. Data/command-registry.tsv has ~1800 rows (~1600
# verified-headless-safe); an unfiltered call was measured live costing
# ~240K input tokens (~$0.75) on a single turn once an agent actually
# started using this tool for view/settings control (2026-08-02).
MAX_LISTED_ROWS = 40

# Spatial-query caps (2026-08-02): classify_site_features at radius=2000
# returned 325 rows (~93K chars) in one live turn while hunting for a
# single named sign the engineer had already offered to click. Same
# footgun class as the unfiltered registry list -- hard-cap what the
# model sees; ask for a click or a tighter radius instead of dumping
# the neighborhood.
MAX_SPATIAL_ROWS = 40
MAX_JOURNAL_LINES = 40


def _cap_spatial_rows(rows: list, tool_name: str, radius: float) -> list:
    """Keep the nearest MAX_SPATIAL_ROWS (by distanceFt when present) and
    append a truncation note so the model knows to narrow radius / ask
    the engineer to click rather than re-querying bigger."""
    total = len(rows)
    if total <= MAX_SPATIAL_ROWS:
        return rows

    def _dist(row) -> float:
        try:
            return float(row.get("distanceFt", row.get("distance", 1e30)))
        except (TypeError, ValueError):
            return 1e30

    kept = sorted(rows, key=_dist)[:MAX_SPATIAL_ROWS]
    kept.append({
        "note": (
            f"{total} elements matched within {radius} ft -- showing nearest "
            f"{MAX_SPATIAL_ROWS}. Narrow radius, add type_filter, or use "
            f"ask_user_choice(allow_point_pick=True) if the engineer can "
            f"click the target rather than paging through a dump."
        ),
        "tool": tool_name,
        "matchedTotal": total,
    })
    return kept


def set_bridge(bridge) -> None:
    """Must be called once before any function below — e.g.
    `wztc_ops.set_bridge(bridge_client.bridge)` in server.py,
    `wztc_ops.set_bridge(bridge_client.chat_bridge)` in chat_driver.py."""
    global _bridge
    _bridge = bridge


def _ok_or_raise(resp: dict, context: str) -> dict:
    if resp["status"] == "ERROR":
        raise RuntimeError(f"{context} failed: {resp.get('note', resp)}")
    return resp


# ================================================================ Query

def find_elements_near(x: float, y: float, radius: float, type_filter: str = "",
                       force: bool = False) -> list[dict]:
    """Find drawn elements within radius (ft) of (x, y) in the active model.
    type_filter narrows by kind (e.g. 'CELL'); empty string matches all
    types. Returns candidates with distance and range (nearest first when
    truncated) — matching is by bounding-box center, so a point near the
    end of a long line matches its midpoint, and multiple close candidates
    are a real ambiguity signal, not noise to collapse to one answer.

    Keep radius tight (tens of feet, not thousands). Results are hard-capped
    at MAX_SPATIAL_ROWS nearest matches — a wide radius that would return
    hundreds of elements will not give you a complete dump. If the engineer
    can point at the target, prefer ask_user_choice(allow_point_pick=True)
    over a fishing expedition.

    If you ALREADY have an elementId (from an element pick or a prior
    result), do NOT search for it with this tool — call
    get_elements_range([id]) instead. Long lines/arcs often sit far from
    where you are looking, and a wide search will miss them under the cap.

    Mid sheet-plan session: wide radius or repeated calls are refused
    (live 2026-08-04 burned MAX_TOOL_ITERATIONS fishing). Prefer
    view_drawing once then FINAL after place_sheet_geometry. force=True
    to override."""
    if _PLAN_SESSION.order_table_built and not force:
        _PLAN_SESSION.find_near_calls += 1
        if radius > 250.0:
            raise ValueError(
                f"find_elements_near radius={radius:g} ft is too wide during a "
                f"sheet plan — keep ≤250 ft or pass force=True. For visual QA "
                f"after place_sheet_geometry use view_drawing "
                f"(not capture_view — that is MCP-only, not a chat tool) "
                f"once, then FINAL; do not fish the model."
            )
        if _PLAN_SESSION.find_near_calls > 6:
            raise ValueError(
                f"find_elements_near called {_PLAN_SESSION.find_near_calls} times "
                f"this plan session — stop fishing. Use view_drawing once for "
                f"visual QA, then FINAL (or pass force=True for a real needed lookup)."
            )
        if _PLAN_SESSION.sheet_geometry_placed and _PLAN_SESSION.find_near_calls > 2:
            raise ValueError(
                "place_sheet_geometry already succeeded this session — do not "
                "keep probing with find_elements_near. Call view_drawing once "
                "(chat tool — do not call capture_view), note any defect, "
                "fix if critical, then FINAL. "
                "Pass force=True only for a targeted delete/fix."
            )
    resp = _ok_or_raise(
        _bridge.call("FIND_ELEMENTS_NEAR", x=x, y=y, radius=radius, typeFilter=type_filter),
        "find_elements_near")
    return _cap_spatial_rows(resp.get("rows", []), "find_elements_near", radius)


def get_elements_range(element_ids) -> dict:
    """Return the combined axis-aligned bbox of one or more element IDs.

    Prefer this whenever you already have elementId(s) — e.g. from
    ask_user_choice(allow_element_pick=True) — instead of find_elements_near
    fishing. Returns lowX/lowY/highX/highY (and centerX/centerY/width/
    height for convenience). Pass a list of ids or a comma-separated
    string. Errors if none of the ids are found in the active model."""
    if isinstance(element_ids, (list, tuple)):
        ids_csv = ",".join(str(i).strip() for i in element_ids if str(i).strip())
    else:
        ids_csv = str(element_ids or "").strip()
    if not ids_csv:
        return {"status": "ERROR", "note": "get_elements_range needs at least one elementId"}
    resp = _ok_or_raise(
        _bridge.call("GET_ELEMENTS_RANGE", elementIds=ids_csv),
        "get_elements_range")
    low_x, low_y = float(resp["lowX"]), float(resp["lowY"])
    high_x, high_y = float(resp["highX"]), float(resp["highY"])
    return {
        "status": "OK",
        "lowX": low_x, "lowY": low_y, "highX": high_x, "highY": high_y,
        "centerX": (low_x + high_x) / 2.0, "centerY": (low_y + high_y) / 2.0,
        "width": max(high_x - low_x, 0.0), "height": max(high_y - low_y, 0.0),
        "elementIds": ids_csv,
    }


def get_elements_in_range_box(low_x: float, low_y: float,
                              high_x: float, high_y: float,
                              max_rows: int = 1500) -> dict:
    """Graphical elements whose Range intersects the world AABB.

    Not center-in-box (unlike find_elements_near). AABB is a prefilter.
    Use check_build_overlap for a verdict."""
    resp = _ok_or_raise(
        _bridge.call(
            "GET_ELEMENTS_IN_RANGE_BOX",
            lowX=low_x, lowY=low_y, highX=high_x, highY=high_y,
            maxRows=int(max_rows or 1500)),
        "get_elements_in_range_box")
    rows = []
    truncated = False
    for r in resp.get("rows") or []:
        if str(r.get("elementId") or "") == "TRUNCATED":
            truncated = True
            continue
        rows.append(r)
    return {"status": "OK", "rows": rows, "truncated": truncated, "count": len(rows)}


def _overlap_session_kwargs() -> dict:
    di = _PLAN_SESSION.designer_inputs
    sheet = di.sheet_num if di else ""
    path = list(_PLAN_SESSION.corridor_path or [])
    if not path:
        path = list((_PLAN_SESSION.work_bay_vertices or []) or [])
    origin = path[0] if path else [0.0, 0.0]
    half = float(_PLAN_SESSION.lateral_half_len or 40.0)
    return {
        "sheet_num": sheet,
        "origin": origin[:2],
        "path_vertices": path,
        "lateral_half_width": half,
    }


def _model_rows_for_path(path: list) -> list[dict]:
    import build_overlap as ov
    bbox = ov.corridor_bbox(path, pad=80.0)
    if not bbox:
        return []
    try:
        got = get_elements_in_range_box(
            bbox["lowX"], bbox["lowY"], bbox["highX"], bbox["highY"])
        return list(got.get("rows") or [])
    except Exception:
        return []


def check_build_overlap(
        sheet_num: str = "",
        origin: list | None = None,
        path_vertices: list | None = None,
        lateral_half_width: float = 0.0,
        sta0: float = 0.0,
        sta1: float = 0.0,
        scan_model: bool = True) -> dict:
    """Caution-not-block overlap check. One tool — do not compose find_near.

    Verdicts: ok | rebuild_same_origin | collision_same_sheet |
    collision_other_sheet | stacked_duplicates. blocking is always False.
    """
    import build_overlap as ov
    import placement_registry as preg
    kw = _overlap_session_kwargs()
    sn = (sheet_num or kw["sheet_num"] or "").strip()
    path = list(path_vertices or kw["path_vertices"] or [])
    orig = origin or kw["origin"]
    half = float(lateral_half_width or kw["lateral_half_width"] or 40.0)
    model_rows = _model_rows_for_path(path) if scan_model and path else []
    ignore = set()
    for r in preg.resolve_latest_placements(sheet_num=sn):
        for eid in r.get("elementIds") or []:
            ignore.add(str(eid))
    caution = ov.classify(
        sheet_num=sn, origin=orig, path_vertices=path,
        lateral_half_width=half, sta0=sta0, sta1=sta1,
        model_rows=model_rows, ignore_ids=ignore)
    return {"status": "OK", "sheetNum": sn, "overlapCaution": caution}


def get_element_vertices(element_id: str) -> dict:
    """Densified vertices for a line, line-string, arc, or complex chain.

    Use after ask_user_choice(allow_element_pick=True). Returns path_vertices
    shape [[x,y,z],…] (capped at 80). Bounding-box get_elements_range is
    not a substitute."""
    import corridor_path as cp
    eid = str(element_id or "").strip()
    if not eid:
        return {"status": "ERROR", "note": "element_id required"}
    resp = _ok_or_raise(
        _bridge.call("GET_ELEMENT_VERTICES", elementId=eid),
        "get_element_vertices")
    rows = resp.get("rows") or []
    raw = []
    for r in rows:
        raw.append([float(r["x"]), float(r["y"]), float(r.get("z") or 0)])
    verts = cp.downsample_polyline(raw)
    length = cp.polyline_length(verts)
    return {
        "status": "OK",
        "elementId": eid,
        "vertexCount": len(verts),
        "lengthFt": round(length, 3),
        "path_vertices": verts,
        "start": verts[0] if verts else None,
        "end": verts[-1] if verts else None,
        "note": "Pass path_vertices to lock_corridor_path (source=element).",
    }


def station_to_point(align_idx: int, sta: float) -> dict:
    """Resolve a station (ft) along a committed alignment to an (x, y, z)
    point and tangent direction (tanX, tanY). The alignment must already be
    drawn and committed via AlignDraw in the current MicroStation session —
    a graceful error here means that hasn't happened yet, not a bug.

    ALWAYS check the returned `clamped` field. If sta exceeds the alignment's
    actual drawn length (totalPathLenFt), the underlying geometry engine
    clamps to the nearest end rather than failing — clamped='True' means the
    returned point is NOT really at the requested station, it's just the end
    of what was physically drawn. Confirmed live: this returns OK with no
    other indication anything is wrong, so treat clamped='True' as a reason
    to stop and tell the engineer the alignment needs to be drawn further,
    not a placement to proceed with."""
    return _ok_or_raise(_bridge.call("STATION_TO_POINT", alignIdx=align_idx, sta=sta), "station_to_point")


def get_alignment_stationing(align_idx: int) -> list[dict]:
    """Return the full stationing breakdown for a committed alignment."""
    resp = _ok_or_raise(_bridge.call("GET_ALIGNMENT_STATIONING", alignIdx=align_idx), "get_alignment_stationing")
    return resp.get("rows", [])


def get_alignment_vertices(align_idx: int) -> list[dict]:
    """Return a committed alignment's raw path segments (straight or arc)
    in design-file master units, one row per segment in path order:
    segIndex, isArc ('Y'/'N'), sx/sy/sz, ex/ey/ez, segLen, and for arcs
    cx/cy/radius/startAngle/sweepAngle (0 for straight segments).

    Fetch this ONCE per alignment and do station->XY interpolation locally
    (see mcp-server/alignment_geometry.py) instead of calling
    station_to_point once per point — that's one bridge round trip for a
    whole sheet's worth of stations instead of one per point. Placement-plan
    compiler Stage 1 (see Data/sheet-specs/STATUS.md)."""
    resp = _ok_or_raise(_bridge.call("GET_ALIGNMENT_VERTICES", alignIdx=align_idx), "get_alignment_vertices")
    return resp.get("rows", [])


_LEVEL_ALIASES_PATH = Path(__file__).resolve().parent.parent / "Data" / "level-aliases.tsv"
_LEVEL_CATEGORIES_PATH = Path(__file__).resolve().parent.parent / "Data" / "level-categories.tsv"
_LEVEL_ALIASES_CACHE: dict[str, tuple[str, ...]] | None = None
_LEVEL_CATEGORIES_CACHE: dict[str, str] | None = None


def _load_level_aliases() -> dict[str, tuple[str, ...]]:
    """alias (lower) → OR-needles for list_levels expansion. Missing/empty
    file → {}. See Data/level-aliases.tsv (feature-specific terms only)."""
    global _LEVEL_ALIASES_CACHE
    if _LEVEL_ALIASES_CACHE is not None:
        return _LEVEL_ALIASES_CACHE
    out: dict[str, tuple[str, ...]] = {}
    try:
        text = _LEVEL_ALIASES_PATH.read_text(encoding="utf-8")
    except OSError:
        _LEVEL_ALIASES_CACHE = out
        return out
    for line in text.splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        if "\t" not in raw:
            continue
        alias, needles = raw.split("\t", 1)
        alias_key = " ".join(alias.strip().lower().split())
        parts = tuple(p.strip() for p in needles.split("|") if p.strip())
        if alias_key and parts:
            out[alias_key] = parts
    _LEVEL_ALIASES_CACHE = out
    return out


def _load_level_categories() -> dict[str, str]:
    """English discipline → HDM Exhibit 20-5 category letter (A–X, plus V/Z).
    See Data/level-categories.tsv."""
    global _LEVEL_CATEGORIES_CACHE
    if _LEVEL_CATEGORIES_CACHE is not None:
        return _LEVEL_CATEGORIES_CACHE
    out: dict[str, str] = {}
    try:
        text = _LEVEL_CATEGORIES_PATH.read_text(encoding="utf-8")
    except OSError:
        _LEVEL_CATEGORIES_CACHE = out
        return out
    for line in text.splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        if "\t" not in raw:
            continue
        alias, letter = raw.split("\t", 1)
        alias_key = " ".join(alias.strip().lower().split())
        letter = letter.strip().upper()
        if alias_key and len(letter) == 1 and letter.isalpha():
            out[alias_key] = letter
    _LEVEL_CATEGORIES_CACHE = out
    return out


def _level_category_letter(name: str) -> str | None:
    """HDM first-character category for a coded level name, or None.

    Matches all-caps feature codes (DCB_P) and Letter_English styles
    (O_Details_…). Skips Draft_* / Default* English layers that happen
    to start with a category letter.
    """
    if not name:
        return None
    upper = name.upper()
    if upper.startswith("DRAFT") or upper.startswith("DEFAULT"):
        return None
    first = name[0].upper()
    if not first.isalpha():
        return None
    token0 = name.split("_", 1)[0]
    # All-caps feature code token (DCB, TWZCD, DSSD, …)
    if token0.isupper() and len(token0) >= 2 and token0.isalnum():
        return first
    # Discipline_English (O_Details_…, U_Electric_…)
    if len(name) >= 2 and name[1] == "_":
        return first
    return None


def _level_search_needles(name_contains: str) -> tuple[list[str], str | None]:
    """Return (needles, alias_hit). Always includes the raw query as a
    needle; if it matches a feature alias key (or is a word in a multi-word
    key), also OR in that alias's prefixes. Exact alias key wins over
    substring of a longer key."""
    raw = (name_contains or "").strip()
    if not raw:
        return [], None
    needles = [raw]
    aliases = _load_level_aliases()
    key = " ".join(raw.lower().split())
    hit = None
    if key in aliases:
        hit = key
        for n in aliases[key]:
            if n not in needles:
                needles.append(n)
    else:
        # Single-token query matching a word inside a multi-word alias
        # (len>=4 to avoid "of" → "right of way").
        for alias_key, parts in aliases.items():
            words = alias_key.split()
            if key == alias_key or (len(key) >= 4 and key in words):
                hit = alias_key
                for n in parts:
                    if n not in needles:
                        needles.append(n)
                break
    return needles, hit


def _level_search_category(name_contains: str) -> tuple[str | None, str | None]:
    """Return (category_letter, category_alias_key) if the query maps to an
    HDM discipline letter via Data/level-categories.tsv."""
    key = " ".join((name_contains or "").strip().lower().split())
    if not key:
        return None, None
    cats = _load_level_categories()
    if key in cats:
        return cats[key], key
    for alias_key, letter in cats.items():
        words = alias_key.split()
        if key == alias_key or (len(key) >= 4 and key in words):
            return letter, alias_key
    return None, None


def _level_prefix_histogram(names: list[str], limit: int = 25) -> str:
    from collections import Counter
    counts = Counter(n.split("_", 1)[0].upper() for n in names if n)
    return ", ".join(f"{k}({v})" for k, v in counts.most_common(limit))


def list_levels(name_contains: str = "") -> list[dict]:
    """List levels in the active design file matching name_contains
    (case-insensitive substring, e.g. 'TWZ', 'Traffic', 'SF_P').

    Matching (OR):
      1. Literal / feature-alias needles (Data/level-aliases.tsv) against
         the level name.
      2. HDM category letter (Data/level-categories.tsv) — e.g. 'drainage'
         matches every coded D* level in the file, not a hand-picked subset.

    name_contains is REQUIRED — this file can have thousands of levels
    (measured live at 3046); an unfiltered dump costs real tokens and
    still won't surface the level you want if it isn't in the first page.
    Results are hard-capped at MAX_LISTED_ROWS matches. Returns name,
    number, isDisplayed; may include matchedVia notes."""
    needle = (name_contains or "").strip()
    if not needle:
        return [{
            "status": "ERROR",
            "note": "list_levels requires name_contains (e.g. 'TWZ', 'SFB', "
                    "'Traffic', or 'drainage'). Refusing unfiltered listing — "
                    "this DGN can have thousands of levels.",
        }]
    needles, alias_hit = _level_search_needles(needle)
    cat_letter, cat_alias = _level_search_category(needle)
    resp = _ok_or_raise(_bridge.call("LIST_LEVELS"), "list_levels")
    rows = resp.get("rows", [])
    upper_needles = [n.upper() for n in needles]

    def _needle_hit(name: str) -> str | None:
        u = name.upper()
        for n, orig in zip(upper_needles, needles):
            if n in u:
                return orig
        return None

    matched: list[dict] = []
    matched_names: list[str] = []
    for r in rows:
        name = str(r.get("name", ""))
        via = _needle_hit(name)
        if via is None and cat_letter:
            if _level_category_letter(name) == cat_letter:
                via = f"category:{cat_letter}"
        if via is None:
            continue
        row = dict(r)
        if via.upper() != needle.upper():
            row["matchedVia"] = via
        matched.append(row)
        matched_names.append(name)

    total = len(matched)
    if total > MAX_LISTED_ROWS:
        matched = matched[:MAX_LISTED_ROWS]
        note = (
            f"{total} levels matched name_contains={name_contains!r}"
        )
        if cat_letter:
            note += f" (HDM category {cat_letter}"
            if cat_alias:
                note += f" via {cat_alias!r}"
            note += ")"
        if len(needles) > 1:
            note += f" (needles={needles})"
        note += (
            f" — showing first {MAX_LISTED_ROWS}. "
            f"Prefixes: {_level_prefix_histogram(matched_names)}. "
            f"Tighten with a feature prefix (e.g. list_levels('DCB'))."
        )
        matched.append({"note": note})
    if total == 0:
        bits = [f"No levels matched {name_contains!r}"]
        if alias_hit:
            bits.append(
                f"feature alias {alias_hit!r}→"
                f"{list(_load_level_aliases().get(alias_hit, ()))}"
            )
        if cat_alias and cat_letter:
            bits.append(f"category {cat_alias!r}→{cat_letter}*")
        bits.append("Ask the engineer for the project prefix, then retry.")
        matched.append({"note": " — ".join(bits)})
    else:
        meta: dict = {}
        if cat_alias and cat_letter:
            meta["categoryExpanded"] = cat_alias
            meta["categoryLetter"] = cat_letter
        if alias_hit and any(
            str(r.get("matchedVia", "")).upper() != f"CATEGORY:{cat_letter}"
            for r in matched
            if isinstance(r, dict) and r.get("matchedVia")
        ):
            meta["aliasExpanded"] = alias_hit
            meta["needles"] = needles
        if meta:
            note_parts = []
            if "categoryLetter" in meta:
                note_parts.append(
                    f"HDM category {meta['categoryLetter']}* via "
                    f"Data/level-categories.tsv ({cat_alias})."
                )
            if "aliasExpanded" in meta:
                note_parts.append(
                    f"Feature alias {alias_hit} → {needles[1:]} "
                    f"(Data/level-aliases.tsv)."
                )
            meta["note"] = " ".join(note_parts)
            matched.append(meta)
    return matched


# Common color-name → RGB for resolve_color. Tables don't store names —
# only indices + RGB — so named requests go through FindClosestColor.
_COLOR_NAME_RGB = {
    "white": (255, 255, 255),
    "black": (0, 0, 0),
    "red": (255, 0, 0),
    "green": (0, 255, 0),
    "blue": (0, 0, 255),
    "yellow": (255, 255, 0),
    "cyan": (0, 255, 255),
    "magenta": (255, 0, 255),
    "orange": (255, 165, 0),
    "purple": (128, 0, 128),
    "violet": (238, 130, 238),
    "brown": (139, 69, 19),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "pink": (255, 192, 203),
    "lime": (0, 255, 0),
    "navy": (0, 0, 128),
    "teal": (0, 128, 128),
    "maroon": (128, 0, 0),
    "olive": (128, 128, 0),
    "coral": (255, 127, 80),
    "gold": (255, 215, 0),
}


def _pack_rgb(r: int, g: int, b: int) -> int:
    """MicroStation color-table Long packing (KB0039791): R + G*256 + B*65536."""
    return (int(r) & 255) + ((int(g) & 255) * 256) + ((int(b) & 255) * 65536)


def _unpack_rgb(packed: int) -> tuple[int, int, int]:
    packed = int(packed) & 0xFFFFFF
    return packed & 255, (packed // 256) & 255, (packed // 65536) & 255


def _active_color_table():
    """Live ColorTable for the open DGN via COM — not WZTCBridge.
    Confirmed ExtractColorTable / GetColors / FindClosestColor work from
    Python against this install (2026-08-02). Kept off the VBA bridge after
    a ColorTable-typed WZTCQuery hot-reload failed to compile and wedged
    the whole bridge until reverted."""
    import pythoncom
    import ms_connect
    pythoncom.CoInitialize()
    return ms_connect.get_microstation_app().ActiveDesignFile.ExtractColorTable()


def list_colors() -> list[dict]:
    """Return every entry in the active DGN's color table (index + RGB).
    ~255 rows — small enough to return in full. Color indices are
    file-specific; never assume index 3 is orange. Prefer resolve_color
    when the engineer names a color."""
    tbl = _active_color_table()
    cols = tbl.GetColors()
    rows = []
    for i, packed in enumerate(cols):
        r, g, b = _unpack_rgb(packed)
        rows.append({"index": str(i), "red": str(r), "green": str(g), "blue": str(b)})
    return rows


def resolve_color(name: str = "", red: int | None = None,
                  green: int | None = None, blue: int | None = None) -> dict:
    """Map a named color or RGB triple to the closest index in THIS DGN's
    color table (ColorTable.FindClosestColor). Always call this before
    change_element_symbology when the engineer asks for a color by name
    (e.g. 'orange') — never guess an index. Pass name='orange' OR
    red/green/blue. Returns index plus the table's actual RGB and the
    RGB that was requested."""
    r, g, b = red, green, blue
    key = (name or "").strip().lower()
    if key:
        if key not in _COLOR_NAME_RGB:
            known = ", ".join(sorted(_COLOR_NAME_RGB))
            return {
                "status": "ERROR",
                "note": f"unknown color name {name!r}. Known names: {known}. "
                        "Or pass red/green/blue directly.",
            }
        r, g, b = _COLOR_NAME_RGB[key]
    if r is None or g is None or b is None:
        return {
            "status": "ERROR",
            "note": "resolve_color needs name= (e.g. 'orange') or red/green/blue.",
        }
    tbl = _active_color_table()
    idx = int(tbl.FindClosestColor(_pack_rgb(r, g, b)))
    ar, ag, ab = _unpack_rgb(tbl.GetColorAtIndex(idx))
    return {
        "status": "OK",
        "index": str(idx),
        "red": str(ar),
        "green": str(ag),
        "blue": str(ab),
        "requestedRed": str(int(r)),
        "requestedGreen": str(int(g)),
        "requestedBlue": str(int(b)),
        "name": key or "",
    }


# Default WZTC symbol library — same path as WZTCBridge.ExecPlaceCell /
# CellPlacer.bas. attach_cell_library() defaults here when lib_path is empty.
DEFAULT_WZTC_CELL_LIB = r"c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
# Folder of NYSDOT plan .cel libraries (utility, roadway, striping, WZTC, …).
DEFAULT_CELL_LIB_DIR = r"c:\pwworking\usny\d0119091"

# Common engineer aliases → exact LineStyles Name keys in a typical NYSDOT
# seed. resolve_line_style also does case-insensitive substring match.
_LINE_STYLE_ALIASES = {
    "solid": "0",
    "continuous": "0",
    "0": "0",
    "bylevel": "STYLE_ByLevel",
    "by level": "STYLE_ByLevel",
    "dashed": "( Dashed )",
    "dash": "( Dashed )",
    "center": "( Center )",
    "hidden": "( Hidden )",
    "dot": "( Dot )",
    "dashdot": "( Dashdot )",
    "phantom": "( Phantom )",
    "border": "( Border )",
    "divide": "( Divide )",
}


def _ms_app():
    """Live MicroStation Application via COM — same pattern as color-table
    helpers. Kept off the VBA bridge so list/resolve tools don't need a
    hot-reload when Discovery APIs change."""
    import pythoncom
    import ms_connect
    pythoncom.CoInitialize()
    return ms_connect.get_microstation_app()


def _iter_line_styles(df):
    """Yield (1-based collectionIndex, name, number) for each LineStyle.
    Collection is 1-based; Name is the stable lookup key (Number is NOT —
    LineStyles(-104) fails; LineStyles('( Dashed )') works)."""
    styles = df.LineStyles
    count = int(styles.Count)
    for i in range(1, count + 1):
        try:
            sty = styles(i)
        except Exception:
            continue
        yield i, str(getattr(sty, "Name", "") or ""), int(getattr(sty, "Number", 0))


def _com_item_by_index(coll, i0: int):
    """Return the item at Python 0-based index i0 from a COM collection.
    Fonts/TextStyles are inconsistently 0- vs 1-based depending on DGN state
    -- try i0 first, then the 1-based fallback; None if neither works."""
    try:
        return coll(i0)
    except Exception:
        pass
    try:
        return coll(i0 + 1)
    except Exception:
        return None


def _resolve_name_hits(raw: str, hits: list, name_of, describe, kind: str) -> dict:
    """Shared exact/unique/ambiguous decision for resolve_line_style,
    resolve_font, and resolve_text_style once each has already gathered its
    own list of substring-matched candidates (the exact-lookup fast path and
    per-type field extraction differ, so only this part is common). `hits`
    is a list of opaque per-type items; name_of(item) extracts the display
    name; describe(item, matched_via) builds the result dict -- pass
    matched_via=None to build the slimmer per-candidate entry used in the
    ambiguous-match error."""
    upper = raw.upper()
    exact = [h for h in hits if name_of(h).upper() == upper]
    if len(exact) == 1:
        return describe(exact[0], "case-insensitive")
    if len(hits) == 1:
        return describe(hits[0], "substring")
    if not hits:
        return {"status": "ERROR", "note": f"no {kind} matched {raw!r}."}
    sample = ", ".join(repr(name_of(h)) for h in hits[:8])
    more = f" (+{len(hits) - 8} more)" if len(hits) > 8 else ""
    return {
        "status": "ERROR",
        "note": f"{len(hits)} {kind}s matched {raw!r}: {sample}{more}. "
                "Pass a more specific name.",
        "candidates": [describe(h, None) for h in hits[:MAX_LISTED_ROWS]],
    }


def list_line_styles(name_contains: str = "") -> list[dict]:
    """List line styles in the active DGN matching name_contains
    (case-insensitive substring). name_contains is REQUIRED — this file
    can have hundreds of styles (measured live at 471). Returns name,
    number (MicroStation Number property), and collectionIndex (1-based).
    Prefer resolve_line_style when you know the style name; pass the
    returned name= to change_element_symbology(line_style_name=...)."""
    needle = (name_contains or "").strip()
    if not needle:
        return [{
            "status": "ERROR",
            "note": "list_line_styles requires name_contains (e.g. 'Dash', "
                    "'Center', 'TWZ', 'Pavt'). Refusing unfiltered listing.",
        }]
    upper = needle.upper()
    rows = []
    for idx, name, number in _iter_line_styles(_ms_app().ActiveDesignFile):
        if upper in name.upper():
            rows.append({
                "name": name,
                "number": str(number),
                "collectionIndex": str(idx),
            })
    total = len(rows)
    if total > MAX_LISTED_ROWS:
        rows = rows[:MAX_LISTED_ROWS]
        rows.append({
            "note": (
                f"{total} line styles matched name_contains={name_contains!r} "
                f"-- showing first {MAX_LISTED_ROWS}. Tighten the filter."
            )
        })
    return rows


def resolve_line_style(name: str = "") -> dict:
    """Map a line-style name (or common alias like 'dashed', 'bylevel') to
    the exact LineStyles Name key for THIS DGN. Always call before
    change_element_symbology when the engineer names a style — pass the
    returned name via line_style_name= (not the Number property; negative
    Numbers are not valid LineStyles() keys). Exact name wins; else unique
    case-insensitive substring match; else ERROR with candidates."""
    raw = (name or "").strip()
    if not raw:
        return {
            "status": "ERROR",
            "note": "resolve_line_style needs name= (e.g. 'dashed', "
                    "'( Center )', 'TWZCD_P').",
        }
    key = raw.lower()
    if key in _LINE_STYLE_ALIASES:
        alias_target = _LINE_STYLE_ALIASES[key]
        if alias_target == "STYLE_ByLevel":
            return {
                "status": "OK",
                "name": "STYLE_ByLevel",
                "number": "2147483647",
                "collectionIndex": "",
                "matchedVia": "alias",
                "note": "ByLevel is not in LineStyles() — cannot pass it to "
                        "change_element_symbology. Use run_registry_command "
                        "ACTIVE_LINESTYLE / LC=ByLevel instead.",
            }
        raw = alias_target
        key = raw.lower()

    df = _ms_app().ActiveDesignFile
    # Exact Name lookup first (fast path).
    try:
        sty = df.LineStyles(raw)
        nm = str(sty.Name)
        num = int(sty.Number)
        coll = ""
        for idx, n, _ in _iter_line_styles(df):
            if n == nm:
                coll = str(idx)
                break
        return {
            "status": "OK",
            "name": nm,
            "number": str(num),
            "collectionIndex": coll,
            "matchedVia": "exact",
        }
    except Exception:
        pass

    upper = key.upper()
    hits = [
        (idx, nm, num)
        for idx, nm, num in _iter_line_styles(df)
        if upper == nm.upper() or upper in nm.upper()
    ]

    def _describe(h, matched_via):
        idx, nm, num = h
        d = {"name": nm, "number": str(num), "collectionIndex": str(idx)}
        if matched_via is not None:
            d["status"] = "OK"
            d["matchedVia"] = matched_via
        return d

    result = _resolve_name_hits(raw, hits, lambda h: h[1], _describe, "line style")
    if result.get("status") == "ERROR" and "candidates" not in result:
        result["note"] = f"no line style matched {name!r}. Try list_line_styles(name_contains=...)."
    return result


def cell_library_status() -> dict:
    """Report whether a cell library is currently attached and its path.
    place_cell auto-attaches the WZTC library, but browse/search via
    list_cells requires an attach first — call attach_cell_library() if
    attached=False."""
    app = _ms_app()
    if not bool(app.IsCellLibraryAttached):
        return {
            "status": "OK",
            "attached": "False",
            "path": "",
            "activeCell": "",
            "note": "No cell library attached. Call attach_cell_library() "
                    f"(defaults to {DEFAULT_WZTC_CELL_LIB}) before list_cells.",
        }
    path = ""
    try:
        path = str(app.AttachedCellLibrary.FullName)
    except Exception:
        path = "(attached, path unavailable)"
    active = ""
    try:
        active = str(app.GetCExpressionValue("tcb->activeCellUtf16", "") or "")
    except Exception:
        pass
    return {
        "status": "OK",
        "attached": "True",
        "path": path,
        "activeCell": active,
    }


def attach_cell_library(lib_path: str = "") -> dict:
    """Attach a .cel cell library (Application.AttachCellLibrary). Empty
    lib_path attaches the default WZTC library (ny_plan_wztc.cel). Idempotent
    if that library is already attached. Call before list_cells when
    cell_library_status shows attached=False."""
    import os
    path = (lib_path or "").strip() or DEFAULT_WZTC_CELL_LIB
    if not os.path.isfile(path):
        return {
            "status": "ERROR",
            "note": f"cell library not found: {path}",
        }
    app = _ms_app()
    try:
        if bool(app.IsCellLibraryAttached):
            try:
                cur = str(app.AttachedCellLibrary.FullName)
                if os.path.normcase(os.path.abspath(cur)) == os.path.normcase(
                        os.path.abspath(path)):
                    return {
                        "status": "OK",
                        "attached": "True",
                        "path": cur,
                        "note": "already attached",
                    }
            except Exception:
                pass
        app.AttachCellLibrary(path)
    except Exception as e:
        return {"status": "ERROR", "note": f"AttachCellLibrary failed: {e}"}
    if not bool(app.IsCellLibraryAttached):
        return {"status": "ERROR", "note": "attach reported success but "
                "IsCellLibraryAttached is still False"}
    attached_path = path
    try:
        attached_path = str(app.AttachedCellLibrary.FullName)
    except Exception:
        pass
    return {"status": "OK", "attached": "True", "path": attached_path}


def list_cell_libraries(name_contains: str = "", lib_dir: str = "") -> dict:
    """List .cel cell libraries in the NY plan cell folder (or lib_dir).

    Default folder is DEFAULT_CELL_LIB_DIR (ny_plan_wztc, ny_plan_utility,
    ny_plan_striping, ny_plan_roadway, …). Use before find_cell / attach when
    the engineer names a theme (utility, striping, drainage) but not a path.
    """
    import os
    folder = (lib_dir or "").strip() or DEFAULT_CELL_LIB_DIR
    if not os.path.isdir(folder):
        return {
            "status": "ERROR",
            "note": f"cell library folder not found: {folder}",
            "libraries": [],
        }
    needle = (name_contains or "").strip().lower()
    libs = []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith(".cel"):
            continue
        if needle and needle not in name.lower():
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        libs.append({
            "name": name,
            "path": path,
            "sizeBytes": os.path.getsize(path),
        })
    return {
        "status": "OK",
        "libDir": folder,
        "count": len(libs),
        "libraries": libs,
        "note": (
            "Pass path= from this list to attach_cell_library / find_cell / "
            "place_cell(library_path=...). Empty name_contains lists all .cel."
        ),
    }


def find_cell(query: str, lib_dir: str = "", library_path: str = "",
              max_results: int = 25) -> dict:
    """Search cell name + description across NY plan .cel libraries.

    Use when the engineer asks to place something by plain language
    ('gas meter', 'catch basin', 'left turn arrow') and you do not know
    the exact cell name or which .cel holds it. Returns matches with
    cellName, description, libraryPath — then place_cell(cellName,
    x, y, library_path=...).

    library_path limits the search to one .cel; otherwise scans every
    .cel under lib_dir (default DEFAULT_CELL_LIB_DIR). Restores the
    previously attached library when finished.
    """
    import os
    q = (query or "").strip()
    if not q:
        return {
            "status": "ERROR",
            "note": "find_cell requires query= (e.g. 'gas meter', 'ARROW LEFT').",
            "matches": [],
        }
    try:
        max_n = max(1, min(int(max_results), 100))
    except (TypeError, ValueError):
        max_n = 25

    prior = cell_library_status()
    prior_path = ""
    if prior.get("attached") == "True":
        prior_path = str(prior.get("path") or "")

    targets: list[str] = []
    one = (library_path or "").strip()
    if one:
        if not os.path.isfile(one):
            return {
                "status": "ERROR",
                "note": f"library_path not found: {one}",
                "matches": [],
            }
        targets = [one]
    else:
        listed = list_cell_libraries(lib_dir=lib_dir)
        if listed.get("status") != "OK":
            return {
                "status": "ERROR",
                "note": listed.get("note") or "list_cell_libraries failed",
                "matches": [],
            }
        targets = [row["path"] for row in listed.get("libraries") or []]

    needle = q.upper()
    tokens = [t for t in needle.replace("-", " ").replace("_", " ").split() if t]
    matches: list[dict] = []
    libs_searched = 0
    errors: list[str] = []

    try:
        for path in targets:
            att = attach_cell_library(path)
            if att.get("status") != "OK":
                errors.append(f"{path}: {att.get('note')}")
                continue
            libs_searched += 1
            rows = list_cells(name_contains="")  # may ERROR if huge — fall back
            if (isinstance(rows, list) and rows and isinstance(rows[0], dict)
                    and rows[0].get("status") == "ERROR"):
                # Large lib: probe with each token / full needle
                probes = [needle] + [t for t in tokens if t != needle]
                seen_names: set[str] = set()
                rows = []
                for p in probes:
                    for r in list_cells(name_contains=p):
                        if not isinstance(r, dict) or not r.get("name"):
                            continue
                        if r["name"] in seen_names:
                            continue
                        seen_names.add(r["name"])
                        rows.append(r)
            for r in rows:
                if not isinstance(r, dict):
                    continue
                nm = str(r.get("name") or "")
                desc = str(r.get("description") or "")
                if not nm or nm.upper() == "DEFAULT":
                    continue
                blob = f"{nm} {desc}".upper()
                if needle in blob or (
                    tokens and all(t in blob for t in tokens)
                ):
                    matches.append({
                        "cellName": nm,
                        "description": desc,
                        "libraryPath": path,
                        "libraryName": os.path.basename(path),
                        "isPoint": r.get("isPoint"),
                        "isGraphic": r.get("isGraphic"),
                    })
                    if len(matches) >= max_n:
                        break
            if len(matches) >= max_n:
                break
    finally:
        # Restore prior attach so we don't strand the session on striping/etc.
        if prior_path and os.path.isfile(prior_path):
            try:
                attach_cell_library(prior_path)
            except Exception:
                pass
        elif prior.get("attached") != "True":
            try:
                attach_cell_library(DEFAULT_WZTC_CELL_LIB)
            except Exception:
                pass

    note = (
        f"searched {libs_searched} library(ies) for {q!r}; "
        f"{len(matches)} match(es)."
    )
    if not matches:
        note += (
            " Try a shorter query, list_cell_libraries(name_contains=…) then "
            "find_cell(query=…, library_path=…)."
        )
    elif len(matches) > 1:
        note += (
            " Multiple matches — ask_user_choice or pick the best "
            "libraryName/description, then place_cell(cellName, x, y, "
            "library_path=…)."
        )
    else:
        note += " Call place_cell(cellName, x, y, library_path=returned path)."

    return {
        "status": "OK",
        "query": q,
        "matchCount": len(matches),
        "matches": matches,
        "libsSearched": libs_searched,
        "errors": errors,
        "note": note,
    }


def list_cells(name_contains: str = "", include_shared: bool = False) -> list[dict]:
    """List cells in the currently attached cell library. name_contains
    filters name OR description (case-insensitive); optional when the
    library is small (WZTC has ~16 cells) but REQUIRED if more than
    MAX_LISTED_ROWS would be returned. Call attach_cell_library first if
    cell_library_status shows nothing attached. Returns name, description,
    isPoint, isGraphic."""
    app = _ms_app()
    if not bool(app.IsCellLibraryAttached):
        return [{
            "status": "ERROR",
            "note": "No cell library attached. Call attach_cell_library() "
                    "first (empty path = default WZTC .cel).",
        }]
    en = app.GetCellInformationEnumerator(bool(include_shared), False)
    rows = []
    needle = (name_contains or "").strip().upper()
    while en.MoveNext():
        ci = en.Current
        nm = str(getattr(ci, "Name", "") or "")
        if nm.upper() == "DEFAULT":
            continue
        desc = str(getattr(ci, "Description", "") or "")
        if needle and needle not in nm.upper() and needle not in desc.upper():
            continue
        rows.append({
            "name": nm,
            "description": desc,
            "isPoint": str(bool(getattr(ci, "IsPoint", False))),
            "isGraphic": str(bool(getattr(ci, "IsGraphic", False))),
        })
    total = len(rows)
    if not needle and total > MAX_LISTED_ROWS:
        return [{
            "status": "ERROR",
            "note": (
                f"{total} cells in attached library — pass name_contains "
                f"(e.g. 'TWZ', 'FLAG', 'Arrow') rather than listing all."
            ),
        }]
    if total > MAX_LISTED_ROWS:
        rows = rows[:MAX_LISTED_ROWS]
        rows.append({
            "note": (
                f"{total} cells matched name_contains={name_contains!r} -- "
                f"showing first {MAX_LISTED_ROWS}. Tighten the filter."
            )
        })
    return rows


def list_fonts(name_contains: str = "") -> list[dict]:
    """List fonts available in the active DGN. Optional name_contains
    (case-insensitive substring). ~24 fonts on a typical seed — small
    enough to return unfiltered. Prefer resolve_font when you know the
    name before ACTIVE FONT / place_text_label."""
    needle = (name_contains or "").strip().upper()
    fonts = _ms_app().ActiveDesignFile.Fonts
    rows = []
    seen = set()
    count = int(fonts.Count)
    for i in range(count):
        f = _com_item_by_index(fonts, i)
        if f is None:
            continue
        nm = str(getattr(f, "Name", "") or "")
        if not nm or nm in seen:
            continue
        seen.add(nm)
        if needle and needle not in nm.upper():
            continue
        rows.append({"name": nm, "id": str(int(getattr(f, "ID", 0)))})
    return rows


def resolve_font(name: str = "") -> dict:
    """Map a font name to the Fonts entry for THIS DGN (Name + ID). Exact
    name wins; else unique case-insensitive / substring match."""
    raw = (name or "").strip()
    if not raw:
        return {"status": "ERROR", "note": "resolve_font needs name= "
                "(e.g. 'Arial', 'Engineering Regular')."}
    df = _ms_app().ActiveDesignFile
    try:
        f = df.Fonts(raw)
        return {
            "status": "OK",
            "name": str(f.Name),
            "id": str(int(f.ID)),
            "matchedVia": "exact",
        }
    except Exception:
        pass
    upper = raw.upper()
    hits = []
    seen = set()
    fonts = df.Fonts
    for i in range(int(fonts.Count)):
        f = _com_item_by_index(fonts, i)
        if f is None:
            continue
        nm = str(getattr(f, "Name", "") or "")
        if not nm or nm in seen:
            continue
        seen.add(nm)
        if upper == nm.upper() or upper in nm.upper():
            hits.append((nm, int(getattr(f, "ID", 0))))

    def _describe(h, matched_via):
        d = {"name": h[0], "id": str(h[1])}
        if matched_via is not None:
            d["status"] = "OK"
            d["matchedVia"] = matched_via
        return d

    result = _resolve_name_hits(raw, hits, lambda h: h[0], _describe, "font")
    if result.get("status") == "ERROR" and "candidates" not in result:
        result["note"] = f"no font matched {name!r}. Try list_fonts()."
    return result


def list_text_styles(name_contains: str = "") -> list[dict]:
    """List text styles in the active DGN (Name, height, width, font).
    Optional name_contains. Annotation scale still comes from
    describe_drawing_state — text Height/Width here are style defaults,
    multiplied by annotation scale when placed as annotation."""
    needle = (name_contains or "").strip().upper()
    styles = _ms_app().ActiveDesignFile.TextStyles
    rows = []
    seen = set()
    for i in range(int(styles.Count)):
        ts = _com_item_by_index(styles, i)
        if ts is None:
            continue
        nm = str(getattr(ts, "Name", "") or "")
        if not nm or nm in seen:
            continue
        seen.add(nm)
        if needle and needle not in nm.upper():
            continue
        font_name = ""
        try:
            font_name = str(ts.Font.Name)
        except Exception:
            pass
        rows.append({
            "name": nm,
            "height": str(getattr(ts, "Height", "")),
            "width": str(getattr(ts, "Width", "")),
            "font": font_name,
            "id": str(int(getattr(ts, "ID", 0))),
        })
    return rows


def resolve_text_style(name: str = "") -> dict:
    """Map a text-style name to Name/height/width/font for THIS DGN.
    Call before placing text when the engineer names a style (e.g.
    'ny_Prop Normal'). Annotation scale is separate — see
    describe_drawing_state."""
    raw = (name or "").strip()
    if not raw:
        return {"status": "ERROR", "note": "resolve_text_style needs name= "
                "(e.g. 'ny_Prop Normal', 'ny_Exist Title')."}
    df = _ms_app().ActiveDesignFile
    try:
        ts = df.TextStyles(raw)
        font_name = ""
        try:
            font_name = str(ts.Font.Name)
        except Exception:
            pass
        return {
            "status": "OK",
            "name": str(ts.Name),
            "height": str(ts.Height),
            "width": str(ts.Width),
            "font": font_name,
            "id": str(int(ts.ID)),
            "matchedVia": "exact",
        }
    except Exception:
        pass
    upper = raw.upper()
    hits = []
    seen = set()
    styles = df.TextStyles
    for i in range(int(styles.Count)):
        ts = _com_item_by_index(styles, i)
        if ts is None:
            continue
        nm = str(getattr(ts, "Name", "") or "")
        if not nm or nm in seen:
            continue
        seen.add(nm)
        if upper == nm.upper() or upper in nm.upper():
            font_name = ""
            try:
                font_name = str(ts.Font.Name)
            except Exception:
                pass
            hits.append({
                "name": nm,
                "height": str(ts.Height),
                "width": str(ts.Width),
                "font": font_name,
                "id": str(int(getattr(ts, "ID", 0))),
            })

    def _describe(h, matched_via):
        d = dict(h)
        if matched_via is not None:
            d["status"] = "OK"
            d["matchedVia"] = matched_via
        return d

    result = _resolve_name_hits(raw, hits, lambda h: h["name"], _describe, "text style")
    if result.get("status") == "ERROR" and "candidates" not in result:
        result["note"] = f"no text style matched {name!r}. Try list_text_styles()."
    return result


def describe_drawing_state() -> dict:
    """Inspect the active model before making any edits: 2D/3D, master/sub
    units and resolution, annotation scale (sign-face cells are Annotation-
    class and PLACE CELL ICON applies this factor — e.g. Scale 960 when
    the factor is 960; place_sign leaves that alone so faces match text),
    active level/color/line
    style/weight, active ACS, open views (center/rotation/which is active),
    reference attachment count, current selection count, and file metadata.
    Call this at the start of a session and again whenever you're unsure
    what you're working in — never assume feet, assume scale 1:1, or assume
    nothing is selected."""
    resp = _ok_or_raise(_bridge.call("DESCRIBE_DRAWING_STATE"), "describe_drawing_state")
    return {row["key"]: row["value"] for row in resp.get("rows", [])}


def classify_site_features(x: float, y: float, radius: float) -> list[dict]:
    """Classify elements near (x, y) by matching level/cell name against
    known WZTC feature names. Site data quality is mixed by design — an
    element that doesn't match a known name/level still comes back
    (kind='unclassified') with its raw geometry rather than being dropped,
    since an unnamed obstruction is still an obstruction the agent must
    reason about.

    Keep radius tight. Results are hard-capped at MAX_SPATIAL_ROWS nearest
    matches (a 2000 ft call was measured live at 325 rows / ~$0.37 of
    follow-on input once it hit history). Do not use this to "find a named
    sign somewhere in the drawing" — if the engineer can click it, use
    ask_user_choice(allow_point_pick=True) instead."""
    resp = _ok_or_raise(_bridge.call("CLASSIFY_SITE_FEATURES", x=x, y=y, radius=radius), "classify_site_features")
    return _cap_spatial_rows(resp.get("rows", []), "classify_site_features", radius)


# =========================================================== Observation

def capture_view() -> dict:
    """Screenshot the live MicroStation window so the caller can actually
    look at the current drawing (spacing, layout, sign placement) instead
    of only reasoning from computed coordinates. OS-level capture (see
    view_capture.py) -- does NOT go through WZTCBridge/CadInputQueue, so
    it has no journal entry and can't hang MicroStation; the only failure
    mode is MicroStation not being open/visible at all. Returns
    {"path": ...} pointing at a PNG on disk; the caller (an MCP tool
    wrapper, or chat_driver.py directly) decides how to surface the actual
    image bytes."""
    return {"path": str(view_capture.capture_microstation())}


def capture_window(title_substring: str) -> dict:
    """Screenshot any visible top-level window whose title contains
    title_substring -- e.g. "WZTC Agent Chat" for the in-MicroStation chat
    panel, which is its own OS window (a modeless UserForm), separate from
    MicroStation's main frame that capture_view() targets. Same OS-level
    mechanism as capture_view; see view_capture.py. Returns {"path": ...}."""
    return {"path": str(view_capture.capture_window(title_substring))}


def adjust_view(zoom_out_percent: float = 0, pan_x: float = 0, pan_y: float = 0,
                 view_num: int = 1,
                 center_x: float | None = None, center_y: float | None = None,
                 width: float | None = None, height: float | None = None,
                 force: bool = False) -> dict:
    """Zoom and/or pan the current MicroStation view by an EXACT amount.
    This is the reliable replacement for the ZOOM_*/PAN_VIEW_* command-
    registry key-ins -- ALL of those are now needs-testing (disabled)
    because they silently activate a tool and wait for a manual datapoint
    click that never arrives when driven headlessly (confirmed live
    2026-08-02 on ZOOM_OUT, ZOOM_OUT_CENTERED, ZOOM_HALF; the rest of the
    family downgraded precautionarily, same pattern). This function sets
    View.Center/Extents directly via COM instead (view_capture.navigate_view
    -- does NOT go through WZTCBridge/CadInputQueue, so it can't hang and
    has no journal entry), which completes headlessly with no click
    needed, and unlike any registry zoom key-in it supports an exact
    percentage.

    Absolute framing (PREFERRED when you know model coords):
      center_x / center_y — absolute model-space view center. Do NOT pass
      absolute coordinates as pan_x/pan_y — those are RELATIVE deltas and
      will fling the view millions of feet away (live miss 2026-08-04).
      width / height — absolute extents in design units (ft). When set,
      zoom_out_percent is ignored.

    Relative framing (from the current view):
      zoom_out_percent: e.g. 40 zooms OUT so ~40% more area becomes visible
      (new width/height = current * 1.40). Negative zooms IN. Must be > -100.
      pan_x / pan_y: shift the CURRENT view center by this many design units.
      Positive pan_x = east/right, positive pan_y = north/up.

    Prefer focus_view_on_elements when you have elementIds. Takes ~2.5s to
    settle — call view_drawing afterward to see the result (chat agent).
    MCP clients may use capture_view instead.
    Returned width/height are what MicroStation actually applied after
    aspect-fit.

    SHEET-PLAN ONLY: after place_sheet_geometry, free pan/zoom is refused
    (use run_visual_qa_captures). General CAD / pre-compiler work is
    unaffected. force=True to override."""
    if (not force
            and not _PLAN_SESSION._qa_capture_active
            and _PLAN_SESSION.sheet_plan_active()
            and _PLAN_SESSION.sheet_geometry_placed
            and not _PLAN_SESSION.visual_qa_passed):
        plan_workflow.raise_plan_gate(
            "Free adjust_view after place_sheet_geometry is refused during a "
            "sheet build — that is how the agent burned MAX_TOOL_ITERATIONS "
            "zooming into unrelated site geometry (live 2026-08-04).",
            tool="adjust_view",
            current_step="visual_qa_passed",
            next_tool="run_visual_qa_captures",
            next_step="Call run_visual_qa_captures() for scripted "
                      "corridor/upstream/work-area/downstream shots, then FINAL. "
                      "Pass force=True only for a targeted engineer-directed pan.",
        )
    state = view_capture.get_view_state(view_num=view_num)

    if center_x is not None and center_y is not None:
        new_center_x = float(center_x) + pan_x
        new_center_y = float(center_y) + pan_y
    else:
        new_center_x = state["centerX"] + pan_x
        new_center_y = state["centerY"] + pan_y

    if width is not None and height is not None:
        new_width = max(float(width), 1.0)
        new_height = max(float(height), 1.0)
    else:
        scale = 1.0 + (zoom_out_percent / 100.0)
        if scale <= 0:
            return {"status": "ERROR",
                    "note": f"zoom_out_percent={zoom_out_percent} would produce a non-positive "
                            "scale factor -- must be greater than -100."}
        new_width = state["width"] * scale
        new_height = state["height"] * scale

    applied = view_capture.navigate_view(
        new_center_x, new_center_y, new_width, new_height,
        z=state["centerZ"], view_num=view_num)
    return {
        "status": "OK",
        "previousWidth": state["width"], "previousHeight": state["height"],
        "newWidth": applied["width"], "newHeight": applied["height"],
        "centerX": applied["centerX"], "centerY": applied["centerY"],
    }


def focus_view_on_elements(element_ids, margin: float = 1.3, view_num: int = 1,
                            min_width: float = 50.0, min_height: float = 50.0) -> dict:
    """Frame the view on the bbox of the given element ID(s).

    One-shot replacement for the find_elements_near + guess-pan dance.
    Calls get_elements_range, then adjust_view with absolute center_x/
    center_y/width/height. margin multiplies the bbox (1.3 = 30% padding).
    Degenerate (zero-area) ranges — e.g. a horizontal line — get at least
    min_width x min_height so the view is still usable."""
    rng = get_elements_range(element_ids)
    if rng.get("status") != "OK":
        return rng
    w = max(rng["width"] * margin, min_width)
    h = max(rng["height"] * margin, min_height)
    # Zero-thickness bbox (pure horizontal/vertical line): give square-ish
    # padding so the line isn't an invisible hairline fill of the view.
    if rng["width"] < 1.0:
        w = max(w, min_width)
    if rng["height"] < 1.0:
        h = max(h, min_height)
    applied = adjust_view(center_x=rng["centerX"], center_y=rng["centerY"],
                          width=w, height=h, view_num=view_num, force=True)
    applied["focusedRange"] = {
        "lowX": rng["lowX"], "lowY": rng["lowY"],
        "highX": rng["highX"], "highY": rng["highY"],
        "elementIds": rng["elementIds"],
    }
    return applied


# ============================================================== Compute

def compute_spacing(speed: int, lane_width: int, shoulder_width: str, road_type: str) -> dict:
    """Deterministic MUTCD/NYSDOT spacing lookup (WZTCRules.bas) —
    downstream taper, buffer space, merging taper, shoulder taper, advance
    warning spacing, roll-ahead distance, barrier/beam up-taper & flare.
    shoulder_width MUST be one of the existing display-label bands this
    tool's underlying table keys on (e.g. '<= 4 ft', '5-7 ft', '8 ft') — not
    an arbitrary number. road_type is 'Freeway' or 'Non-Freeway'. Always
    call this tool for these values; never estimate them yourself — they
    must stay traceable to WZTCRules.bas for a PE review, not to a model
    guess."""
    return _ok_or_raise(
        _bridge.call("COMPUTE_SPACING", speed=speed, laneWidth=lane_width,
                     shoulderWidth=shoulder_width, roadType=road_type),
        "compute_spacing")


def get_sheet_requirements(sheet_num: str) -> dict:
    """Look up required signs/elements for a 619-series standard sheet
    (e.g. '619-302') from Data/sheet-registry.tsv (all 91 DesignerRef
    sheets; some stubs have empty signs when not in the 2026 Book 3 PDF).
    Check notes for stub/catalog rows. A 'found: false' result means the
    sheet number is unknown to the registry — ask the engineer rather
    than guessing.

    When Data/sheet-specs/<sheet>.build.md exists (or sheet.buildGuide),
    attaches buildGuidePath + full buildGuide text — live tips/prefs the
    agent must follow on the next build (not agent-log only)."""
    resp = _bridge.call("GET_SHEET_REQUIREMENTS", sheetNum=sheet_num)
    if resp["status"] == "ERROR":
        return {"found": False, "note": resp.get("note", "")}
    resp["found"] = True
    guide = _attach_build_guide_fields(sheet_num, resp)
    if guide is None and sheet_spec.has_spec(sheet_num):
        resp["buildGuidePath"] = None
        resp["buildGuideNote"] = (
            f"No Data/sheet-specs/{sheet_num}.build.md yet — follow the "
            "JSON spec + prompts; add a .build.md when live tips accumulate."
        )
    if sheet_spec.has_spec(sheet_num):
        caution = _highway_caution_for_sheet(sheet_num)
        resp["highwayKinds"] = caution.get("highwayKinds")
        resp["highwayRoadway"] = caution.get("roadway")
        resp["highwayCaution"] = caution
    return resp


def get_required_designer_inputs(sheet_num: str = "") -> dict:
    """Table-driven designer-input ask list from Data/sheet-specs/<sheet>.json.

    Call this BEFORE ask_user_choice for a named 619 sheet. Returns toAsk
    (with ask_user_choice option payloads from allowed[]), derived fields
    the spec already determines (cite them; do not re-ask), and locked
    fields already in this session. Do not invent speed/area_type. Do not
    offer values outside allowed[] (e.g. 60 mph is not on 619-311)."""
    sn = (sheet_num or "").strip()
    locked_raw = _PLAN_SESSION.get_locked_inputs_dict()
    if not sn:
        sn = str(locked_raw.get("sheet_num") or "").strip()
    if not sn:
        return {
            "status": "ERROR",
            "found": False,
            "note": "sheet_num required (or lock a sheet first)",
        }
    spec = sheet_spec.load(sn)
    if spec is None:
        return {
            "status": "ERROR",
            "found": False,
            "sheetNum": sn,
            "note": f"No Data/sheet-specs/{sn}.json — ask from applicability or escalate.",
        }
    locked_fields = {}
    if locked_raw.get("locked"):
        for k in (
            "speed", "lane_width", "shoulder_width", "area_type",
            "exposure_condition", "closure_type", "road_type",
        ):
            if k in locked_raw:
                locked_fields[k] = locked_raw[k]
    out = sheet_spec.required_designer_inputs(spec, locked_fields)
    out["found"] = True
    out["highwayCaution"] = _highway_caution_for_sheet(sn)
    return out


def get_sheet_build_guide(sheet_num: str) -> dict:
    """Return the durable live-build playbook for a named 619 sheet.

    Companion file Data/sheet-specs/<sheet>.build.md (override via
    sheet.buildGuide in the JSON). Machine-enforced prefs stay in the
    JSON; this markdown holds tips, QA checklist, and gotchas so the
    next build reuses them. Call after get_sheet_requirements when
    buildGuide was truncated or you need a fresh copy mid-turn."""
    sn = (sheet_num or "").strip()
    if not sn:
        return {"status": "ERROR", "found": False, "note": "sheet_num required"}
    guide = sheet_spec.load_build_guide(sn)
    if guide is None:
        return {
            "status": "OK",
            "found": False,
            "sheetNum": sn,
            "hasSpec": sheet_spec.has_spec(sn),
            "note": (
                f"No build guide at Data/sheet-specs/{sn}.build.md "
                "(and no sheet.buildGuide override). Use the JSON spec "
                "+ get_sheet_requirements; author a .build.md for tips."
            ),
        }
    return {
        "status": "OK",
        "found": True,
        "sheetNum": guide["sheetNum"],
        "path": guide["path"],
        "charCount": guide["charCount"],
        "text": guide["text"],
        "hasSpec": sheet_spec.has_spec(sn),
        "note": (
            "Follow this playbook on named-sheet builds. Encode compiler "
            "prefs in the JSON; keep human tips here (not only agent-log)."
        ),
    }


def _remember_placed_road(
        *, road_type: str, lanes: int, lane_width_ft: float,
        shoulder_width_ft: float, yellow_gap_ft: float, side: str,
        verts, x1: float, y1: float, x2: float, y2: float, length: float,
) -> None:
    global _LAST_PLACED_ROAD
    import corridor_path as cp
    if verts is not None and len(verts) >= 2:
        raw = verts
    else:
        raw = [[x1, y1], [x2, y2]]
    path = cp.downsample_polyline(raw)
    if len(path) < 2:
        return
    _LAST_PLACED_ROAD = {
        "roadType": road_type,
        "lanes": int(lanes),
        "laneWidthFt": float(lane_width_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "yellowGapFt": float(yellow_gap_ft),
        "side": side,
        "lengthFt": float(length),
        "edgeRole": "first_travel_outer",
        "path_vertices": path,
    }


def _placed_highway_kind() -> str:
    last = _LAST_PLACED_ROAD or {}
    return str(last.get("roadType") or last.get("road_type") or "")


def _highway_caution_for_sheet(sheet_num: str) -> dict:
    """Wrong-highway caution for any 619 spec (not 619-311 only)."""
    spec = sheet_spec.load(sheet_num)
    if spec is None:
        return {"mismatch": False, "sheetNum": sheet_num, "note": "no spec"}
    return sheet_spec.highway_kind_match(spec, _placed_highway_kind())


def _record_sheet_build_ledger(path_vertices=None) -> dict:
    """One ledger row after geometry is down (even if visual QA fails)."""
    import build_ledger
    import build_overlap as ov
    import placement_registry as preg
    di = _PLAN_SESSION.designer_inputs
    if di is None:
        return {}
    path = list(path_vertices or _PLAN_SESSION.corridor_path or [])
    origin = path[0][:2] if path else [0.0, 0.0]
    ids = []
    for r in preg.resolve_latest_placements(sheet_num=di.sheet_num):
        ids.extend(str(x) for x in (r.get("elementIds") or []) if str(x).isdigit())
    nums = [int(x) for x in ids] if ids else []
    return build_ledger.append_build(
        sheet_num=di.sheet_num,
        origin=origin,
        path_vertices=path,
        lateral_half_width=float(_PLAN_SESSION.lateral_half_len or 40.0),
        element_id_min=str(min(nums)) if nums else "",
        element_id_max=str(max(nums)) if nums else "",
        bbox=ov.corridor_bbox(path) if path else {},
    )


def _derived_closed_side(sheet_num: str) -> Optional[str]:
    spec = sheet_spec.load(sheet_num) if sheet_num else None
    if not spec:
        return None
    c = str((spec.get("applicability") or {}).get("closure") or "").lower()
    if "right" in c:
        return "right"
    if "left" in c:
        return "left"
    return None


def _approach_for_locked_sheet() -> Optional[dict]:
    import corridor_path as cp
    di = _PLAN_SESSION.designer_inputs
    if di is None:
        return None
    spec = sheet_spec.load(di.sheet_num)
    if spec is None:
        return None
    try:
        resolved = sheet_spec.resolve(
            spec, di.speed, di.lane_width, di.shoulder_width, di.area_type)
        return cp.sheet_approach_ft(spec, resolved)
    except Exception:
        return None


def _centerline_offset_ft() -> float:
    last = _LAST_PLACED_ROAD or {}
    di = _PLAN_SESSION.designer_inputs
    lane = float(last.get("laneWidthFt") or (di.lane_width if di else 12) or 12)
    sh = float(last.get("shoulderWidthFt") or 0)
    if sh <= 0 and di is not None:
        b = str(di.shoulder_width or "")
        if ">=" in b or "8" in b:
            sh = 8.0
        elif "5" in b:
            sh = 6.0
        elif "4" in b:
            sh = 4.0
    yel = float(last.get("yellowGapFt") or 2.0)
    lanes = int(last.get("lanes") or 4)
    per_dir = max(1, lanes // 2)
    return yel / 2.0 + per_dir * lane + sh


def propose_corridor_source() -> dict:
    """Phase B ladder: which roadway to build along. Call after designer inputs."""
    import corridor_path as cp
    last = _LAST_PLACED_ROAD
    options = []
    if last and last.get("path_vertices"):
        ln = last.get("lengthFt") or cp.polyline_length(last["path_vertices"])
        options.append({
            "label": "The road I just placed (Recommended)",
            "description": (
                f"{last.get('lanes')}-lane {last.get('roadType')} "
                f"{ln:.0f} ft, {len(last['path_vertices'])} vertices. "
                "Reuse this session's striping (no click)."
            ),
            "value": "last_placed",
            "recommended": True,
        })
    options.append({
        "label": "I'll click the roadway",
        "description": (
            "One element pick. Agent calls get_element_vertices. "
            "If you click yellow centerline, say so — agent offsets to the outer edge."
        ),
        "value": "element",
    })
    options.append({
        "label": "Trace it from a level",
        "description": "find_reference_linework; confirm the longest chain.",
        "value": "level",
    })
    options.append({
        "label": "I'll click points along it",
        "description": "Last resort. Click ON the road, not 38 ft off centerline.",
        "value": "points",
    })
    return {
        "status": "OK",
        "lastPlacedAvailable": bool(last and last.get("path_vertices")),
        "askUserChoice": {
            "question": "Which roadway should I build along?",
            "options": options,
        },
        "note": (
            "Call lock_corridor_path with source=last_placed|element|level|points. "
            "Do not ask the engineer to eyeball the closed-lane left edge."
        ),
    }


def lock_corridor_path(
        source: str,
        element_id: str = "",
        vertices: list | None = None,
        reverse: bool = False,
        edge_role: str = "first_travel_outer",
        level_name_contains: str = "",
) -> dict:
    """Lock first-travel-outer path_vertices from the Phase B answer."""
    import corridor_path as cp
    src = (source or "").strip().lower()
    role = (edge_role or "first_travel_outer").strip().lower()
    pts = None
    if src == "last_placed":
        if not _LAST_PLACED_ROAD or not _LAST_PLACED_ROAD.get("path_vertices"):
            return {"status": "ERROR", "note": "No road placed this session. Pick click/level/points."}
        pts = list(_LAST_PLACED_ROAD["path_vertices"])
        role = str(_LAST_PLACED_ROAD.get("edgeRole") or "first_travel_outer")
    elif src == "element":
        got = get_element_vertices(element_id)
        if got.get("status") != "OK":
            return got
        pts = got["path_vertices"]
    elif src == "level":
        chains = find_reference_linework(level_name_contains)
        if isinstance(chains, dict) and chains.get("status") == "ERROR":
            return chains
        rows = chains if isinstance(chains, list) else []
        if not rows:
            return {"status": "ERROR", "note": "No chains on that level. Pick click or points."}
        if len(rows) > 1:
            ranked = sorted(rows, key=lambda r: float(r.get("totalLengthFt") or 0), reverse=True)
            return {
                "status": "OK",
                "needsChoice": True,
                "candidates": [
                    {
                        "chainIdx": r.get("chainIdx"),
                        "lengthFt": r.get("totalLengthFt"),
                        "vertexCount": r.get("vertexCount"),
                    }
                    for r in ranked[:4]
                ],
                "note": "Ask which chain. Then lock_corridor_path(source=points, vertices=that chain).",
            }
        tsv = rows[0].get("verticesTSV") or ""
        pts = []
        for part in tsv.split("|"):
            nums = [float(x) for x in part.split(",") if x]
            if len(nums) >= 2:
                pts.append(nums[:3])
    elif src == "points":
        pts = list(vertices or [])
    else:
        return {"status": "ERROR", "note": f"unknown source {source!r}"}

    pts = cp.downsample_polyline(pts or [])
    if len(pts) < 2:
        return {"status": "ERROR", "note": "Need >= 2 path vertices"}
    if reverse:
        pts = cp.reverse_polyline(pts)
    if role in ("centerline", "center", "yellow"):
        pts = cp.offset_polyline(pts, _centerline_offset_ft())
        role = "first_travel_outer"

    length = cp.polyline_length(pts)
    approach = _approach_for_locked_sheet()
    check = cp.length_check(length, approach, None) if approach else None
    closed = None
    di = _PLAN_SESSION.designer_inputs
    if di:
        closed = _derived_closed_side(di.sheet_num)
    _PLAN_SESSION.corridor_path = pts
    _PLAN_SESSION.corridor_meta = {
        "source": src, "edgeRole": role, "lengthFt": length, "reversed": bool(reverse),
        "closedSideDerived": closed,
    }
    _save_sheet_plan()
    out = {
        "status": "OK",
        "source": src,
        "edgeRole": role,
        "vertexCount": len(pts),
        "lengthFt": round(length, 3),
        "start": pts[0],
        "end": pts[-1],
        "path_vertices": pts,
        "closedSideDerived": closed,
        "travelAskUserChoice": {
            "question": "Travel direction through the work bay?",
            "options": cp.travel_choice_options(pts),
        },
        "lengthCheck": check,
        "note": (
            "If travel is already known, skip the travel question. "
            "Then propose_work_area_on_path. Do not ask closed_side when "
            "closedSideDerived is set. If lengthCheck.ok is false, tell the "
            "engineer the shortfall and offer to extend — do not build."
        ),
    }
    if closed:
        out["closed_side"] = closed
    di = _PLAN_SESSION.designer_inputs
    if di and di.sheet_num:
        out["highwayCaution"] = _highway_caution_for_sheet(di.sheet_num)
    return out


def propose_work_area_on_path() -> dict:
    """Phase C: work bay along the locked corridor (station only)."""
    pts = _PLAN_SESSION.corridor_path
    if not pts:
        return {"status": "ERROR", "note": "lock_corridor_path first"}
    return {
        "status": "OK",
        "askUserChoice": {
            "question": "Where is the work area, along that road?",
            "options": [
                {
                    "label": "Click the two ends (Recommended)",
                    "description": "Picks snap to the path. You only choose station.",
                    "value": "ends",
                },
                {
                    "label": "Type start station + length",
                    "description": "Exact, PE-auditable.",
                    "value": "station_length",
                },
                {
                    "label": "Click the middle + a length",
                    "description": "One click plus a number.",
                    "value": "mid_length",
                },
            ],
        },
        "note": "Then snap_work_area_to_path. Lateral comes from resolve_sheet_lateral.",
    }


def snap_work_area_to_path(
        mode: str,
        p1: list | None = None,
        p2: list | None = None,
        start_sta: float | None = None,
        length_ft: float | None = None,
        mid: list | None = None,
) -> dict:
    """Snap work-bay ends onto the locked corridor. Returns edges for resolve_sheet_lateral."""
    import corridor_path as cp
    pts = _PLAN_SESSION.corridor_path
    if not pts:
        return {"status": "ERROR", "note": "lock_corridor_path first"}
    m = (mode or "ends").strip().lower()
    path_len = cp.polyline_length(pts)
    if m == "station_length":
        if start_sta is None or length_ft is None:
            return {"status": "ERROR", "note": "start_sta and length_ft required"}
        sta_a = float(start_sta)
        sta_b = sta_a + float(length_ft)
    elif m == "mid_length":
        if mid is None or length_ft is None or len(mid) < 2:
            return {"status": "ERROR", "note": "mid [x,y] and length_ft required"}
        n = cp.nearest_station(pts, float(mid[0]), float(mid[1]))
        half = float(length_ft) / 2.0
        sta_a = n["stationFt"] - half
        sta_b = n["stationFt"] + half
    else:
        if p1 is None or p2 is None or len(p1) < 2 or len(p2) < 2:
            return {"status": "ERROR", "note": "p1 and p2 [x,y] required"}
        n1 = cp.nearest_station(pts, float(p1[0]), float(p1[1]))
        n2 = cp.nearest_station(pts, float(p2[0]), float(p2[1]))
        sta_a, sta_b = n1["stationFt"], n2["stationFt"]
        snap_dist = max(n1["distFt"], n2["distFt"])
    if sta_a > sta_b:
        sta_a, sta_b = sta_b, sta_a
    work_len = sta_b - sta_a
    if work_len < 1.0:
        return {"status": "ERROR", "note": f"work bay too short ({work_len:.2f} ft)"}
    up = cp.point_at_station(pts, sta_a)
    dn = cp.point_at_station(pts, sta_b)
    bay = cp.sample_span(pts, sta_a, sta_b)
    approach = _approach_for_locked_sheet()
    check = cp.length_check(path_len, approach, work_len) if approach else None
    _PLAN_SESSION.work_area_edges = {
        "upstream": up, "downstream": dn, "sta0": sta_a, "sta1": sta_b,
    }
    _PLAN_SESSION.work_bay_vertices = bay
    _save_sheet_plan()
    closed = (_PLAN_SESSION.corridor_meta or {}).get("closedSideDerived")
    out = {
        "status": "OK",
        "mode": m,
        "upstream_edge": up,
        "downstream_edge": dn,
        "workLenFt": round(work_len, 3),
        "upstreamStaFt": round(sta_a, 3),
        "downstreamStaFt": round(sta_b, 3),
        "path_vertices": pts,
        "work_bay_vertices": bay,
        "lengthCheck": check,
        "closed_side": closed,
        "note": (
            "Call resolve_sheet_lateral(upstream_edge, downstream_edge, "
            "closed_side=closed_side or ask, path_vertices=path_vertices, "
            "real_road_edge=True). Then run_sheet_build with the same edges "
            "and path_vertices."
        ),
    }
    if m == "ends":
        out["snapDistFt"] = round(snap_dist, 3)
    if check and not check["ok"]:
        out["status"] = "ERROR"
        out["note"] = check["note"]
    return out


def _attach_build_guide_fields(sheet_num: str, resp: dict) -> Optional[dict]:
    """Merge load_build_guide into a tool response dict. Returns guide or None."""
    guide = sheet_spec.load_build_guide(sheet_num)
    if guide is None:
        return None
    resp["buildGuidePath"] = guide["path"]
    resp["buildGuideCharCount"] = guide["charCount"]
    resp["buildGuide"] = guide["text"]
    resp["buildGuideNote"] = (
        "Durable live-build playbook — follow tips/QA/gotchas. "
        "Machine prefs are in the sheet JSON (annotationStyle, etc.). "
        "Re-fetch via get_sheet_build_guide if needed."
    )
    return guide


def resolve_sign_code(code: str) -> list[dict]:
    """Translate a raw sign code as printed on a 619 sheet (from
    get_sheet_requirements' `signs` field, e.g. 'W20-1') into
    SignLibrary.bas's zero-padded, message/side-suffixed key that
    place_sign actually needs (e.g. 'W20-01RA'). ALWAYS call this before
    place_sign for a code that came from get_sheet_requirements rather
    than one an engineer typed directly in library form already.

    Returns one row per match, each with matchType:
    - 'exact'/'normalized': a single unambiguous match — use signNumber as-is.
    - 'candidate': the base sign has multiple message/side variants (e.g.
      distance-to-work Ahead/Feet/Mile, or Road/Street, or Left/Right) —
      every row is a real possibility. Do not guess between them; pick
      based on context you already have (e.g. side of a divided highway)
      or ask the engineer which variant.
    - Empty list: the sign is not in SignLibrary.bas yet — a real content
      gap, not a typo. Say so plainly rather than inventing a CellName."""
    resp = _ok_or_raise(_bridge.call("RESOLVE_SIGN_CODE", code=code), "resolve_sign_code")
    return resp.get("rows", [])


# ================================================================== Draw
# Every draw op takes an optional `reason`. It rides through untouched to
# WZTCBridge's journal (Bridge/wztc-journal.tsv) alongside every other
# param — pass it whenever a placement isn't the default/expected one (an
# obstruction dodge, a non-standard station) so a PE reviewing the journal
# later can see *why*, not just *what*.

@dataclass
class DesignerInputs:
    """Snapshot of the designer inputs (WZTCDesigner.frm equivalent) locked
    in by the most recent successful build_wztc_order_table call. Real
    persisted state — not something later code has to re-derive by
    rereading chat history (that's what silently regressed live
    2026-08-04 into a re-asked area_type)."""
    sheet_num: str
    speed: int
    road_type: str
    lane_width: int
    shoulder_width: str
    area_type: str = ""
    closure_type: str = ""
    exposure_condition: str = ""
    protective_vehicle_gvw: int = 0


@dataclass
class PlanSession:
    """In-process plan-session state (chat_driver process lifetime).

    SCOPE: checklist / PLAN_GATE / anti-fish-after-compiler apply ONLY while
    a named 619 sheet build is active (order_table_built + designer_inputs).
    General CAD, one-offs, and non-sheet WZTC tasks stay unconstrained —
    the agent still reasons freely there. Cleared by reset() (exit_mode)
    or rebuilt when build_wztc_order_table runs.

    Soft memory so place_perp_line/place_sign/heuristic PlaceOrderTable*
    tools can refuse incomplete sheet-plan patterns (live 2026-08-02
    incomplete-sketch miss)."""
    placed_workspace: bool = False
    order_table_built: bool = False
    stations_placed_aligns: set[int] = field(default_factory=set)
    designer_inputs: Optional[DesignerInputs] = None
    locked_sign_rows: set[tuple[int, str]] = field(default_factory=set)
    locked_sign_details: list[dict] = field(default_factory=list)
    aligns_ready: set[int] = field(default_factory=set)
    # Anti-fishing counters (live 2026-08-04) — only incremented/enforced
    # while sheet plan active (see find_elements_near).
    find_near_calls: int = 0
    sheet_geometry_placed: bool = False
    # Sheet-build checklist bits (ignored when sheet_plan_active is False).
    required_aligns: set[int] = field(default_factory=set)
    signs_placed_rows: set[tuple[int, str]] = field(default_factory=set)
    sign_attrs_applied: bool = False
    geometry_qa_passed: bool = False
    visual_qa_passed: bool = False
    # Last place_order_table_stations rows per align (for run_sheet_build tips).
    last_station_rows: dict[int, list] = field(default_factory=dict)
    # True only while run_visual_qa_captures drives adjust_view internally.
    _qa_capture_active: bool = False
    # Durable plan extras (Bridge/sheet-plan.json).
    work_area_edges: Optional[dict] = None
    # Closed-lane / roadway edge polyline through the work bay (up→dn).
    # Used by compile_hatch on curved corridors; None = straight chord.
    work_bay_vertices: Optional[list] = None
    plan_updated_at: str = ""
    # Post-placement scorecard + phase-boundary replan / reflection.
    last_scorecard: Optional[dict] = None
    last_compiled: Optional[dict] = None
    last_failed_phase: str = ""
    last_replan: Optional[dict] = None
    visual_qa_failures: list[str] = field(default_factory=list)
    reflection_log: list[dict] = field(default_factory=list)
    sandbox: Optional[dict] = None
    # Closed-lane lateral (resolve_sheet_lateral). Used by run_sheet_build /
    # place_sheet_geometry when set — real-road half_len + outward_sign.
    lateral_outward_sign: Optional[float] = None
    lateral_half_len: Optional[float] = None
    closed_side: str = ""
    real_road_edge: bool = False
    closed_outward_x: float = 0.0
    closed_outward_y: float = 0.0
    opposite_half_len: Optional[float] = None
    # Locked first-travel-outer polyline for this sheet build.
    corridor_path: Optional[list] = None
    corridor_meta: Optional[dict] = None

    def reset(self) -> None:
        """Drop plan-flow memory (call from exit_mode so a later general/
        wztc task doesn't inherit a prior plan's gate state)."""
        self.placed_workspace = False
        self.order_table_built = False
        self.stations_placed_aligns = set()
        self.designer_inputs = None
        self.locked_sign_rows = set()
        self.locked_sign_details = []
        self.aligns_ready = set()
        self.find_near_calls = 0
        self.sheet_geometry_placed = False
        self.required_aligns = set()
        self.signs_placed_rows = set()
        self.sign_attrs_applied = False
        self.geometry_qa_passed = False
        self.visual_qa_passed = False
        self.last_station_rows = {}
        self._qa_capture_active = False
        self.work_area_edges = None
        self.work_bay_vertices = None
        self.plan_updated_at = ""
        self.last_scorecard = None
        self.last_compiled = None
        self.last_failed_phase = ""
        self.last_replan = None
        self.visual_qa_failures = []
        self.reflection_log = []
        self.sandbox = None
        self.lateral_outward_sign = None
        self.lateral_half_len = None
        self.closed_side = ""
        self.real_road_edge = False
        self.closed_outward_x = 0.0
        self.closed_outward_y = 0.0
        self.opposite_half_len = None
        self.corridor_path = None
        self.corridor_meta = None

    def lock_designer_inputs(self, **kwargs) -> None:
        self.designer_inputs = DesignerInputs(**kwargs)

    def get_locked_inputs_dict(self) -> dict:
        if self.designer_inputs is None:
            return {"locked": False}
        out = {"locked": True, **asdict(self.designer_inputs)}
        if self.lateral_outward_sign is not None or self.lateral_half_len is not None:
            out["lateral"] = {
                "outward_sign": self.lateral_outward_sign,
                "half_len": self.lateral_half_len,
                "opposite_half_len": self.opposite_half_len,
                "closed_side": self.closed_side or None,
                "real_road_edge": self.real_road_edge,
                "closed_outward": [self.closed_outward_x, self.closed_outward_y],
            }
        return out

    def lock_sign_rows(self, sign_rows: list[dict]) -> None:
        self.locked_sign_details = [dict(r) for r in (sign_rows or [])]
        self.locked_sign_rows = {
            (int(r["align_idx"]), str(r["sign_num"]).strip().upper())
            for r in self.locked_sign_details
        }

    def mark_align_ready(self, align_idx: int) -> bool:
        """Returns True once BOTH align 1 and align 2 are ready."""
        self.aligns_ready.add(align_idx)
        return not ({1, 2} - self.aligns_ready)

    def sheet_plan_active(self) -> bool:
        return plan_workflow.sheet_plan_active(self)


_PLAN_SESSION = PlanSession()
_LAST_PLACED_ROAD: Optional[dict] = None


def _iso_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _save_sheet_plan() -> Optional[Path]:
    """Persist PlanSession checklist to Bridge/sheet-plan.json.
    No-op (returns None) when no sheet plan is active."""
    s = _PLAN_SESSION
    if not s.sheet_plan_active():
        return None
    s.plan_updated_at = _iso_now()
    di = s.designer_inputs
    payload = {
        "schemaVersion": "1",
        "updatedAt": s.plan_updated_at,
        "sheetNum": di.sheet_num if di else "",
        "designerInputs": asdict(di) if di else None,
        "requiredAligns": sorted(s.required_aligns),
        "alignsReady": sorted(s.aligns_ready),
        "stationsPlacedAligns": sorted(s.stations_placed_aligns),
        "signsPlaced": [
            f"{a}:{c}" for a, c in sorted(s.signs_placed_rows)
        ],
        "lockedSignRows": list(s.locked_sign_details),
        "checklist": {
            "inputs_locked": di is not None,
            "order_table_built": s.order_table_built,
            "corridor_ready": bool(s.aligns_ready & {1, 2}) and not (
                (s.required_aligns or {1, 2}) - s.aligns_ready
            ),
            "stations_placed": sorted(s.stations_placed_aligns),
            "signs_placed": [
                f"{a}:{c}" for a, c in sorted(s.signs_placed_rows)
            ],
            "sign_attrs_applied": s.sign_attrs_applied,
            "compiler_placed": s.sheet_geometry_placed,
            "geometry_qa_passed": s.geometry_qa_passed,
            "visual_qa_passed": s.visual_qa_passed,
        },
        "workAreaEdges": s.work_area_edges,
        "workBayVertices": s.work_bay_vertices,
        "lateral": {
            "outward_sign": s.lateral_outward_sign,
            "half_len": s.lateral_half_len,
            "opposite_half_len": s.opposite_half_len,
            "closed_side": s.closed_side or None,
            "real_road_edge": s.real_road_edge,
            "closed_outward": [s.closed_outward_x, s.closed_outward_y],
        },
        "lastStationRowsByAlign": {
            str(k): v for k, v in s.last_station_rows.items()
        },
        "placedWorkspace": s.placed_workspace,
        "lastFailedPhase": s.last_failed_phase or None,
        "lastReplan": s.last_replan,
        "scorecardPassed": (
            None if s.last_scorecard is None
            else bool(s.last_scorecard.get("passed"))
        ),
        "visualQaFailures": list(s.visual_qa_failures or []),
        "corridorPath": s.corridor_path,
        "corridorMeta": s.corridor_meta,
        "lastPlacedRoad": (
            None if not _LAST_PLACED_ROAD else {
                k: v for k, v in _LAST_PLACED_ROAD.items() if k != "path_vertices"
            } | {"path_vertices": _LAST_PLACED_ROAD.get("path_vertices")}
        ),
    }
    SHEET_PLAN_PATH.parent.mkdir(parents=True, exist_ok=True)
    SHEET_PLAN_PATH.write_text(
        json.dumps(payload, indent=2), encoding="utf-8")
    return SHEET_PLAN_PATH


def _clear_sheet_plan_file() -> None:
    try:
        if SHEET_PLAN_PATH.exists():
            SHEET_PLAN_PATH.unlink()
    except OSError:
        pass


def _load_sheet_plan(path: Optional[Path] = None) -> dict:
    """Load Bridge/sheet-plan.json into _PLAN_SESSION. Returns status dict."""
    p = path or SHEET_PLAN_PATH
    if not p.exists():
        return {"status": "OK", "loaded": False, "note": "no sheet-plan.json"}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        return {"status": "ERROR", "loaded": False, "note": str(e)}
    if not data.get("order_table_built") and not (data.get("checklist") or {}).get(
            "order_table_built"):
        # tolerate either top-level or checklist-only
        if not data.get("designerInputs"):
            return {"status": "OK", "loaded": False, "note": "empty plan file"}

    s = _PLAN_SESSION
    s.reset()
    di = data.get("designerInputs")
    if di:
        s.lock_designer_inputs(**{
            k: di[k] for k in (
                "sheet_num", "speed", "road_type", "lane_width", "shoulder_width",
                "area_type", "closure_type", "exposure_condition",
                "protective_vehicle_gvw",
            ) if k in di
        })
    cl = data.get("checklist") or {}
    s.order_table_built = bool(
        data.get("order_table_built", cl.get("order_table_built", True)))
    s.required_aligns = set(int(x) for x in (data.get("requiredAligns") or []))
    s.aligns_ready = set(int(x) for x in (data.get("alignsReady") or []))
    s.stations_placed_aligns = set(
        int(x) for x in (data.get("stationsPlacedAligns")
                         or cl.get("stations_placed") or []))
    signs = data.get("signsPlaced") or cl.get("signs_placed") or []
    s.signs_placed_rows = set()
    for item in signs:
        if isinstance(item, str) and ":" in item:
            a, c = item.split(":", 1)
            try:
                s.signs_placed_rows.add((int(a), c.upper()))
            except ValueError:
                pass
    if data.get("lockedSignRows"):
        s.lock_sign_rows(data["lockedSignRows"])
    s.sign_attrs_applied = bool(
        data.get("sign_attrs_applied", cl.get("sign_attrs_applied", False)))
    s.sheet_geometry_placed = bool(
        data.get("sheet_geometry_placed", cl.get("compiler_placed", False)))
    s.geometry_qa_passed = bool(
        data.get("geometry_qa_passed", cl.get("geometry_qa_passed", False)))
    s.visual_qa_passed = bool(
        data.get("visual_qa_passed", cl.get("visual_qa_passed", False)))
    s.placed_workspace = bool(data.get("placedWorkspace", False))
    s.work_area_edges = data.get("workAreaEdges")
    wb = data.get("workBayVertices")
    s.work_bay_vertices = list(wb) if isinstance(wb, list) and len(wb) >= 2 else None
    cp = data.get("corridorPath")
    s.corridor_path = list(cp) if isinstance(cp, list) and len(cp) >= 2 else None
    s.corridor_meta = data.get("corridorMeta") if isinstance(data.get("corridorMeta"), dict) else None
    global _LAST_PLACED_ROAD
    lp = data.get("lastPlacedRoad")
    if isinstance(lp, dict) and isinstance(lp.get("path_vertices"), list) and len(lp["path_vertices"]) >= 2:
        _LAST_PLACED_ROAD = lp
    lat = data.get("lateral") or {}
    if lat.get("outward_sign") is not None:
        s.lateral_outward_sign = float(lat["outward_sign"])
    if lat.get("half_len") is not None:
        s.lateral_half_len = float(lat["half_len"])
    if lat.get("opposite_half_len") is not None:
        s.opposite_half_len = float(lat["opposite_half_len"])
    s.closed_side = str(lat.get("closed_side") or "")
    s.real_road_edge = bool(lat.get("real_road_edge"))
    co = lat.get("closed_outward") or [0.0, 0.0]
    if isinstance(co, (list, tuple)) and len(co) >= 2:
        s.closed_outward_x = float(co[0])
        s.closed_outward_y = float(co[1])
    s.plan_updated_at = str(data.get("updatedAt") or "")
    s.last_failed_phase = str(data.get("lastFailedPhase") or "")
    s.last_replan = data.get("lastReplan")
    s.visual_qa_failures = list(data.get("visualQaFailures") or [])
    lsr = data.get("lastStationRowsByAlign") or {}
    s.last_station_rows = {
        int(k): list(v) for k, v in lsr.items()
        if str(k).isdigit()
    }
    return {
        "status": "OK",
        "loaded": True,
        "sheetNum": (s.designer_inputs.sheet_num if s.designer_inputs else ""),
        "persistedPath": str(p),
        "updatedAt": s.plan_updated_at,
        "sheetPlanActive": s.sheet_plan_active(),
    }


def try_restore_sheet_plan() -> dict:
    """Chat-driver startup: restore incomplete plan across process restarts."""
    if not SHEET_PLAN_PATH.exists():
        return {"status": "OK", "loaded": False}
    try:
        data = json.loads(SHEET_PLAN_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"status": "OK", "loaded": False}
    cl = data.get("checklist") or {}
    complete = bool(cl.get("visual_qa_passed")) and bool(cl.get("compiler_placed"))
    if complete:
        return {
            "status": "OK",
            "loaded": False,
            "note": "prior sheet plan already complete — not restoring gates",
        }
    return _load_sheet_plan()


def reset_plan_session_flags() -> None:
    """Drop plan-flow memory (call from exit_mode so a later general/wztc
    task doesn't inherit a prior plan's gate state)."""
    global _LAST_PLACED_ROAD
    _PLAN_SESSION.reset()
    _LAST_PLACED_ROAD = None
    _clear_sheet_plan_file()
    placement_registry.clear_registry()


def _check_corridor_topology_if_ready(align_idx: int, force: bool) -> Optional[str]:
    """Called after commit_alignment/adopt_alignment succeeds. Once BOTH
    align 1 and align 2 are ready (and this build has a real sheet spec),
    runs check_corridor_topology immediately — not only later at
    compile_hatch/run_rules_gate time, by which point a bad corridor could
    already have a full plan computed against it. Returns a warning string
    (force=True) or raises ValueError (force=False) when the check fails;
    returns None when not yet applicable (only one alignment ready, or no
    spec locked for this build)."""
    if not _PLAN_SESSION.mark_align_ready(align_idx):
        return None
    _save_sheet_plan()
    inputs = _PLAN_SESSION.designer_inputs
    if inputs is None:
        return None
    spec = sheet_spec.load(inputs.sheet_num)
    if spec is None:
        return None
    resolved = sheet_spec.resolve(
        spec, inputs.speed, inputs.lane_width, inputs.shoulder_width,
        inputs.area_type or None, inputs.closure_type or None,
        inputs.exposure_condition or None,
        protective_vehicle_gvw=inputs.protective_vehicle_gvw or None)
    import alignment_geometry as ag
    segs1 = ag.parse_vertices(get_alignment_vertices(1))
    segs2 = ag.parse_vertices(get_alignment_vertices(2))
    width_ft = float(inputs.lane_width or 0.0)
    fails = sheet_spec.check_corridor_topology(spec, resolved, segs1, segs2, width_ft=width_ft)
    if not fails:
        return None
    msg = "; ".join(fails)
    if force:
        return f"corridorTopologyWarning (force=True, proceeding anyway): {msg}"
    raise ValueError(
        f"corridor-topology check failed after both alignments became ready: {msg} "
        f"The alignment commit/adopt itself already succeeded (nothing to undo) — "
        f"fix align2's geometry and recommit/re-adopt it, or pass force=True if this "
        f"corridor is intentional."
    )


def get_locked_designer_inputs() -> dict:
    """Returns the designer inputs (speed/road_type/lane_width/
    shoulder_width/area_type/sheet_num/...) locked in by the most recent
    successful build_wztc_order_table call this session, or
    {"locked": False} if none yet. Real persisted state, not "reread it
    from chat history" -- call this instead of re-deriving/re-asking the
    engineer for values already established earlier in the same build,
    including after a turn that hit MAX_TOOL_ITERATIONS and had to
    continue in a fresh turn (history re-reading is what silently regressed
    live 2026-08-04 into a re-asked area_type)."""
    return _PLAN_SESSION.get_locked_inputs_dict()


def get_plan_status() -> dict:
    """Sheet-build checklist for the CURRENT named 619 plan only.

    After build_wztc_order_table, call this (or follow each tool's nextStep)
    instead of rediscovering progress from chat. Returns checklist done
    flags, currentStep, nextTool, remainingSigns, stationsNeeded.

    Outside a sheet build (no order table locked): returns
    sheetPlanActive=False — general CAD tasks are NOT gated; keep
    reasoning freely. force/one_off escapes still apply on individual tools."""
    if not _PLAN_SESSION.sheet_plan_active():
        return {
            "status": "OK",
            "sheetPlanActive": False,
            "note": (
                "No named sheet plan active (build_wztc_order_table has not "
                "locked a sheet this session). Deterministic checklist does "
                "NOT apply — reason freely for general CAD / one-offs / "
                "questions. Call build_wztc_order_table when starting a 619 "
                "standard-sheet build."
            ),
            "nextTool": None,
            "nextStep": None,
        }
    out = plan_workflow.build_status_dict(_PLAN_SESSION)
    out["sheetPlanActive"] = True
    out["persistedPath"] = str(SHEET_PLAN_PATH) if SHEET_PLAN_PATH.exists() else None
    out["updatedAt"] = _PLAN_SESSION.plan_updated_at or None
    if _PLAN_SESSION.work_area_edges:
        out["workAreaEdges"] = _PLAN_SESSION.work_area_edges
    if _PLAN_SESSION.last_failed_phase:
        out["lastFailedPhase"] = _PLAN_SESSION.last_failed_phase
    if _PLAN_SESSION.last_replan:
        out["lastReplan"] = _PLAN_SESSION.last_replan
    if _PLAN_SESSION.last_scorecard is not None:
        out["scorecardPassed"] = bool(_PLAN_SESSION.last_scorecard.get("passed"))
        out["scorecardFailureCount"] = len(
            _PLAN_SESSION.last_scorecard.get("failures") or [])
    if _PLAN_SESSION.visual_qa_failures:
        out["visualQaFailures"] = list(_PLAN_SESSION.visual_qa_failures)
    di = _PLAN_SESSION.designer_inputs
    if di is not None and di.sheet_num:
        guide = sheet_spec.load_build_guide(di.sheet_num)
        if guide is not None:
            out["buildGuidePath"] = guide["path"]
            out["buildGuideCharCount"] = guide["charCount"]
            # Short excerpt so checklist turns see tips without a second call;
            # full text via get_sheet_build_guide / get_sheet_requirements.
            excerpt = guide["text"][:2000]
            if len(guide["text"]) > 2000:
                excerpt += "\n\n…[truncated — call get_sheet_build_guide for full playbook]"
            out["buildGuideExcerpt"] = excerpt
            out["nextStepHint"] = (
                (out.get("nextStep") or "")
                + " Follow buildGuidePath / buildGuideExcerpt "
                  "(get_sheet_build_guide for full text)."
            ).strip()
        else:
            out["buildGuidePath"] = None
    return out


def _attach_plan_next(resp: dict) -> dict:
    """Stamp nextStep/nextTool from checklist onto a successful tool result
    during an active sheet plan (no-op otherwise)."""
    if not isinstance(resp, dict):
        return resp
    if not _PLAN_SESSION.sheet_plan_active():
        return resp
    st = plan_workflow.build_status_dict(_PLAN_SESSION)
    resp.setdefault("planCurrentStep", st.get("currentStep"))
    resp.setdefault("nextTool", st.get("nextTool"))
    resp.setdefault("nextStep", st.get("nextStep"))
    return resp


def _merge_locked_designer_inputs(
        sheet_num: str, speed: int, lane_width: int, shoulder_width: str,
        area_type: str = "", closure_type: str = "",
        exposure_condition: str = "", protective_vehicle_gvw: int = 0,
        force: bool = False) -> dict:
    """Fill blank compile/place kwargs from locked DesignerInputs.
    Live miss 2026-08-04: place_sheet_geometry(area_type='') after the
    engineer already locked RURAL. Conflicts raise unless force=True."""
    inputs = _PLAN_SESSION.designer_inputs
    out = {
        "sheet_num": sheet_num,
        "speed": speed,
        "lane_width": lane_width,
        "shoulder_width": shoulder_width,
        "area_type": area_type or "",
        "closure_type": closure_type or "",
        "exposure_condition": exposure_condition or "",
        "protective_vehicle_gvw": protective_vehicle_gvw or 0,
        "filledFromLock": [],
    }
    if inputs is None:
        return out

    def _blank(v) -> bool:
        return v is None or (isinstance(v, str) and not str(v).strip()) or v == 0

    pairs = [
        ("sheet_num", inputs.sheet_num, sheet_num),
        ("speed", inputs.speed, speed),
        ("lane_width", inputs.lane_width, lane_width),
        ("shoulder_width", inputs.shoulder_width, shoulder_width),
        ("area_type", inputs.area_type, area_type),
        ("closure_type", inputs.closure_type, closure_type),
        ("exposure_condition", inputs.exposure_condition, exposure_condition),
        ("protective_vehicle_gvw", inputs.protective_vehicle_gvw, protective_vehicle_gvw),
    ]
    for key, locked, passed in pairs:
        if _blank(passed) and not _blank(locked):
            out[key] = locked
            out["filledFromLock"].append(key)
            continue
        if _blank(passed) or _blank(locked):
            continue
        # Normalize light string compare for shoulder/area
        lp = str(locked).strip()
        pp = str(passed).strip()
        if key in ("speed", "lane_width", "protective_vehicle_gvw"):
            try:
                same = int(locked) == int(passed)
            except (TypeError, ValueError):
                same = lp == pp
        else:
            same = lp.upper() == pp.upper()
        if not same and not force:
            raise ValueError(
                f"{key}={passed!r} conflicts with locked designer input "
                f"{key}={locked!r} from build_wztc_order_table. Reuse the "
                f"locked value (or call get_locked_designer_inputs). "
                f"Pass force=True only for an intentional override."
            )
    return out


def _refuse_if_spec_path_available(tool_name: str, force: bool) -> None:
    """Refuses a legacy heuristic PlaceOrderTable*/place_sheet_symbol_cells
    call when the current build has a real Data/sheet-specs/<sheet>.json —
    that spec-driven geometry should go through place_sheet_geometry (the
    placement-plan compiler + run_rules_gate), not this generic heuristic
    path, which has no rules-gate validation at all (confirmed live
    2026-08-04: this path is how a bad corridor could still get fully
    drawn even after run_rules_gate's corridor-topology check existed,
    simply by never calling compile_sheet_plan in the first place). Only
    checked when build_wztc_order_table already locked a sheet_num this
    session — an unrelated one-off drawing op outside a sheet build is
    unaffected. Pass force=True to use the heuristic path anyway (e.g. the
    compiler doesn't cover this layer yet for this sheet family)."""
    if force:
        return
    inputs = _PLAN_SESSION.designer_inputs
    if inputs is None:
        return
    sheet_num = inputs.sheet_num
    if not sheet_num or sheet_spec.load(sheet_num) is None:
        return
    raise ValueError(
        f"{tool_name} is the generic heuristic path with no rules-gate validation. "
        f"Data/sheet-specs/{sheet_num}.json exists for this build — use "
        f"place_sheet_geometry instead, which compiles from that spec and runs "
        f"run_rules_gate (corridor-topology, taper-continuity, cone-spacing, etc.) "
        f"before placing anything. Pass force=True only if the compiler genuinely "
        f"doesn't cover what you need here.")


def place_perp_line(align_idx: int, sta: float, half_len: float = 40,
                    reason: str = "", one_off: bool = False) -> dict:
    """Place a SINGLE perpendicular reference tick line (2*half_len ft
    long, default 80ft) at a station along a committed alignment. For a
    full-plan run, prefer place_order_table_stations instead — it places
    every order-table item's tick line (and records sign geometry) in
    ONE call rather than one call per item, which is the whole point of
    the batched op. Use this one only for a genuinely one-off tick line
    outside the order-table flow (e.g. an ad hoc reference the engineer
    asks for directly) and pass one_off=True — without that flag this
    tool refuses when the session already looks like a plan (workspace
    placed and/or order table built but stations not yet batched)."""
    if not one_off:
        if _PLAN_SESSION.order_table_built and align_idx not in _PLAN_SESSION.stations_placed_aligns:
            raise ValueError(
                f"Order table exists but place_order_table_stations has not been called "
                f"for align_idx={align_idx} yet. Call place_order_table_stations instead "
                f"(not place_perp_line item-by-item). Pass one_off=True only if the "
                f"engineer explicitly asked for a single ad-hoc tick outside the order table."
            )
        if _PLAN_SESSION.placed_workspace and not _PLAN_SESSION.order_table_built:
            raise ValueError(
                "place_workspace already ran this session — treat this as a work-zone plan. "
                "Call build_wztc_order_table (show the engineer the table), commit_alignment "
                "if needed, then place_order_table_stations. Do not place_perp_line + "
                "place_sign by hand for the plan's stations. Pass one_off=True only if the "
                "engineer explicitly asked for a single ad-hoc tick."
            )
    return _ok_or_raise(
        _bridge.call("PLACE_PERP_LINE", alignIdx=align_idx, sta=sta, halfLen=half_len, reason=reason),
        "place_perp_line")


def place_sign(sign_num: str, road_type: str, side: str,
               pt1x: float, pt1y: float, pt1z: float, dir1x: float, dir1y: float,
               pt2x: Optional[float] = None, pt2y: Optional[float] = None, pt2z: Optional[float] = None,
               dir2x: Optional[float] = None, dir2y: Optional[float] = None,
               reason: str = "", align_idx: int = 0, one_off: bool = False,
               post_angle_deg: Optional[float] = None) -> dict:
    """Place a sign assembly (post + edge-connected stem + face + label).

    If build_wztc_order_table already ran this session, sign_num MUST match
    one of its resolved sign_rows (the deterministic Table-driven legend/
    side pick) — this refuses an ad-hoc sign_num that wasn't in that list,
    the same guard shape place_perp_line already uses for stations. This
    exists because a live miss picked a manually-guessed legend variant
    (W20-01RA) instead of the order table's own resolved one (W20-01RF) by
    calling resolve_sign_code + place_sign directly rather than trusting
    build_wztc_order_table's Table 311-03-driven pick. Pass one_off=True
    only for a genuine ad-hoc sign the engineer explicitly asked for
    outside the order table.

    pt1 is the ATTACHMENT POINT ON THE PERPENDICULAR TICK — typically the
    outward tip of the 80ft perp line from place_order_table_stations /
    place_perp_line — NOT the alignment centerline station and NOT the
    face center. dir1 is the unit OUTWARD direction along that perp
    (away from the alignment). Do NOT pass the alignment tangent as dir1
    (that was a live miss: assembly built along the road instead of off
    the tick).

    From an order-table station row (ptX/ptY/tanX/tanY), compute:
      outward = rotate tan 90deg toward the chosen side (e.g. (-tanY, tanX))
      tip = (ptX, ptY) + outward * half_len   # half_len default 40
      place_sign(..., pt1=tip, dir1=outward)

    The stem connects the post's outer edge to the face's inner edge only
    (never through the face center). sign_num MUST be a SignLibrary.bas key.
    side is 'One Side' or 'Both Sides'; pt2/dir2 required only for Both Sides.

    align_idx (1=Upstream, 2=Downstream): journaled so a later
    clear_plan_elements(align_idx=…) / place_order_table_stations(
    clear_prior=True) can wipe only that alignment's signs. Always pass
    it when placing from an order-table row.

    post_angle_deg rotates TWZSGN_P with travel tangent (arm downstream,
    stem upstream, T on the curve). Omit to leave the post at view angle
    (MUTCD face/text still use view angle either way).
    """
    if side.strip().lower() == "both sides":
        missing = [n for n, v in
                   [("pt2x", pt2x), ("pt2y", pt2y), ("dir2x", dir2x), ("dir2y", dir2y)] if v is None]
        if missing:
            raise ValueError(f"side='Both Sides' requires {missing}")
    if not one_off and _PLAN_SESSION.order_table_built and _PLAN_SESSION.locked_sign_rows:
        key = str(sign_num).strip().upper()
        locked = _PLAN_SESSION.locked_sign_rows
        locked_codes = {c for _, c in locked}
        if align_idx and align_idx > 0:
            ok = (align_idx, key) in locked
        else:
            ok = key in locked_codes
        if not ok:
            plan_workflow.raise_plan_gate(
                f"sign_num={sign_num!r} is not in build_wztc_order_table's "
                f"resolved sign_rows for this sheet build.",
                tool="place_sign",
                current_step="signs_placed",
                missing=[f"requested:{sign_num}"],
                accepted=sorted(locked_codes),
                next_tool="place_sign",
                next_step=(
                    "Use a sign_num from the locked order table (accepted list). "
                    "Pass one_off=True only for a genuine ad-hoc sign outside the sheet."
                ),
            )
        # Sheet plan: stations for this align must exist before signs.
        if (align_idx and align_idx > 0
                and align_idx not in _PLAN_SESSION.stations_placed_aligns):
            plan_workflow.raise_plan_gate(
                f"stations not placed yet for align_idx={align_idx}.",
                tool="place_sign",
                current_step="stations_placed",
                missing=[f"align_idx={align_idx}"],
                next_tool="place_order_table_stations",
                next_step=f"place_order_table_stations(align_idx={align_idx}) first",
            )
    kwargs = dict(signNum=sign_num, roadType=road_type, side=side,
                  pt1X=pt1x, pt1Y=pt1y, pt1Z=pt1z, dir1X=dir1x, dir1Y=dir1y,
                  pt2X=pt2x, pt2Y=pt2y, pt2Z=pt2z, dir2X=dir2x, dir2Y=dir2y,
                  reason=reason)
    if align_idx and align_idx > 0:
        kwargs["alignIdx"] = align_idx
    if post_angle_deg is not None:
        kwargs["postAngleDeg"] = float(post_angle_deg)
    resp = _ok_or_raise(_bridge.call("PLACE_SIGN", **kwargs), "place_sign")
    if align_idx and align_idx > 0 and _PLAN_SESSION.sheet_plan_active() and not one_off:
        _PLAN_SESSION.signs_placed_rows.add(
            (int(align_idx), str(sign_num).strip().upper()))
        _save_sheet_plan()
    if isinstance(resp, dict):
        ids = placement_registry.parse_created_ids(resp)
        if ids:
            sheet = ""
            if _PLAN_SESSION.designer_inputs:
                sheet = _PLAN_SESSION.designer_inputs.sheet_num
            placement_registry.append_placement(
                sheet_num=sheet,
                align_idx=int(align_idx or 0),
                kind="sign",
                primitive_id=f"{int(align_idx or 0)}:{str(sign_num).strip().upper()}:sign",
                bridge_op="PLACE_SIGN",
                element_ids=ids,
                spec_ref={"signNum": str(sign_num).strip().upper(),
                          "zone": None, "run": None, "alignIdx": int(align_idx or 0)},
                req_id=str(resp.get("reqId") or resp.get("req_id") or ""),
                extra={
                    "x": float(pt1x), "y": float(pt1y), "z": float(pt1z or 0),
                    "signNum": str(sign_num).strip().upper(),
                },
            )
    return _attach_plan_next(resp) if isinstance(resp, dict) else resp

def place_workspace(vertices: list[list[float]], reason: str = "") -> dict:
    """Place the work space boundary (unfilled shape) + hatch stripes.
    vertices is an ordered list of [x, y, z] points — do not repeat the
    first point to close it. Creates an unfilled TWZWS2_P shape plus
    associative pattern AND real diagonal line stripes (view Patterns
    attribute often hides associative hatch alone). Response must include
    elementId — verify the shape exists with find_elements_near before
    continuing the plan. Do not substitute place_block for this."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    resp = _ok_or_raise(_bridge.call("PLACE_WORKSPACE", verticesTSV=verts_tsv, reason=reason), "place_workspace")
    _PLAN_SESSION.placed_workspace = True
    return resp


# ========================================================== Full-plan flow
# Agent-driven-8-step-wizard plan (~/.claude/plans/polished-purring-reef.md).
# These orchestrate what the manual WZTCDesigner->DrawWorkSpace->AlignDraw->
# PlacePerp wizard steps do, without opening any form. The manual wizard
# is untouched and remains the fallback.

def build_wztc_order_table(speed: int, road_type: str, lane_width: int, shoulder_width: str,
                            sign_rows: Optional[list[dict]] = None,
                            category: str = "", sheet_num: str = "",
                            area_type: str = "", closure_type: str = "",
                            exposure_condition: str = "",
                            protective_vehicle_gvw: int = 0) -> dict:
    """Headless equivalent of WZTCDesigner.frm's Submit & Draw — builds the
    full per-alignment order table and writes the same SharedState the manual
    form writes. Never estimate spacing yourself; it comes from here.

    Data/sheet-specs/<sheet_num>.json MUST exist — THE SHEET DRIVES THE TABLE:
    the station sequence, every spacing, the sign order and each sign's
    SignLibrary key all come from the sheet, and sign_rows/area_type are only
    needed to disambiguate. There is no generic fallback: every real 619
    sheet now has either a `done` spec or a documented blocker in
    Data/sheet-specs/STATUS.md (no not-started sheets remain), so a missing
    spec means either a real gap that should be reported, not guessed
    around, or a genuinely blocked sheet (missing source PDF, etc.) that
    cannot be drawn yet. The old fallback emitted the same 7 upstream rows
    for every sheet, including stations (Vehicle Space, temporary barrier,
    box/corr beam) most sheets do not have, and interpolated shoulder taper
    values tables don't print — raises ValueError now instead of silently
    drawing that.
    Pass area_type ("URBAN"/"RURAL"/"FREEWAY") only when the sheet's
    tableRoles include advanceWarningSpacing; shoulder/freeway sheets
    without that role (e.g. 619-301) do not need it. Pass
    protective_vehicle_gvw (lbs) when roll-ahead is GVW-keyed; 0 means
    use sheet_spec's default (22000).

    sign_rows (optional when a spec exists): list of dicts, each
    {"align_idx": 1|2, "sign_num": SignLibrary key, "side": "One Side"|
    "Both Sides", "spacing_ft": optional, "size": optional}.

    Returns the order table (rows: alignIdx, alignName, rowNum, type, label,
    spacing, size, side) — show it to the engineer before drawing."""
    sign_rows = list(sign_rows or [])

    spec = sheet_spec.load(sheet_num) if sheet_num else None
    if spec is None:
        raise ValueError(
            f"no verified sheet spec for {sheet_num!r} "
            f"(Data/sheet-specs/{sheet_num}.json does not exist) — refusing to draw a "
            f"generic/guessed order table. Check Data/sheet-specs/STATUS.md: this sheet "
            f"is either genuinely blocked (tell the engineer why) or its spec needs to "
            f"be authored first, not worked around.")

    roles = spec.get("tableRoles") or {}
    needs_area = bool(roles.get("advanceWarningSpacing"))
    if needs_area and not area_type:
        raise ValueError(
            f"sheet {sheet_num} has an advance-warning spacing table; pass "
            f"area_type='URBAN'/'RURAL'/'FREEWAY' as the sheet's table keys it. "
            f"(Sheets without that role — e.g. freeway shoulder 619-301 — do not "
            f"need area_type.)")
    gvw = protective_vehicle_gvw if protective_vehicle_gvw and protective_vehicle_gvw > 0 else None
    resolved = sheet_spec.resolve(
        spec, speed, lane_width, shoulder_width,
        area_type or None,
        closure_type or None, exposure_condition or None,
        protective_vehicle_gvw=gvw)
    payload = sheet_spec.order_table_rows(
        spec, resolved,
        size_class="FREEWAY" if road_type.strip().lower() == "freeway" else "NON-FREEWAY")
    spec_rows_tsv = "|".join(payload["nonSignRows"])
    if not sign_rows:
        sign_rows = [{"align_idx": int(s.split(":")[0]), "sign_num": s.split(":")[1],
                      "side": s.split(":")[2], "spacing_ft": s.split(":")[3],
                      "size": s.split(":")[4]}
                     for s in payload["signRows"]]
    # Per-taper counts, not totals: wztcSkipLines is merge + shoulder +
    # buffer + roll ahead, and the sheet gives no skip count for the last
    # two. VBA substitutes only the taper terms and keeps ComputeSpacing's
    # buffer/roll-ahead skips rather than inventing sheet-less numbers.
    # Optional fields: shoulder-only / mobile / barrier sheets omit some.
    lane = resolved.get("laneTaper") or {}
    sh = resolved.get("shoulderTaper") or {}
    roll = resolved.get("rollAheadFt") or {}
    overrides_tsv = "|".join([
        f"bufferSpace={resolved.get('bufferFt', '')}",
        f"mergingTaper={lane.get('ft', '')}",
        f"shoulderTapers={sh.get('ft', '')}",
        f"rollAhead={roll.get('min', '')}",
        f"laneTaperSkips={lane.get('skipLines', '')}",
        f"shoulderTaperSkips={sh.get('skipLines', '')}",
        f"laneTaperDevices={lane.get('devices', '')}",
        f"shoulderTaperDevices={sh.get('devices', '')}",
    ])
    spec_info = {
        "specDriven": True,
        "sheet": spec["sheet"]["number"],
        "shoulderBandUsed": resolved.get("shoulderBand"),
        "signLegends": resolved.get("legend") or {},
        "overlays": payload["overlays"],
        "stationWalk": sheet_spec.station_walk(spec, resolved),
        "note": "Stations, spacings and SignLibrary keys came from the standard "
                "sheet spec, not WZTCRules defaults.",
    }

    rows_tsv = "|".join(
        f"{r['align_idx']}:{r['sign_num']}:{r.get('side', 'One Side')}:"
        f"{r.get('spacing_ft', '')}:{r.get('size', '')}"
        for r in sign_rows
    )
    resp = _ok_or_raise(
        _bridge.call("BUILD_WZTC_ORDER_TABLE", category=category, sheetNum=sheet_num,
                     speed=speed, roadType=road_type, laneWidth=lane_width,
                     shoulderWidth=shoulder_width, signRowsTSV=rows_tsv,
                     nonSignRowsTSV=spec_rows_tsv, spacingOverridesTSV=overrides_tsv),
        "build_wztc_order_table")
    resp.update(spec_info)
    _attach_build_guide_fields(sheet_num, resp)
    _PLAN_SESSION.order_table_built = True
    _PLAN_SESSION.stations_placed_aligns = set()
    _PLAN_SESSION.find_near_calls = 0
    _PLAN_SESSION.sheet_geometry_placed = False
    _PLAN_SESSION.signs_placed_rows = set()
    _PLAN_SESSION.sign_attrs_applied = False
    _PLAN_SESSION.geometry_qa_passed = False
    _PLAN_SESSION.visual_qa_passed = False
    _PLAN_SESSION.last_station_rows = {}
    req_aligns: set[int] = set()
    for a in (spec.get("orderTable") or {}).get("alignments") or []:
        try:
            req_aligns.add(int(a.get("alignIdx") or a.get("align_idx") or 0))
        except (TypeError, ValueError):
            pass
    for r in sign_rows:
        try:
            req_aligns.add(int(r["align_idx"]))
        except (KeyError, TypeError, ValueError):
            pass
    req_aligns.discard(0)
    _PLAN_SESSION.required_aligns = req_aligns or {1, 2}
    _PLAN_SESSION.lock_designer_inputs(
        sheet_num=sheet_num, speed=speed, road_type=road_type,
        lane_width=lane_width, shoulder_width=shoulder_width,
        area_type=area_type, closure_type=closure_type,
        exposure_condition=exposure_condition,
        protective_vehicle_gvw=protective_vehicle_gvw,
    )
    _PLAN_SESSION.lock_sign_rows(sign_rows)
    placement_registry.clear_registry()
    _save_sheet_plan()
    resp["highwayCaution"] = _highway_caution_for_sheet(sheet_num)
    return _attach_plan_next(resp)


def find_reference_linework(level_name_contains: str, include_references: bool = False,
                            ref_name_contains: str = "", force: bool = False) -> list[dict]:
    """Locate connected line/line-string chains on a level, for auto-
    tracing an alignment or work-space boundary without clicks. Ask the
    engineer which level holds the roadway centerline first — never guess
    a level name. include_references=False (default) scans only the
    active model; pass True to also scan attached reference files (their
    own geometry can be genuinely unavailable session to session — treat
    that as a normal, recoverable condition, not a bug, and fall back to
    asking the engineer to click points if nothing plausible comes back).
    Arc segments are not included (line-based geometry only) — a chain
    with true arcs will come back broken into separate pieces.
    Returns one row per disconnected candidate chain (chainIdx, source,
    segmentCount, vertexCount, totalLengthFt, verticesTSV) — usually the
    longest is the intended roadway, but don't assume; a short/odd result
    should be confirmed with the engineer rather than used blindly.
    verticesTSV feeds straight into define_alignment_segment/
    place_workspace with no re-encoding.

    After build_wztc_order_table: refuse vague Default/RDEFAULT fishing
    (live 2026-08-04) — prefer assemble_corridor with work-area edge
    point-picks. force=True to override."""
    needle = (level_name_contains or "").strip().lower()
    vague = needle in ("default", "rdefault", "def", "level default")
    if _PLAN_SESSION.order_table_built and vague and not force:
        raise ValueError(
            f"find_reference_linework(level={level_name_contains!r}) is too "
            f"broad during a sheet plan (Default matches dozens of elements "
            f"and will be refused by the bridge). Prefer "
            f"assemble_corridor(upstream_edge, downstream_edge) after "
            f"ask_user_choice(allow_point_pick=True) for the two WORK AREA "
            f"edges. Pass a specific CL/ROAD level name, or force=True only "
            f"if the engineer named Default explicitly."
        )
    resp = _ok_or_raise(
        _bridge.call("FIND_REFERENCE_LINEWORK", levelNameContains=level_name_contains,
                     includeReferences="Y" if include_references else "N",
                     refNameContains=ref_name_contains),
        "find_reference_linework")
    return resp.get("rows", [])


def define_alignment_segment(align_idx: int, vertices: list[list[float]],
                             reason: str = "", force: bool = False) -> dict:
    """Create straight alignment line segments from vertices (Default
    level/color 0/weight 0) and record them as one drawing session for
    align_idx — the same bookkeeping AlignDraw's interactive clicking
    produces. vertices come from find_reference_linework's verticesTSV
    (parsed back to [[x,y,z],...]) or from repeated ask_user_choice
    point-picks when no usable reference geometry exists. Call this one
    or more times per alignment, then commit_alignment once when done.
    align_idx convention: 1=Upstream, 2=Downstream (matches
    build_wztc_order_table).

    When a sheet order table is locked and alignments are not both ready,
    prefer assemble_corridor over freestyle define+commit pairs (live
    2026-08-04 Downstream-on-Upstream miss). For curved corridors pass
    path_vertices to assemble_corridor / run_sheet_build instead of
    freestyle define. Pass force=True only for engineer-directed redefine
    or adopt-recovery edge cases."""
    if (_PLAN_SESSION.order_table_built
            and _PLAN_SESSION.designer_inputs is not None
            and sheet_spec.has_spec(_PLAN_SESSION.designer_inputs.sheet_num)
            and not ({1, 2} <= _PLAN_SESSION.aligns_ready)
            and not force):
        raise ValueError(
            "define_alignment_segment refused during a sheet-spec plan — "
            "call assemble_corridor(upstream_edge, downstream_edge, "
            "path_vertices=… if curved) after point-picking the two "
            "work-area edges (prevents Downstream committed along "
            "Upstream's line). Pass force=True only when the engineer "
            "explicitly asked to define segments by hand."
        )
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(
        _bridge.call("DEFINE_ALIGNMENT_SEGMENT", alignIdx=align_idx, verticesTSV=verts_tsv, reason=reason),
        "define_alignment_segment")


def commit_alignment(align_idx: int, force: bool = False) -> dict:
    """Group every segment recorded by define_alignment_segment for
    align_idx into a graphic group, marking that alignment ready for
    place_order_table_stations. Call once per alignment after all its
    define_alignment_segment calls.

    Once BOTH align 1 and align 2 are committed/adopted (and this build has
    a locked sheet spec), this runs the corridor-topology check immediately
    — catching a bad corridor (e.g. Downstream committed further along
    Upstream's own line instead of at a distinct work-area edge) right when
    it's created, not only later when place_sheet_geometry compiles a full
    plan against it. Raises if the check fails; pass force=True to proceed
    anyway (the commit itself already succeeded either way — nothing to
    undo)."""
    resp = _ok_or_raise(_bridge.call("COMMIT_ALIGNMENT", alignIdx=align_idx), "commit_alignment")
    warning = _check_corridor_topology_if_ready(align_idx, force)
    if warning:
        resp["corridorTopologyWarning"] = warning
    if not _PLAN_SESSION.order_table_built:
        resp["nextStep"] = (
            "build_wztc_order_table (show engineer), then place_order_table_stations — "
            "do not place_perp_line/place_sign by hand for plan stations"
        )
    elif align_idx not in _PLAN_SESSION.stations_placed_aligns:
        resp["nextStep"] = f"place_order_table_stations(align_idx={align_idx}, reset_session=...)"
    return resp


def adopt_alignment(align_idx: int, element_id: str, force: bool = False) -> dict:
    """Re-bind SharedState for align_idx to an EXISTING LINE element in the
    model — without redrawing it. Use after a VBA hot-reload / IDE Reset
    wiped in-memory alignment session state while the corridor geometry
    is still on screen (live miss 2026-08-04). Also use when the engineer
    element-picks an existing centerline to adopt as Upstream/Downstream.

    element_id must be a LINE (not a complex chain). align_idx: 1=Upstream,
    2=Downstream. After adopt, place_order_table_stations / station_to_point
    / get_alignment_vertices work again for that align_idx.

    Once BOTH align 1 and align 2 are committed/adopted (and this build has
    a locked sheet spec), this runs the corridor-topology check immediately
    — same as commit_alignment. Raises if it fails; force=True to proceed
    anyway (the adopt itself already succeeded either way)."""
    resp = _ok_or_raise(
        _bridge.call("ADOPT_ALIGNMENT_ELEMENT", alignIdx=align_idx, elementId=element_id),
        "adopt_alignment")
    warning = _check_corridor_topology_if_ready(align_idx, force)
    if warning:
        resp["corridorTopologyWarning"] = warning
    return resp


def _pt3(p) -> list[float]:
    if not isinstance(p, (list, tuple)) or len(p) < 2:
        raise ValueError(f"point must be [x,y] or [x,y,z], got {p!r}")
    return [float(p[0]), float(p[1]), float(p[2]) if len(p) > 2 else 0.0]


def resolve_sheet_lateral(
        upstream_edge: list[float],
        downstream_edge: list[float],
        closed_side: str,
        lane_width_ft: float = 0.0,
        shoulder_width_ft: float = 0.0,
        real_road_edge: bool = True,
        yellow_gap_ft: float = 2.0,
        opposing_lanes: int = 2,
        path_vertices: list | None = None) -> dict:
    """Lock outward_sign + half_len for a named-sheet build from travel
    and closed-lane side (Cursor real-road method, 2026-08-10).

    Contract matches assemble_corridor: travel = unit(up→dn). Align1
    stations increase UPSTREAM (−travel). closed_side is relative to
    travel ('right' | 'left'). Right-lane sheets (619-311) use 'right'.

    path_vertices: optional curved closed-lane / first-travel outer
    polyline. When set, travel is the path tangent at mid work-bay
    (same orientation rules as assemble_corridor), not the chord.

    outward_sign: +1 for closed right of travel, −1 for closed left —
    so Align1's OutwardUnit points into the closed lane / toward the
    shoulder (verified: EB +X travel, Align1 tan west, right → +1 → −Y).

    half_len: when real_road_edge, lane_width_ft + shoulder_width_ft so
    sign/AP tips land on outer EOP (12+8→20). When false or widths
    missing, half_len=40 (abstract ticks).

    Also locks closed_outward (world XY) so Align2 one-side signs
    (including G20-2) tip on the SAME closed shoulder as Align1 advance
    signs — Align2 tan alone flips _outward_unit across the road.

    Locks PlanSession lateral_* used by run_sheet_build /
    place_sheet_geometry. Call after work-area edges are known and
    BEFORE run_sheet_build. Ask the engineer for closed_side / travel
    if unknown — do not invent.
    """
    import math
    import alignment_geometry as ag

    side = (closed_side or "").strip().lower()
    if side in ("r", "right-lane", "right_lane", "outer-right"):
        side = "right"
    if side in ("l", "left-lane", "left_lane", "outer-left"):
        side = "left"
    if side not in ("right", "left"):
        raise ValueError(
            "resolve_sheet_lateral: closed_side must be 'right' or 'left' "
            f"(relative to travel through the work bay); got {closed_side!r}")

    up = _pt3(upstream_edge)
    dn = _pt3(downstream_edge)
    curved = False
    if path_vertices is not None and len(path_vertices) >= 2:
        path_pts = [_pt3(p) for p in path_vertices]
        segs = ag.segments_from_polyline(path_pts)
        sta_up, _ = ag.nearest_station(segs, up[0], up[1])
        sta_dn, _ = ag.nearest_station(segs, dn[0], dn[1])
        if sta_dn < sta_up:
            path_pts = list(reversed(path_pts))
            segs = ag.segments_from_polyline(path_pts)
            sta_up, _ = ag.nearest_station(segs, up[0], up[1])
            sta_dn, _ = ag.nearest_station(segs, dn[0], dn[1])
        work_len = abs(sta_dn - sta_up)
        if work_len < 1.0:
            raise ValueError(
                f"resolve_sheet_lateral: edges only {work_len:.3f} ft apart "
                "along path_vertices")
        mid = 0.5 * (sta_up + sta_dn)
        _, _, tx, ty = ag.point_at_extended(segs, mid)
        curved = True
    else:
        dx, dy = dn[0] - up[0], dn[1] - up[1]
        work_len = math.hypot(dx, dy)
        if work_len < 1.0:
            raise ValueError(
                f"resolve_sheet_lateral: edges only {work_len:.3f} ft apart")
        tx, ty = dx / work_len, dy / work_len
    # Align1 tan at sta0 points upstream (opposite travel).
    a1_tx, a1_ty = -tx, -ty
    outward_sign = 1.0 if side == "right" else -1.0
    from sheet_compile import _outward_unit
    out_x, out_y = _outward_unit(a1_tx, a1_ty, outward_sign)

    lw = float(lane_width_ft or 0.0)
    sh = float(shoulder_width_ft or 0.0)
    if _PLAN_SESSION.designer_inputs is not None:
        if lw <= 0:
            lw = float(_PLAN_SESSION.designer_inputs.lane_width or 0)
        if sh <= 0:
            sh = _shoulder_ft_from_band(
                _PLAN_SESSION.designer_inputs.shoulder_width)

    use_real = bool(real_road_edge) and lw > 0
    half_len = (lw + max(sh, 0.0)) if use_real else 40.0

    _PLAN_SESSION.lateral_outward_sign = outward_sign
    _PLAN_SESSION.lateral_half_len = half_len
    _PLAN_SESSION.closed_side = side
    _PLAN_SESSION.real_road_edge = use_real
    _PLAN_SESSION.closed_outward_x = float(out_x)
    _PLAN_SESSION.closed_outward_y = float(out_y)
    _PLAN_SESSION.opposite_half_len = None
    if _PLAN_SESSION.work_area_edges is None:
        _PLAN_SESSION.work_area_edges = {
            "upstream": list(up),
            "downstream": list(dn),
        }
    _save_sheet_plan()

    return _attach_plan_next({
        "status": "OK",
        "closed_side": side,
        "outward_sign": outward_sign,
        "half_len": half_len,
        "real_road_edge": use_real,
        "curved": curved,
        "workAreaLengthFt": round(work_len, 3),
        "travelUnit": [round(tx, 6), round(ty, 6)],
        "align1TanUpstream": [round(a1_tx, 6), round(a1_ty, 6)],
        "outwardUnit": [round(out_x, 6), round(out_y, 6)],
        "closed_outward": [round(out_x, 6), round(out_y, 6)],
        "lane_width_ft": lw,
        "shoulder_width_ft": sh,
        "note": (
            f"Locked lateral for run_sheet_build: outward_sign={outward_sign:g}, "
            f"half_len={half_len:g} "
            f"({('real EOP' if use_real else 'abstract 40')}"
            f"{', curved path' if curved else ''}). "
            "closed_outward keeps Align1+Align2 one-side signs (incl. G20-2) "
            "on the closed shoulder. Pass the same edges (+ path_vertices) "
            "to run_sheet_build; locked values apply unless "
            "use_locked_lateral=False."
        ),
    })


def _shoulder_ft_from_band(band: str) -> float:
    """Best-effort feet from a sheet shoulder band or '8 ft' label."""
    import re
    s = (band or "").strip().lower()
    if not s:
        return 0.0
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", s)]
    if not nums:
        return 0.0
    if ">=" in s or "≥" in s:
        return nums[0]
    if "<=" in s or "≤" in s:
        return nums[0]
    if "-" in s or " to " in s:
        return nums[-1]
    return nums[0]


def _apply_locked_lateral(outward_sign: float, half_len: float,
                          use_locked_lateral: bool) -> tuple[float, float, dict]:
    """Prefer PlanSession lateral_* when resolve_sheet_lateral ran."""
    meta = {"usedLockedLateral": False}
    if not use_locked_lateral:
        return outward_sign, half_len, meta
    s = _PLAN_SESSION
    if s.lateral_outward_sign is not None:
        outward_sign = float(s.lateral_outward_sign)
        meta["usedLockedLateral"] = True
        meta["locked_outward_sign"] = outward_sign
    if s.lateral_half_len is not None:
        half_len = float(s.lateral_half_len)
        meta["usedLockedLateral"] = True
        meta["locked_half_len"] = half_len
    return outward_sign, half_len, meta


def assemble_corridor(upstream_edge: list[float], downstream_edge: list[float],
                      approach_length_ft: float = 0.0,
                      force: bool = False,
                      path_vertices: list | None = None) -> dict:
    """Build both plan alignments from the two work-area edge points.

    Contract (sheet orderTable.alignments[].station0):
      Align1 sta0 = upstream work-area edge; station increases AWAY upstream
      Align2 sta0 = downstream work-area edge; station increases AWAY downstream
    Work length = distance between the two edges (compile_hatch uses that).

    Vertices drawn (first vertex = station 0):
      Upstream:   [up_edge, up_edge - T * approach]
      Downstream: [dn_edge, dn_edge + T * approach]
    where T is the unit travel direction through the work bay
    (upstream_edge → downstream_edge).

    path_vertices: optional polyline (>=2 points) along the closed-lane /
    first-travel outer edge spanning the work bay (and preferably beyond).
    When set, Align1/Align2 follow that path (extended by approach past the
    projected edges) instead of a straight chord — required for curved /
    S-shaped real-road corridors. Work-bay hatch also follows this path.

    approach_length_ft=0 (default) auto-sizes from the locked sheet's
    station_walk max + 50 ft slack so ticks never clamp. Requires
    build_wztc_order_table first (locked DesignerInputs).

    Prefer this over freestyle define_alignment_segment pairs — live
    2026-08-04 Downstream was committed +1000 ft along Upstream's own
    line, which topology now catches but this primitive prevents.

    If alignments are already ready this session, pass force=True to
    clear_plan_elements(keep_alignments=False) first (wipes corridor +
    plan geometry and resets VBA alignment bookkeeping)."""
    import math
    import alignment_geometry as ag

    inputs = _PLAN_SESSION.designer_inputs
    if inputs is None:
        raise ValueError(
            "assemble_corridor requires build_wztc_order_table first "
            "(locked designer inputs + sheet spec drive approach length).")
    up = _pt3(upstream_edge)
    dn = _pt3(downstream_edge)

    spec = sheet_spec.load(inputs.sheet_num)
    if spec is None:
        raise ValueError(
            f"assemble_corridor: no sheet spec for {inputs.sheet_num!r}")
    resolved = sheet_spec.resolve(
        spec, inputs.speed, inputs.lane_width, inputs.shoulder_width,
        inputs.area_type or None, inputs.closure_type or None,
        inputs.exposure_condition or None,
        protective_vehicle_gvw=inputs.protective_vehicle_gvw or None)
    walk = sheet_spec.station_walk(spec, resolved)
    max_need = max((float(w["stationFt"]) for w in walk), default=0.0)
    approach = float(approach_length_ft) if approach_length_ft and approach_length_ft > 0 else (
        max_need + 50.0)
    if approach < max_need:
        raise ValueError(
            f"assemble_corridor: approach_length_ft={approach:.1f} is shorter "
            f"than station_walk max {max_need:.1f} ft — ticks will clamp. "
            f"Omit approach_length_ft for auto, or pass >= {max_need + 50:.1f}.")

    overlap = check_build_overlap(
        sheet_num=inputs.sheet_num,
        origin=list(upstream_edge)[:2],
        path_vertices=path_vertices,
        lateral_half_width=float(_PLAN_SESSION.lateral_half_len or 40.0),
        scan_model=True,
    )

    if _PLAN_SESSION.aligns_ready & {1, 2}:
        if not force:
            raise ValueError(
                "assemble_corridor: alignments already ready this session. "
                "Pass force=True to wipe corridor via "
                "clear_plan_elements(keep_alignments=False) and rebuild, "
                "or adopt_alignment if recovering SharedState only.")
        clear_plan_elements(keep_alignments=False)

    curved = False
    work_bay: list[list[float]]
    path_meta: dict = {}

    if path_vertices is not None:
        if not isinstance(path_vertices, (list, tuple)) or len(path_vertices) < 2:
            raise ValueError(
                "assemble_corridor: path_vertices needs >= 2 points "
                f"(got {path_vertices!r})")
        path_pts = [_pt3(p) for p in path_vertices]
        segs = ag.segments_from_polyline(path_pts)
        sta_up, d_up = ag.nearest_station(segs, up[0], up[1])
        sta_dn, d_dn = ag.nearest_station(segs, dn[0], dn[1])
        max_off = max(d_up, d_dn)
        if max_off > 80.0:
            raise ValueError(
                f"assemble_corridor: work-area edge is {max_off:.1f} ft from "
                "path_vertices (max 80). Pass the closed-lane / first-travel "
                "outer polyline that the upstream/downstream picks sit on.")
        if abs(sta_dn - sta_up) < 1.0:
            raise ValueError(
                "assemble_corridor: upstream and downstream edges project to "
                f"nearly the same path station ({sta_up:.1f} / {sta_dn:.1f}).")
        if sta_dn < sta_up:
            path_pts = list(reversed(path_pts))
            segs = ag.segments_from_polyline(path_pts)
            sta_up, d_up = ag.nearest_station(segs, up[0], up[1])
            sta_dn, d_dn = ag.nearest_station(segs, dn[0], dn[1])
            if sta_dn < sta_up:
                raise ValueError(
                    "assemble_corridor: could not orient path_vertices so "
                    "downstream is downstream of upstream along the path.")
        work_len = sta_dn - sta_up
        # Snap sta0 to the path so Align1/2 + hatch share one geometry.
        up_s = list(ag.point_at_extended(segs, sta_up)[:2]) + [up[2]]
        dn_s = list(ag.point_at_extended(segs, sta_dn)[:2]) + [dn[2]]
        # Align1: sta0 at up, stations increase AWAY upstream (= −path).
        up_verts = ag.sample_path_vertices(
            segs, sta_up, sta_up - approach, step_ft=10.0)
        # Align2: sta0 at dn, stations increase AWAY downstream (= +path).
        dn_verts = ag.sample_path_vertices(
            segs, sta_dn, sta_dn + approach, step_ft=10.0)
        _PLAN_SESSION.corridor_path = [[float(p[0]), float(p[1])] for p in path_pts]
        # Ensure first vertex is the snapped edge (sample may duplicate).
        up_verts[0] = up_s
        dn_verts[0] = dn_s
        work_bay = ag.sample_path_vertices(
            segs, sta_up, sta_dn, step_ft=5.0)
        mid_sta = 0.5 * (sta_up + sta_dn)
        _, _, tx, ty = ag.point_at_extended(segs, mid_sta)
        curved = True
        path_meta = {
            "pathVertexCount": len(path_pts),
            "staUpstreamFt": round(sta_up, 3),
            "staDownstreamFt": round(sta_dn, 3),
            "edgeOffsetFt": {
                "upstream": round(d_up, 3),
                "downstream": round(d_dn, 3),
            },
            "snappedSta0": True,
        }
    else:
        dx, dy = dn[0] - up[0], dn[1] - up[1]
        work_len = math.hypot(dx, dy)
        if work_len < 1.0:
            raise ValueError(
                f"assemble_corridor: upstream and downstream edges are only "
                f"{work_len:.3f} ft apart — need two distinct work-area edges.")
        tx, ty = dx / work_len, dy / work_len
        up_out = [up[0] - tx * approach, up[1] - ty * approach, up[2]]
        dn_out = [dn[0] + tx * approach, dn[1] + ty * approach, dn[2]]
        up_verts = [up, up_out]
        dn_verts = [dn, dn_out]
        work_bay = [list(up), list(dn)]
        up_s, dn_s = up, dn

    # force=True on define: freestyle define is refused mid-plan; this
    # primitive is the allowed path and must not trip its own gate.
    d1 = define_alignment_segment(
        1, up_verts,
        reason=(
            f"assemble_corridor Upstream edge→away ({approach:.0f} ft"
            f"{', curved' if curved else ''})"
        ),
        force=True)
    c1 = commit_alignment(1, force=force)
    d2 = define_alignment_segment(
        2, dn_verts,
        reason=(
            f"assemble_corridor Downstream edge→away ({approach:.0f} ft"
            f"{', curved' if curved else ''})"
        ),
        force=True)
    c2 = commit_alignment(2, force=force)

    _PLAN_SESSION.work_area_edges = {
        "upstream": list(up),
        "downstream": list(dn),
        "upstreamSta0": list(up_s),
        "downstreamSta0": list(dn_s),
    }
    _PLAN_SESSION.work_bay_vertices = work_bay
    _save_sheet_plan()

    return {
        "status": "OK",
        "workAreaLengthFt": round(work_len, 3),
        "approachLengthFt": round(approach, 3),
        "stationWalkMaxFt": round(max_need, 3),
        "travelUnit": [round(tx, 6), round(ty, 6)],
        "curved": curved,
        "pathMeta": path_meta or None,
        "workBayVertexCount": len(work_bay),
        "upstream": {
            "sta0": up_s, "vertexCount": len(up_verts),
            "define": d1, "commit": c1,
        },
        "downstream": {
            "sta0": dn_s, "vertexCount": len(dn_verts),
            "define": d2, "commit": c2,
        },
        "nextStep": (
            "place_order_table_stations per alignment (runs cross_validate), "
            "then place_sign / place_sheet_geometry"
        ),
        "overlapCaution": (overlap or {}).get("overlapCaution"),
    }


def cross_validate_stations(align_idx: int = 0, tol_ft: float = 0.5,
                            force: bool = False) -> dict:
    """Compare VBA get_alignment_stationing vs Python station_walk, and
    ensure the drawn path is long enough for the farthest walk station.

    align_idx=0 checks every locked order-table alignment that is ready
    (typically 1 and 2). Called automatically by place_order_table_stations
    and place_sheet_geometry unless force=True on those ops.

    Raises ValueError listing mismatches unless force=True (then returns
    them under failures / warning)."""
    inputs = _PLAN_SESSION.designer_inputs
    if inputs is None:
        raise ValueError(
            "cross_validate_stations requires build_wztc_order_table first "
            "(locked designer inputs).")
    spec = sheet_spec.load(inputs.sheet_num)
    if spec is None:
        raise ValueError(
            f"cross_validate_stations: no sheet spec for {inputs.sheet_num!r}")
    resolved = sheet_spec.resolve(
        spec, inputs.speed, inputs.lane_width, inputs.shoulder_width,
        inputs.area_type or None, inputs.closure_type or None,
        inputs.exposure_condition or None,
        protective_vehicle_gvw=inputs.protective_vehicle_gvw or None)
    walk_all = sheet_spec.station_walk(spec, resolved)

    idxs = [align_idx] if align_idx and align_idx > 0 else sorted(
        {int(a["alignIdx"]) for a in (spec.get("orderTable") or {}).get("alignments") or []}
        or [1, 2]
    )

    import alignment_geometry as ag
    failures: list[str] = []
    per_align: list[dict] = []
    for aidx in idxs:
        vba_rows = get_alignment_stationing(aidx)
        walk_rows = [w for w in walk_all if int(w["alignIdx"]) == aidx]
        table_fails = sheet_spec.compare_station_tables(vba_rows, walk_rows, tol_ft=tol_ft)
        segs = ag.parse_vertices(get_alignment_vertices(aidx))
        path_len = ag.total_length(segs)
        max_walk = max((float(w["stationFt"]) for w in walk_rows), default=0.0)
        path_fails: list[str] = []
        if path_len + tol_ft < max_walk:
            path_fails.append(
                f"cross-validate: align {aidx} path length {path_len:.1f} ft "
                f"< station_walk max {max_walk:.1f} ft — extend via "
                f"assemble_corridor (longer approach) or redefine the "
                f"alignment; otherwise place_order_table_stations will clamp."
            )
        all_fails = table_fails + path_fails
        failures.extend(all_fails)
        per_align.append({
            "alignIdx": aidx,
            "vbaRowCount": len(vba_rows),
            "walkRowCount": len([w for w in walk_rows if w.get("rowNum") is not None]),
            "pathLengthFt": round(path_len, 3),
            "stationWalkMaxFt": round(max_walk, 3),
            "failures": all_fails,
        })

    result = {
        "status": "OK" if not failures else "FAIL",
        "tolFt": tol_ft,
        "alignments": per_align,
        "failures": failures,
    }
    if failures and not force:
        raise ValueError(
            "cross_validate_stations failed: " + "; ".join(failures) +
            " Fix the corridor (prefer assemble_corridor) or rebuild the "
            "order table; pass force=True only to proceed knowingly."
        )
    if failures and force:
        result["warning"] = (
            f"cross_validate_stations failures ignored (force=True): "
            f"{'; '.join(failures)}"
        )
    return result


def place_order_table_stations(align_idx: int, reset_session: bool = False,
                                clear_prior: bool = False,
                                force: bool = False) -> dict:
    """Batched replacement for PlacePerp.frm's interactive walk — places
    an 80ft perpendicular tick line at EVERY row in align_idx's order
    table in one call (instead of one place_perp_line-equivalent call per
    item), using the same station math build_wztc_order_table's spacing
    values drive. Requires build_wztc_order_table AND commit_alignment
    for this align_idx first.
    ALWAYS call this — not repeated place_perp_line calls — once an
    alignment is committed as part of a full-plan run. Calling
    place_perp_line once per order-table item defeats the entire purpose
    of this tool (collapsing N tool-call round-trips into 1) and costs
    real money for no benefit; only reach for place_perp_line directly
    for a genuinely one-off tick line outside this flow.
    reset_session=True clears prior placed-sign bookkeeping — pass True
    for the FIRST alignment in a fresh plan run, False (default) for any
    subsequent alignment in the same run so sign geometry accumulates
    correctly across alignments rather than being wiped.
    clear_prior=True calls clear_plan_elements(align_idx=…) first — scoped
    to THIS alignment only (keeps the other alignment's ticks/signs). Use
    clear_plan_elements() with no align_idx for a full plan wipe. Without
    a clear, a second place stacks ticks / cells / channelizing on top of
    the previous run (the non-idempotent failure mode). If stations were
    already placed for this align_idx this session and clear_prior/force
    are both False, this refuses.
    Returns one row per order-table item (itemNum, label, type,
    cumulativeStationFt, ptX, ptY, ptZ, tanX, tanY, isSign). isSign=N rows
    get a tick only at this step — follow with place_order_table_labels
    (Non-Sign names), place_order_table_dimensions (spacings), and
    place_sheet_symbol_cells (ProtectiveVehicle/ArrowPanel). For isSign=Y
    rows: resolve_sign_code, then place_sign at the OUTWARD PERP TIP — NOT
    at (ptX,ptY) and NOT with dir=tangent. Compute:
      outward = rotate (tanX,tanY) 90deg toward the chosen side
      tip = (ptX,ptY) + outward * half_len   # half_len default 40 (80ft tick)
      place_sign(..., pt1=tip, dir1=outward)
    VBA builds the edge-connected stem/post/face from that tip; wrong pt1/dir
    is what produced assemblies along the road or floating off the tick."""
    if _PLAN_SESSION.sheet_plan_active():
        if align_idx not in _PLAN_SESSION.aligns_ready:
            plan_workflow.raise_plan_gate(
                f"align_idx={align_idx} is not committed/adopted yet.",
                tool="place_order_table_stations",
                current_step="corridor_ready",
                missing=[f"align_idx={align_idx}"],
                next_tool="assemble_corridor",
                next_step=(
                    "assemble_corridor(upstream_edge, downstream_edge) after "
                    "point-picking work-area edges (or commit/adopt this align)."
                ),
            )
    already = align_idx in _PLAN_SESSION.stations_placed_aligns
    if already and not clear_prior and not force:
        raise ValueError(
            f"stations already placed for align_idx={align_idx} this session. "
            f"Call clear_plan_elements() (or pass clear_prior=True) before "
            f"rebuilding — otherwise ticks/cells/channelizing stack on the "
            f"previous run. Pass force=True only for intentional additive placement."
        )
    # Preflight: VBA order-table stations must match Python station_walk and
    # the drawn path must be long enough (else ticks clamp at path end).
    # force=True softens to a warning so intentional overrides still work.
    xv = None
    if _PLAN_SESSION.designer_inputs is not None:
        try:
            xv = cross_validate_stations(align_idx=align_idx, force=force)
        except ValueError:
            raise
    cleared = None
    if clear_prior:
        cleared = clear_plan_elements(keep_alignments=True, align_idx=align_idx)
    resp = _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_STATIONS", alignIdx=align_idx,
                     resetSession="Y" if reset_session else "N"),
        "place_order_table_stations")
    _PLAN_SESSION.stations_placed_aligns.add(align_idx)
    if isinstance(resp, dict) and resp.get("rows") is not None:
        _PLAN_SESSION.last_station_rows[int(align_idx)] = list(resp.get("rows") or [])
    _save_sheet_plan()
    if cleared is not None:
        resp["clearedPrior"] = cleared
    if xv is not None:
        resp["crossValidate"] = xv
    return _attach_plan_next(resp)


def place_order_table_labels(align_idx: int, outward_sign: float = -1.0,
                             text_extra_along: float = 20.0,
                             sheet_elements: str = "", force: bool = False) -> dict:
    """Name labels BELOW tip-to-tip dims (X-centered). sheet_elements from
    get_sheet_requirements gates optional tapers (Must include ShoulderTaper
    when the official sheet shows it). Core Roll Ahead / Vehicle Space /
    Buffer always. Dims are separate — place_order_table_dimensions does
    every tick and is not sheet-gated.

    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry then). force=True
    to override."""
    _refuse_if_spec_path_available("place_order_table_labels", force)
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_LABELS", alignIdx=align_idx,
                     outwardSign=outward_sign, textExtraAlong=text_extra_along,
                     sheetElements=sheet_elements),
        "place_order_table_labels")


def place_order_table_dimensions(align_idx: int, outward_sign: float = -1.0,
                                 offset_dist: float = 15.0,
                                 sheet_elements: str = "", force: bool = False) -> dict:
    """Real ny_Plan Linear Size dims tip-to-tip between EVERY consecutive
    tick (including Sign spacings). Length above the dim line.
    sheet_elements is not used for gating (API compat only).

    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry then). force=True
    to override."""
    _refuse_if_spec_path_available("place_order_table_dimensions", force)
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_DIMENSIONS", alignIdx=align_idx,
                     outwardSign=outward_sign, offsetDist=offset_dist,
                     sheetElements=sheet_elements),
        "place_order_table_dimensions")


def place_sheet_symbol_cells(align_idx: int, sheet_elements: str,
                             outward_sign: float = -1.0, force: bool = False) -> dict:
    """ProtectiveVehicle→TWZWVA_P in Vehicle Space bay (Buffer Space fallback
    when the sheet has no VS — e.g. 619-301); ArrowPanel→TWZAP_P at
    Shoulder Taper tip (fallback Merging/Lane taper).

    Generic heuristic (fixed offset, not lane/shoulder-width-derived), no
    rules-gate validation — refuses when a sheet spec exists for this build
    (prefer place_sheet_geometry, whose compile_symbols derives a real
    lateral position from lane/shoulder width). force=True to override."""
    _refuse_if_spec_path_available("place_sheet_symbol_cells", force)
    return _ok_or_raise(
        _bridge.call("PLACE_SHEET_SYMBOL_CELLS", alignIdx=align_idx,
                     sheetElements=sheet_elements, outwardSign=outward_sign),
        "place_sheet_symbol_cells")


def place_order_table_workspace(align_idx: int, outward_sign: float = -1.0,
                                lane_width: float = 12.0, force: bool = False) -> dict:
    """Hatched work-space box in the closed lane from path start through
    Vehicle Space end (Buffer Space end when the sheet has no VS). Not
    freeform vertices.

    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry, whose compile_hatch
    derives the work-area bounds from both alignments' own station-0 points
    and is checked by run_rules_gate's corridor-topology gate). force=True
    to override."""
    _refuse_if_spec_path_available("place_order_table_workspace", force)
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_WORKSPACE", alignIdx=align_idx,
                     outwardSign=outward_sign, laneWidth=lane_width),
        "place_order_table_workspace")


def place_order_table_channelizing(align_idx: int, outward_sign: float = -1.0,
                                   lane_width: float = 12.0, force: bool = False) -> dict:
    """Sheet-bounded channelizing: merging/lane taper diagonal (or shoulder
    taper alone on shoulder-only sheets) + longitudinal closed-lane run
    from taper toe to path start. Does not use freeform AccuDraw vertices.

    Generic heuristic, no rules-gate validation (taper-continuity/
    cone-spacing checks live in run_rules_gate, only reachable via
    compile_sheet_plan) — refuses when a sheet spec exists for this build
    (prefer place_sheet_geometry). force=True to override."""
    _refuse_if_spec_path_available("place_order_table_channelizing", force)
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_CHANNELIZING", alignIdx=align_idx,
                     outwardSign=outward_sign, laneWidth=lane_width),
        "place_order_table_channelizing")


def place_dimension(x1: float, y1: float, x2: float, y2: float,
                    ox: float, oy: float, z: float = 0.0,
                    style_name: str = "ny_Plan", reason: str = "",
                    override_text: str = "") -> dict:
    """Place a real Linear Size DimensionElement (msdDimTypeSizeArrow)
    between (x1,y1)-(x2,y2); dim-line offset toward (ox,oy). Uses DGN
    DimensionStyles (default ny_Plan). Prefer place_order_table_dimensions
    for full-plan spacing annotation.

    override_text: optional PrimaryText (sheet/table length). Empty =
    measured length (straight-sheet default).
    """
    kwargs = dict(x1=x1, y1=y1, x2=x2, y2=y2, ox=ox, oy=oy, z=z,
                  styleName=style_name, reason=reason)
    if override_text != "":
        kwargs["overrideText"] = override_text
    return _ok_or_raise(_bridge.call("PLACE_DIMENSION", **kwargs), "place_dimension")


def place_arc_size_dimension(cx: float, cy: float,
                             x1: float, y1: float, x2: float, y2: float,
                             ox: float, oy: float, z: float = 0.0,
                             style_name: str = "ny_Plan", reason: str = "",
                             override_text: str = "") -> dict:
    """One continuous curved DimensionElement (msdDimTypeArcSize) + ny_Plan.

    Prefer place_curved_plan_dimension for roadside-hugging bend dims —
    Arc Size often sweeps the wrong way on this install.
    """
    kwargs = dict(cx=cx, cy=cy, x1=x1, y1=y1, x2=x2, y2=y2, ox=ox, oy=oy, z=z,
                  styleName=style_name, reason=reason)
    if override_text != "":
        kwargs["overrideText"] = override_text
    return _ok_or_raise(
        _bridge.call("PLACE_ARC_SIZE_DIMENSION", **kwargs),
        "place_arc_size_dimension")


def place_curved_plan_dimension(cx: float, cy: float,
                                x1: float, y1: float, x2: float, y2: float,
                                ox: float, oy: float, z: float = 0.0,
                                reason: str = "",
                                override_text: str = "") -> dict:
    """Dim line = true ArcElement concentric with the bend (FOLLOWS THE CURVE).

    Radial extension lines + SizeArrow-style filled tips + sheet-length text.
    """
    kwargs = dict(cx=cx, cy=cy, x1=x1, y1=y1, x2=x2, y2=y2, ox=ox, oy=oy, z=z,
                  reason=reason)
    if override_text != "":
        kwargs["overrideText"] = override_text
    r = _ok_or_raise(
        _bridge.call("PLACE_CURVED_PLAN_DIMENSION", **kwargs),
        "place_curved_plan_dimension")
    tx = r.get("textX")
    ty = r.get("textY")
    txt = override_text or r.get("overrideText") or ""
    if tx is not None and ty is not None and str(txt).strip():
        try:
            tr = place_text_label(
                str(txt), float(tx), float(ty),
                reason=reason or "curved plan dim text")
            ids = placement_registry.parse_created_ids(r)
            ids.extend(placement_registry.parse_created_ids(tr))
            r = dict(r)
            r["createdElementIds"] = ids
            r["textElementIds"] = placement_registry.parse_created_ids(tr)
        except Exception as e:
            r = dict(r)
            r["textError"] = str(e)
    return r


def delete_dimension_elements_in_range(low_x: float, low_y: float,
                                       high_x: float, high_y: float,
                                       reason: str = "") -> dict:
    """Delete DimensionElements + color-2 arc/line/shape in bbox (leftover wipe).

    Engineer QA: remove bad chord/tip-chain dims AND prior curved-plan graphics
    before rebuild — DimensionElement-only wipe left arcs/chords visible.
    """
    return _ok_or_raise(
        _bridge.call(
            "DELETE_DIMENSION_ELEMENTS_IN_RANGE",
            lowX=low_x, lowY=low_y, highX=high_x, highY=high_y,
            reason=reason or "wipe leftover dims"),
        "delete_dimension_elements_in_range")


def arc_dim_line_radius(cx: float, cy: float, r: float,
                       mid: tuple[float, float],
                       ox: float, oy: float, pad: float = 15.0) -> float:
    """Radius of the Arc Size dim line.

    Keep the dim on the same roadside as (ox, oy). If the closed shoulder
    is inside the curve, ox is closer to the center than the tips — use
    r-pad (negative DimHeight). Never force r+pad; that draws through the
    pavement on the inside of an S-bend (live 2026-08-14).
    """
    import math
    vx, vy = mid[0] - cx, mid[1] - cy
    oxv, oyv = ox - cx, oy - cy
    r_ox = math.hypot(oxv, oyv)
    same = vx * oxv + vy * oyv
    if r < 1.0:
        return r + pad
    if same >= 0.0 and r_ox < r - 0.5:
        return max(1.0, min(r_ox, r - pad))
    return max(r + pad, r_ox)


def _format_ny_plan_dim_text(text: str) -> str:
    """Match straight-sheet ny_Plan look (e.g. 495'-0\") from lengthOnly '495'."""
    s = str(text or "").strip()
    if not s:
        return s
    if "'" in s or '"' in s:
        return s
    try:
        v = float(s)
        if abs(v - round(v)) < 1e-6:
            return f"{int(round(v))}'-0\""
        return f"{v:g}'"
    except ValueError:
        return s


def _fit_circle_2d(pts: list[tuple[float, float]],
                   min_sag: float = 1.0) -> tuple[float, float, float] | None:
    """Algebraic circle fit. Returns (cx, cy, r) or None if degenerate.

    Reject near-straight spans (tiny sagitta / huge R). Those must stay
    one SizeArrow — Arc Size with R>>chord draws a giant construction arc
    (live QA 2026-08-13 Buffer ~495' got center ~112k ft away).
    """
    import math
    if len(pts) < 3:
        return None
    # Prefer endpoints + mid for a stable roadside arc; fall back to all pts.
    sample = [pts[0], pts[len(pts) // 2], pts[-1]]
    (x1, y1), (x2, y2), (x3, y3) = sample
    d = 2.0 * (x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2))
    if abs(d) < 1e-9:
        return None
    ux = ((x1 * x1 + y1 * y1) * (y2 - y3)
          + (x2 * x2 + y2 * y2) * (y3 - y1)
          + (x3 * x3 + y3 * y3) * (y1 - y2)) / d
    uy = ((x1 * x1 + y1 * y1) * (x3 - x2)
          + (x2 * x2 + y2 * y2) * (x1 - x3)
          + (x3 * x3 + y3 * y3) * (x2 - x1)) / d
    r = math.hypot(x1 - ux, y1 - uy)
    if r < 1.0 or r > 1.0e7:
        return None
    chord = math.hypot(pts[-1][0] - pts[0][0], pts[-1][1] - pts[0][1])
    if chord < 1.0:
        return None
    half = min(chord * 0.5, r)
    sag = r - math.sqrt(max(0.0, r * r - half * half))
    if sag < min_sag:
        return None
    # Cap radius: allow highway curves (C-curve R~3000 on a 50' downstream
    # span). Reject only giant construction arcs (Buffer 495' → R~1e5).
    if r > max(12000.0, 80.0 * chord):
        return None
    return (ux, uy, r)


# Bowed spans normally use the constructed PLACE_CURVED_PLAN_DIMENSION
# (ArcElement + radial extensions + tip fans). Set True to place a REAL
# annotative DimensionElement (msdDimTypeArcSize) instead. Arc Size was
# thought broken on this install until 2026-08-13, when the actual defect
# turned out to be our own tip order — Arc Size measures counter-clockwise
# and we passed tips in path order, so clockwise bends swept the long way.
# Fixed in ExecPlaceArcSizeDimension; see scripts/diag_arc_size_root_cause.py
# (6/6 hug after fix). Left opt-in pending engineer visual QA that it matches
# the straight-sheet ny_Plan SizeArrow look.
ARC_SIZE_BEND_DIMS = True


def place_path_hugging_dimension(path: list, text: str,
                                 offset: list | tuple,
                                 reason: str = "",
                                 force_arc: bool = False) -> dict:
    """Curved-corridor dimension: ArcElement dim-line that FOLLOWS THE CURVE.

    Concentric with the tip-path bend + radial extensions + sheet length text.
    Falls back to one SizeArrow only when the tip path is essentially straight
    and force_arc is False.
    """
    import math
    if not path or len(path) < 2:
        raise ValueError("place_path_hugging_dimension needs >= 2 path points")
    verts: list[tuple[float, float]] = []
    for p in path:
        if isinstance(p, (list, tuple)) and len(p) >= 2:
            verts.append((float(p[0]), float(p[1])))
    cleaned: list[tuple[float, float]] = [verts[0]]
    for x, y in verts[1:]:
        if math.hypot(x - cleaned[-1][0], y - cleaned[-1][1]) >= 0.5:
            cleaned.append((x, y))
    if len(cleaned) < 2:
        raise ValueError("place_path_hugging_dimension: tip path too short after cleanup")

    ox = float(offset[0])
    oy = float(offset[1])
    sheet_txt = _format_ny_plan_dim_text(text)
    x1, y1 = cleaned[0]
    x2, y2 = cleaned[-1]

    # Ensure 3 non-collinear samples for circle fit on short spans.
    if len(cleaned) == 2:
        mx = 0.5 * (x1 + x2)
        my = 0.5 * (y1 + y2)
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy) or 1.0
        nx, ny = -dy / L, dx / L
        if (ox - mx) * nx + (oy - my) * ny < 0:
            nx, ny = -nx, -ny
        bulge = max(2.0, 0.05 * L)
        cleaned = [cleaned[0], (mx + nx * bulge, my + ny * bulge), cleaned[1]]

    min_sag = 0.01 if force_arc else 0.25
    fit = _fit_circle_2d(cleaned, min_sag=min_sag)
    if fit is None and force_arc and len(cleaned) >= 3:
        # Last resort: circumcircle of start/mid/end, but still reject
        # near-straight huge-R fits (Buffer on approach).
        sample = [cleaned[0], cleaned[len(cleaned) // 2], cleaned[-1]]
        (x1a, y1a), (x2a, y2a), (x3a, y3a) = sample
        d = 2.0 * (x1a * (y2a - y3a) + x2a * (y3a - y1a) + x3a * (y1a - y2a))
        if abs(d) >= 1e-9:
            ux = ((x1a * x1a + y1a * y1a) * (y2a - y3a)
                  + (x2a * x2a + y2a * y2a) * (y3a - y1a)
                  + (x3a * x3a + y3a * y3a) * (y1a - y2a)) / d
            uy = ((x1a * x1a + y1a * y1a) * (x3a - x2a)
                  + (x2a * x2a + y2a * y2a) * (x1a - x3a)
                  + (x3a * x3a + y3a * y3a) * (x2a - x1a)) / d
            rr = math.hypot(x1a - ux, y1a - uy)
            chord = math.hypot(x2 - x1, y2 - y1)
            half = min(chord * 0.5, rr) if rr > 0 else 0.0
            sag = rr - math.sqrt(max(0.0, rr * rr - half * half)) if rr > 0 else 0.0
            if 1.0 < rr <= max(12000.0, 80.0 * chord) and sag >= 0.03:
                fit = (ux, uy, rr)
    if fit is not None:
        cx, cy, r = fit
        chord = math.hypot(x2 - x1, y2 - y1)
        # Even with force_arc, do not place ArcElement for near-straight spans
        # (huge R) — those stay SizeArrow like the straight sheet.
        half = min(chord * 0.5, r)
        sag = r - math.sqrt(max(0.0, r * r - half * half))
        # force_arc: keep mild bows on the bend (50' Downstream) as arcs;
        # only reject near-straight / giant-R (Buffer approach).
        min_keep = 0.03 if force_arc else 0.35
        # Highway C-curves are R~3000; old cap 2500 forced 50' downstream
        # onto a SizeArrow chord. Reject only giant construction arcs.
        r_cap = max(12000.0, 80.0 * chord) if force_arc else max(1200.0, 25.0 * chord)
        if sag < min_keep or r > r_cap:
            fit = None
    if fit is not None:
        cx, cy, r = fit
        a1 = math.atan2(y1 - cy, x1 - cx)
        a2 = math.atan2(y2 - cy, x2 - cx)
        da = (a2 - a1 + math.pi) % (2.0 * math.pi) - math.pi
        amid = a1 + 0.5 * da
        mid = cleaned[len(cleaned) // 2]
        r_off = arc_dim_line_radius(cx, cy, r, mid, ox, oy, pad=15.0)
        hx = cx + r_off * math.cos(amid)
        hy = cy + r_off * math.sin(amid)
        if ARC_SIZE_BEND_DIMS:
            resp = place_arc_size_dimension(
                cx, cy, x1, y1, x2, y2, hx, hy,
                reason=reason or f"curve arc-size dim {text}",
                override_text=sheet_txt)
            ids = placement_registry.parse_created_ids(
                resp if isinstance(resp, dict) else {})
            return {
                "status": "OK",
                "curved": True,
                "dimType": "ArcSize",
                "text": sheet_txt,
                "center": [cx, cy],
                "createdElementIds": ids,
                "note": ("curved dim = real annotative DimensionElement "
                         "(msdDimTypeArcSize, ny_Plan) concentric with bend"),
            }
        resp = place_curved_plan_dimension(
            cx, cy, x1, y1, x2, y2, hx, hy,
            reason=reason or f"curve arc dim {text}",
            override_text=sheet_txt)
        ids = placement_registry.parse_created_ids(
            resp if isinstance(resp, dict) else {})
        return {
            "status": "OK",
            "curved": True,
            "dimType": "CurvedPlanArc",
            "text": sheet_txt,
            "center": [cx, cy],
            "createdElementIds": ids,
            "note": (
                "curved dim = ArcElement dim-line concentric with bend "
                "(follows the curve) + radial extensions + sheet length text"
            ),
        }

    # Essentially straight tip path — one SizeArrow (matches straight sheet).
    r = place_dimension(
        x1, y1, x2, y2, ox, oy,
        reason=reason or f"curve fallback size-arrow {text}",
        override_text=sheet_txt)
    ids = placement_registry.parse_created_ids(r if isinstance(r, dict) else {})
    return {
        "status": "OK",
        "curved": False,
        "dimType": "SizeArrow",
        "text": sheet_txt,
        "createdElementIds": ids,
        "note": "tip path nearly straight; one ny_Plan SizeArrow",
    }


def hatch_element(element_id: str, spacing: float = 10.0, angle_deg: float = 45.0,
                  own_element_only: bool = True, reason: str = "") -> dict:
    """Apply associative hatch to an existing closed shape by element ID.
    Does not create a new element — sets HasPattern on the shape. spacing
    is in master units; angle_deg in degrees."""
    return _ok_or_raise(
        _bridge.call("HATCH_ELEMENT", elementId=element_id, spacing=spacing,
                     angleDeg=angle_deg, ownElementOnly=own_element_only, reason=reason),
        "hatch_element")


def place_arc(x1: float, y1: float, x2: float, y2: float, x3: float, y3: float,
              z: float = 0.0, reason: str = "") -> dict:
    """Place a 3-point arc (placeArcModeEx=3). Point order: start, end, bulge."""
    return _ok_or_raise(
        _bridge.call("PLACE_ARC", x1=x1, y1=y1, x2=x2, y2=y2, x3=x3, y3=y3, z=z, reason=reason),
        "place_arc")


def place_text_label(text: str, x: float, y: float, z: float = 0.0,
                     reason: str = "", angle_deg: float = 0.0) -> dict:
    """Place a single-line text label. angle_deg rotates about Z (tangent
    for Non-Sign dim labels on a curve; 0 = view-identity / +X)."""
    return _ok_or_raise(
        _bridge.call("PLACE_TEXT_LABEL", text=text, x=x, y=y, z=z,
                     angleDeg=angle_deg, reason=reason),
        "place_text_label")


def place_circle(cx: float, cy: float, radius: float, z: float = 0.0, reason: str = "") -> dict:
    """Place a circle (CreateEllipseElement2 with equal radii)."""
    return _ok_or_raise(_bridge.call("PLACE_CIRCLE", cx=cx, cy=cy, radius=radius, z=z, reason=reason), "place_circle")


def place_ellipse(cx: float, cy: float, primary_radius: float, secondary_radius: float,
                  angle_deg: float = 0.0, z: float = 0.0, reason: str = "") -> dict:
    """Place an ellipse via CreateEllipseElement2."""
    return _ok_or_raise(
        _bridge.call("PLACE_ELLIPSE", cx=cx, cy=cy, primaryRadius=primary_radius,
                     secondaryRadius=secondary_radius, angleDeg=angle_deg, z=z, reason=reason),
        "place_ellipse")


def place_block(x1: float, y1: float, x2: float, y2: float, z: float = 0.0, reason: str = "") -> dict:
    """Place an axis-aligned rectangle (CreateShapeElement1)."""
    return _ok_or_raise(_bridge.call("PLACE_BLOCK", x1=x1, y1=y1, x2=x2, y2=y2, z=z, reason=reason), "place_block")


def place_polyline(vertices: list[list[float]], reason: str = "") -> dict:
    """Place an open polyline. vertices is [[x,y,z?], ...]."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(_bridge.call("PLACE_POLYLINE", verticesTSV=verts_tsv, reason=reason), "place_polyline")


def _place_road_line_segments(
    segs: list[dict], *, reason_prefix: str, need_yellow: bool = True,
) -> tuple[list[dict], list[str], int | None]:
    """Place striping segments with Default/weight0; yellow kinds use resolve_color.

    Each seg may be a 2-point line (x1,y1,x2,y2) or a curved/S polyline via
    ``vertices=[[x,y],…]`` (from path-offset highway builders).
    """
    yellow_idx: int | None = None
    if need_yellow and any((s.get("kind") or "") == "yellow" for s in segs):
        yc = resolve_color(name="yellow")
        if yc.get("status") != "OK" or yc.get("index") is None:
            return [], [f"resolve_color('yellow') failed: {yc.get('note') or yc}"], None
        yellow_idx = int(yc["index"])

    placed: list[dict] = []
    errors: list[str] = []
    for seg in segs:
        if (seg.get("style") or "") == "meta":
            continue
        kind = seg.get("kind") or "lane"
        try:
            verts = seg.get("vertices")
            if verts and len(verts) >= 2:
                poly = [[float(p[0]), float(p[1]), 0.0] for p in verts]
            else:
                poly = [
                    [float(seg["x1"]), float(seg["y1"]), 0.0],
                    [float(seg["x2"]), float(seg["y2"]), 0.0],
                ]
            r = place_polyline(
                poly,
                reason=f"{reason_prefix} {kind} {seg['style']} row={seg['row']}",
            )
            eid = str(r.get("elementId") or "")
            color = yellow_idx if kind == "yellow" and yellow_idx is not None else 0
            if eid:
                change_element_level(eid, "Default", own_element_only=True,
                                     reason="road strip align-like level")
                change_element_symbology(eid, color=color, weight=0, own_element_only=True,
                                         reason="road strip color/weight")
            placed.append({
                "elementId": eid, "style": seg["style"], "kind": kind,
                "row": seg["row"],
                "arm": seg.get("arm") or "",
                "x1": float(poly[0][0]), "y1": float(poly[0][1]),
                "x2": float(poly[-1][0]), "y2": float(poly[-1][1]),
                "vertexCount": len(poly),
                "color": color,
            })
        except Exception as e:
            errors.append(str(e))
    return placed, errors, yellow_idx


def _road_strip_counts(placed: list[dict]) -> dict:
    return {
        "solidWhiteCount": sum(
            1 for p in placed
            if p["style"] == "solid" and p["kind"] in ("edge", "shoulder", "gore")
        ),
        "solidYellowCount": sum(
            1 for p in placed if p["style"] == "solid" and p["kind"] == "yellow"
        ),
        "dashedYellowSegmentCount": sum(
            1 for p in placed if p["style"] == "dashed" and p["kind"] == "yellow"
        ),
        "dashedSegmentCount": sum(1 for p in placed if p["style"] == "dashed"),
        "shoulderCount": sum(1 for p in placed if p["kind"] == "shoulder"),
        "goreMarkCount": sum(1 for p in placed if p["kind"] == "gore"),
        "stopBarCount": sum(1 for p in placed if p["kind"] == "stop_bar"),
        "crosswalkCount": sum(1 for p in placed if p["kind"] == "crosswalk"),
        "placedCount": len(placed),
    }


def _resolve_road_edge_args(
    x1: float, y1: float, x2: float, y2: float,
    vertices: list | None,
) -> tuple[float, float, float, float, list | None, float]:
    """Return (x1,y1,x2,y2,vertices,lengthFt). vertices wins when len>=2."""
    import lane_highway as lh

    if vertices is not None and len(vertices) >= 2:
        segs = lh.vertices_to_path_segments(vertices)
        length = lh.path_length(segs)
        first, last = vertices[0], vertices[-1]
        return (
            float(first[0]), float(first[1]),
            float(last[0]), float(last[1]),
            [[float(p[0]), float(p[1])] for p in vertices],
            length,
        )
    length = ((float(x2) - float(x1)) ** 2 + (float(y2) - float(y1)) ** 2) ** 0.5
    return float(x1), float(y1), float(x2), float(y2), None, length


def place_lane_highway(lanes: int, x1: float = 0.0, y1: float = 0.0,
                       x2: float = 0.0, y2: float = 0.0,
                       lane_width_ft: float = 12.0, shoulder_width_ft: float = 0.0,
                       dash_ft: float = 10.0, gap_ft: float = 30.0,
                       side: str = "right", reason: str = "",
                       vertices: list | None = None) -> dict:
    """Draw an N-lane one-way highway strip (general CAD).

    Two solid white outer travel edges + (lanes-1) dashed separators.
    Optional shoulder_width_ft > 0 adds solid white EOP lines outside both
    travel outers (sheet 'paved shoulder'). Pass vertices=[[x,y],…] for a
    curved/S first-travel-outer polyline (overrides x1..y2). Ask for missing
    inputs — do not invent site coordinates. Not a 619 sheet-plan tool.
    """
    import lane_highway as lh

    side_n = (side or "right").strip().lower()
    try:
        x1, y1, x2, y2, verts, length = _resolve_road_edge_args(
            x1, y1, x2, y2, vertices,
        )
        segs = lh.lane_highway_lines(
            int(lanes), float(x1), float(y1), float(x2), float(y2),
            lane_width_ft=float(lane_width_ft),
            shoulder_width_ft=float(shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            side=side_n,  # type: ignore[arg-type]
            vertices=verts,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    placed, errors, _ = _place_road_line_segments(
        segs, reason_prefix=reason or f"one-way {lanes}-lane", need_yellow=False,
    )
    counts = _road_strip_counts(placed)
    path_note = f"; pathVerts={len(verts)}" if verts else ""
    _remember_placed_road(
        road_type="one_way", lanes=int(lanes),
        lane_width_ft=float(lane_width_ft),
        shoulder_width_ft=float(shoulder_width_ft),
        yellow_gap_ft=0.0, side=side_n,
        verts=verts, x1=x1, y1=y1, x2=x2, y2=y2, length=length,
    )
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "one_way",
        "lanes": int(lanes),
        "lengthFt": round(length, 3),
        "pathVertexCount": len(verts) if verts else 2,
        "laneWidthFt": float(lane_width_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "dashFt": float(dash_ft),
        "gapFt": float(gap_ft),
        "side": side_n,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"{lanes}-lane one-way: 2 travel edges + {max(lanes - 1, 0)} dashed "
            f"row(s); shoulder={shoulder_width_ft}ft; dash={dash_ft}/{gap_ft}ft"
            f"{path_note}"
        ),
    }


def place_two_way_highway(lanes: int, x1: float = 0.0, y1: float = 0.0,
                          x2: float = 0.0, y2: float = 0.0,
                          lane_width_ft: float = 12.0, yellow_gap_ft: float = 2.0,
                          shoulder_width_ft: float = 0.0,
                          dash_ft: float = 10.0, gap_ft: float = 30.0,
                          side: str = "right", reason: str = "",
                          vertices: list | None = None) -> dict:
    """Draw an even-N undivided two-way road (double solid yellow center).

    Optional shoulder_width_ft. Pass vertices=[[x,y],…] for a curved/S
    first-travel-outer polyline (overrides x1..y2). Ask for missing
    lanes/width/endpoints/side/path.
    """
    import lane_highway as lh

    side_n = (side or "right").strip().lower()
    try:
        x1, y1, x2, y2, verts, length = _resolve_road_edge_args(
            x1, y1, x2, y2, vertices,
        )
        segs = lh.two_way_highway_lines(
            int(lanes), float(x1), float(y1), float(x2), float(y2),
            lane_width_ft=float(lane_width_ft),
            yellow_gap_ft=float(yellow_gap_ft),
            shoulder_width_ft=float(shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            side=side_n,  # type: ignore[arg-type]
            vertices=verts,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    placed, errors, yellow_idx = _place_road_line_segments(
        segs, reason_prefix=reason or f"two-way {lanes}-lane",
    )
    if errors and yellow_idx is None and not placed:
        return {"status": "ERROR", "note": errors[0], "errors": errors}

    per_dir = int(lanes) // 2
    counts = _road_strip_counts(placed)
    path_note = f"; pathVerts={len(verts)}" if verts else ""
    _remember_placed_road(
        road_type="two_way_undivided", lanes=int(lanes),
        lane_width_ft=float(lane_width_ft),
        shoulder_width_ft=float(shoulder_width_ft),
        yellow_gap_ft=float(yellow_gap_ft), side=side_n,
        verts=verts, x1=x1, y1=y1, x2=x2, y2=y2, length=length,
    )
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "two_way_undivided",
        "lanes": int(lanes),
        "lanesPerDirection": per_dir,
        "lengthFt": round(length, 3),
        "pathVertexCount": len(verts) if verts else 2,
        "laneWidthFt": float(lane_width_ft),
        "yellowGapFt": float(yellow_gap_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "yellowColorIndex": yellow_idx,
        "dashFt": float(dash_ft),
        "gapFt": float(gap_ft),
        "side": side_n,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"{lanes}-lane two-way undivided: double yellow gap={yellow_gap_ft}ft; "
            f"shoulder={shoulder_width_ft}ft; yellow idx={yellow_idx}{path_note}"
        ),
    }


def place_divided_highway(lanes_per_direction: int, x1: float = 0.0,
                          y1: float = 0.0, x2: float = 0.0, y2: float = 0.0,
                          median_width_ft: float = 0.0,
                          lane_width_ft: float = 12.0,
                          shoulder_width_ft: float = 0.0,
                          dash_ft: float = 10.0, gap_ft: float = 30.0,
                          side: str = "right", reason: str = "",
                          vertices: list | None = None) -> dict:
    """Draw a divided multilane / freeway dual carriageway (619-302-style).

    Each direction: white outer, (N-1) dashed white, yellow median edge.
    median_width_ft is the empty gap between the two yellows (required —
    ask; do not invent). Optional outer shoulders. Pass vertices=[[x,y],…]
    for a curved/S first-travel-outer polyline (overrides x1..y2).
    """
    import lane_highway as lh

    side_n = (side or "right").strip().lower()
    try:
        x1, y1, x2, y2, verts, length = _resolve_road_edge_args(
            x1, y1, x2, y2, vertices,
        )
        segs = lh.divided_highway_lines(
            int(lanes_per_direction), float(x1), float(y1), float(x2), float(y2),
            median_width_ft=float(median_width_ft),
            lane_width_ft=float(lane_width_ft),
            shoulder_width_ft=float(shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            side=side_n,  # type: ignore[arg-type]
            vertices=verts,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    placed, errors, yellow_idx = _place_road_line_segments(
        segs,
        reason_prefix=reason or f"divided {lanes_per_direction}+{lanes_per_direction}",
    )
    if errors and yellow_idx is None and not placed:
        return {"status": "ERROR", "note": errors[0], "errors": errors}

    counts = _road_strip_counts(placed)
    path_note = f"; pathVerts={len(verts)}" if verts else ""
    _remember_placed_road(
        road_type="divided", lanes=int(lanes_per_direction) * 2,
        lane_width_ft=float(lane_width_ft),
        shoulder_width_ft=float(shoulder_width_ft),
        yellow_gap_ft=0.0, side=side_n,
        verts=verts, x1=x1, y1=y1, x2=x2, y2=y2, length=length,
    )
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "divided",
        "lanesPerDirection": int(lanes_per_direction),
        "lengthFt": round(length, 3),
        "pathVertexCount": len(verts) if verts else 2,
        "medianWidthFt": float(median_width_ft),
        "laneWidthFt": float(lane_width_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "yellowColorIndex": yellow_idx,
        "dashFt": float(dash_ft),
        "gapFt": float(gap_ft),
        "side": side_n,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"divided {lanes_per_direction}/dir + median {median_width_ft}ft; "
            f"shoulder={shoulder_width_ft}ft; yellow idx={yellow_idx}{path_note}"
        ),
    }


def place_twlt_highway(lanes_per_direction: int, x1: float = 0.0,
                       y1: float = 0.0, x2: float = 0.0, y2: float = 0.0,
                       twlt_width_ft: float = 12.0,
                       lane_width_ft: float = 12.0,
                       shoulder_width_ft: float = 0.0,
                       dash_ft: float = 10.0, gap_ft: float = 30.0,
                       side: str = "right", reason: str = "",
                       vertices: list | None = None) -> dict:
    """Draw multilane undivided with center TWLT (619-312-style).

    lanes_per_direction = travel lanes each way (TWLT not counted).
    Center turn lane bounded by two dashed yellow lines twlt_width_ft
    apart. Pass vertices=[[x,y],…] for a curved/S first-travel-outer
    polyline (overrides x1..y2). Do NOT use place_two_way_highway for TWLT.
    """
    import lane_highway as lh

    side_n = (side or "right").strip().lower()
    try:
        x1, y1, x2, y2, verts, length = _resolve_road_edge_args(
            x1, y1, x2, y2, vertices,
        )
        segs = lh.twlt_highway_lines(
            int(lanes_per_direction), float(x1), float(y1), float(x2), float(y2),
            twlt_width_ft=float(twlt_width_ft),
            lane_width_ft=float(lane_width_ft),
            shoulder_width_ft=float(shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            side=side_n,  # type: ignore[arg-type]
            vertices=verts,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    placed, errors, yellow_idx = _place_road_line_segments(
        segs,
        reason_prefix=reason or f"twlt {lanes_per_direction}+TWLT+{lanes_per_direction}",
    )
    if errors and yellow_idx is None and not placed:
        return {"status": "ERROR", "note": errors[0], "errors": errors}

    counts = _road_strip_counts(placed)
    path_note = f"; pathVerts={len(verts)}" if verts else ""
    _remember_placed_road(
        road_type="twlt", lanes=int(lanes_per_direction) * 2,
        lane_width_ft=float(lane_width_ft),
        shoulder_width_ft=float(shoulder_width_ft),
        yellow_gap_ft=float(twlt_width_ft), side=side_n,
        verts=verts, x1=x1, y1=y1, x2=x2, y2=y2, length=length,
    )
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "twlt",
        "lanesPerDirection": int(lanes_per_direction),
        "lengthFt": round(length, 3),
        "pathVertexCount": len(verts) if verts else 2,
        "twltWidthFt": float(twlt_width_ft),
        "laneWidthFt": float(lane_width_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "yellowColorIndex": yellow_idx,
        "dashFt": float(dash_ft),
        "gapFt": float(gap_ft),
        "side": side_n,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"TWLT: {lanes_per_direction}/dir + center {twlt_width_ft}ft "
            f"(dashed yellow bounds); shoulder={shoulder_width_ft}ft{path_note}"
        ),
    }


def place_orthogonal_intersection(
    junction_x: float, junction_y: float,
    primary_road_type: str, secondary_road_type: str,
    primary_length_ft: float, secondary_stub_ft: float,
    primary_bearing_deg: float = 0.0,
    junction: str = "plus", tee_side: str = "right",
    primary_lanes: int | None = None, secondary_lanes: int | None = None,
    primary_lanes_per_direction: int | None = None,
    secondary_lanes_per_direction: int | None = None,
    lane_width_ft: float = 12.0, yellow_gap_ft: float = 2.0,
    primary_median_width_ft: float = 0.0,
    secondary_median_width_ft: float = 0.0,
    primary_twlt_width_ft: float = 12.0,
    secondary_twlt_width_ft: float = 12.0,
    primary_shoulder_width_ft: float = 0.0,
    secondary_shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0, gap_ft: float = 30.0,
    side: str = "right",
    crosswalks: bool = True, stop_bars: bool = True,
    has_turning_lanes: bool | None = None,
    turn_arrows: bool = True,
    primary_lanes_out: int | None = None,
    secondary_lanes_out: int | None = None,
    reason: str = "",
) -> dict:
    """Draw a + or T intersection with MUTCD box striping rules.

    Edge lines meet the box (arms connect). Yellow/dashed lane lines stop
    at the stop bar. Defaults: crosswalks + stop bars on every approach;
    turn arrows from ny_plan_striping.cel (SAS through; SAL/SAR + SLONLY
    only when lanes_in > lanes_out via primary_lanes_out /
    secondary_lanes_out). Dotted yellow center when has_turning_lanes,
    TWLT, or dedicated > 0. Ask for missing inputs.
    """
    import road_junctions as rj

    side_n = (side or "right").strip().lower()
    try:
        segs = rj.orthogonal_intersection_lines(
            float(junction_x), float(junction_y),
            primary_road_type=primary_road_type,
            secondary_road_type=secondary_road_type,
            primary_length_ft=float(primary_length_ft),
            secondary_stub_ft=float(secondary_stub_ft),
            primary_bearing_deg=float(primary_bearing_deg),
            junction=junction,  # type: ignore[arg-type]
            tee_side=tee_side,  # type: ignore[arg-type]
            primary_lanes=primary_lanes,
            secondary_lanes=secondary_lanes,
            primary_lanes_per_direction=primary_lanes_per_direction,
            secondary_lanes_per_direction=secondary_lanes_per_direction,
            lane_width_ft=float(lane_width_ft),
            yellow_gap_ft=float(yellow_gap_ft),
            primary_median_width_ft=float(primary_median_width_ft),
            secondary_median_width_ft=float(secondary_median_width_ft),
            primary_twlt_width_ft=float(primary_twlt_width_ft),
            secondary_twlt_width_ft=float(secondary_twlt_width_ft),
            primary_shoulder_width_ft=float(primary_shoulder_width_ft),
            secondary_shoulder_width_ft=float(secondary_shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            side=side_n,  # type: ignore[arg-type]
            crosswalks=bool(crosswalks),
            stop_bars=bool(stop_bars),
            has_turning_lanes=has_turning_lanes,
            turn_arrows=bool(turn_arrows),
            primary_lanes_out=primary_lanes_out,
            secondary_lanes_out=secondary_lanes_out,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    arrow_metas = [s for s in segs if s.get("kind") == "turn_arrow"]
    placeable = rj.strip_placeable_segments(segs)
    placed, errors, yellow_idx = _place_road_line_segments(
        placeable,
        reason_prefix=reason or f"intersection {junction} {primary_road_type}/{secondary_road_type}",
    )

    arrows_placed: list[dict] = []
    for meta in arrow_metas:
        try:
            r = place_cell(
                str(meta["cellName"]),
                float(meta["x"]), float(meta["y"]), 0.0,
                angle_deg=float(meta.get("angleDeg") or 0.0),
                library_path=str(meta.get("libraryPath") or rj.DEFAULT_STRIPING_CELL_LIB),
                reason=reason or f"intersection turn arrow {meta['cellName']}",
            )
            arrows_placed.append({
                "elementId": str(r.get("elementId") or ""),
                "cellName": meta["cellName"],
                "arm": meta.get("arm"),
                "x": meta["x"], "y": meta["y"],
                "angleDeg": meta.get("angleDeg"),
            })
        except Exception as e:
            errors.append(f"turn_arrow {meta.get('cellName')}: {e}")

    if errors and not placed and not arrows_placed:
        return {"status": "ERROR", "note": errors[0], "errors": errors}

    arms = sorted({p.get("arm") or "" for p in placed if p.get("arm")})
    turning = any(
        (p.get("arm") or "").startswith("center_extension") for p in placed
    )
    counts = _road_strip_counts(placed)
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "orthogonal_intersection",
        "junction": (junction or "plus").strip().lower(),
        "teeSide": (tee_side or "right").strip().lower(),
        "junctionX": float(junction_x),
        "junctionY": float(junction_y),
        "primaryRoadType": primary_road_type,
        "secondaryRoadType": secondary_road_type,
        "primaryLengthFt": float(primary_length_ft),
        "secondaryStubFt": float(secondary_stub_ft),
        "primaryBearingDeg": float(primary_bearing_deg),
        "crosswalks": bool(crosswalks),
        "stopBars": bool(stop_bars),
        "turnArrows": bool(turn_arrows),
        "dottedCenterExtension": turning,
        "yellowColorIndex": yellow_idx,
        "arms": arms,
        "arrowCount": len(arrows_placed),
        "arrows": arrows_placed,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"{junction} intersection @ ({junction_x},{junction_y}): "
            f"edges-to-box; center/lane stop at stop-bar; "
            f"CW={crosswalks} SB={stop_bars} arrows={len(arrows_placed)} "
            f"dottedCenter={turning}"
        ),
    }


def place_ramp_gore(
    x1: float = 0.0, y1: float = 0.0, x2: float = 0.0, y2: float = 0.0,
    mainline_lanes: int = 2, ramp_angle_deg: float = 15.0,
    gore_station_ft: float = 0.0, ramp_length_ft: float = 200.0,
    ramp_lanes: int = 1, side: str = "right",
    gore_mark_ft: float = 40.0,
    lane_width_ft: float = 12.0, shoulder_width_ft: float = 0.0,
    dash_ft: float = 10.0, gap_ft: float = 30.0,
    reason: str = "",
    vertices: list | None = None,
) -> dict:
    """Draw mainline one-way + diverging ramp meeting at a gore nose.

    (x1,y1)->(x2,y2) or vertices=[[x,y],…] = mainline first travel outer
    edge. Gore nose on the ramp-side outer edge at gore_station_ft from
    start; ramp diverges by ramp_angle_deg toward side. Optional solid
    white gore V marks (gore_mark_ft). Ask for missing angle/station/
    lengths — do not invent.
    """
    import road_junctions as rj

    side_n = (side or "right").strip().lower()
    try:
        x1, y1, x2, y2, verts, length = _resolve_road_edge_args(
            x1, y1, x2, y2, vertices,
        )
        segs = rj.ramp_gore_lines(
            float(x1), float(y1), float(x2), float(y2),
            mainline_lanes=int(mainline_lanes),
            ramp_angle_deg=float(ramp_angle_deg),
            gore_station_ft=float(gore_station_ft),
            ramp_length_ft=float(ramp_length_ft),
            ramp_lanes=int(ramp_lanes),
            side=side_n,  # type: ignore[arg-type]
            gore_mark_ft=float(gore_mark_ft),
            lane_width_ft=float(lane_width_ft),
            shoulder_width_ft=float(shoulder_width_ft),
            dash_ft=float(dash_ft), gap_ft=float(gap_ft),
            vertices=verts,
        )
    except ValueError as e:
        return {"status": "ERROR", "note": str(e)}

    nose = next((s for s in segs if s.get("kind") == "gore_nose"), None)
    placeable = rj.strip_placeable_segments(segs)
    placed, errors, yellow_idx = _place_road_line_segments(
        placeable,
        reason_prefix=reason or f"ramp gore {mainline_lanes}+{ramp_lanes}",
        need_yellow=False,
    )
    if errors and not placed:
        return {"status": "ERROR", "note": errors[0], "errors": errors}

    counts = _road_strip_counts(placed)
    path_note = f"; pathVerts={len(verts)}" if verts else ""
    return {
        "status": "OK" if not errors else "ERROR",
        "roadType": "ramp_gore",
        "mainlineLanes": int(mainline_lanes),
        "rampLanes": int(ramp_lanes),
        "mainlineLengthFt": round(length, 3),
        "pathVertexCount": len(verts) if verts else 2,
        "rampLengthFt": float(ramp_length_ft),
        "rampAngleDeg": float(ramp_angle_deg),
        "goreStationFt": float(gore_station_ft),
        "goreMarkFt": float(gore_mark_ft),
        "goreNose": (
            {"x": nose["x1"], "y": nose["y1"]} if nose else None
        ),
        "laneWidthFt": float(lane_width_ft),
        "shoulderWidthFt": float(shoulder_width_ft),
        "side": side_n,
        "yellowColorIndex": yellow_idx,
        **counts,
        "placed": placed,
        "errors": errors,
        "note": (
            f"ramp gore: mainline {mainline_lanes}-lane + ramp {ramp_lanes}-lane "
            f"@ {ramp_angle_deg}deg station {gore_station_ft}ft{path_note}"
        ),
    }


def place_polygon(cx: float, cy: float, radius: float, sides: int, z: float = 0.0, reason: str = "") -> dict:
    """Place a regular n-gon centered at (cx,cy)."""
    return _ok_or_raise(
        _bridge.call("PLACE_POLYGON", cx=cx, cy=cy, radius=radius, sides=sides, z=z, reason=reason),
        "place_polygon")


def change_element_symbology(element_id: str, color: int | None = None, weight: int | None = None,
                             line_style_index: int | None = None, line_style_name: str = "",
                             own_element_only: bool = True, reason: str = "") -> dict:
    """Set element color and/or line weight (and optional line style).
    color is a MicroStation color-table INDEX for this DGN — not a universal
    hue. When the engineer names a color ('orange', 'yellow'), call
    resolve_color first and use the returned index; never guess (confirmed
    live: guessing 3 for orange painted the element red).

    For line style prefer line_style_name= from resolve_line_style (exact
    Name key like '( Dashed )') — the Number property is NOT a valid
    LineStyles() lookup key. line_style_index remains the 1-based
    collectionIndex fallback. ByLevel cannot be assigned this way."""
    params = {"elementId": element_id, "ownElementOnly": ("Y" if own_element_only else "N"), "reason": reason}
    if color is not None:
        params["color"] = color
    if weight is not None:
        params["weight"] = weight
    ls_name = (line_style_name or "").strip()
    if ls_name:
        params["lineStyleName"] = ls_name
    if line_style_index is not None:
        params["lineStyleIndex"] = line_style_index
    return _ok_or_raise(_bridge.call("CHANGE_ELEMENT_SYMBOLOGY", **params), "change_element_symbology")


def copy_parallel(element_id: str, distance: float, own_element_only: bool = True, reason: str = "") -> dict:
    """Perpendicular offset-copy of a LINE. distance>0 = left of start->end."""
    return _ok_or_raise(
        _bridge.call("COPY_PARALLEL", elementId=element_id, distance=distance,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "copy_parallel")


def crosshatch_element(element_id: str, spacing: float = 10.0, angle_deg: float = 45.0,
                       own_element_only: bool = True, reason: str = "") -> dict:
    """Apply crosshatch pattern to a closed element."""
    return _ok_or_raise(
        _bridge.call("CROSSHATCH_ELEMENT", elementId=element_id, spacing=spacing, angleDeg=angle_deg,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "crosshatch_element")


def remove_hatch(element_id: str, own_element_only: bool = True, reason: str = "") -> dict:
    """Remove associative hatch/pattern from a closed element."""
    return _ok_or_raise(
        _bridge.call("REMOVE_HATCH", elementId=element_id,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "remove_hatch")


def break_line(element_id: str, x: float, y: float, z: float = 0.0,
               own_element_only: bool = True, reason: str = "") -> dict:
    """Break a line into two segments at (x,y)."""
    return _ok_or_raise(
        _bridge.call("BREAK_LINE", elementId=element_id, x=x, y=y, z=z,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "break_line")


def extend_line(element_id: str, new_length: float, own_element_only: bool = True, reason: str = "") -> dict:
    """Set line length from start point (extend or shorten)."""
    return _ok_or_raise(
        _bridge.call("EXTEND_LINE", elementId=element_id, newLength=new_length,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "extend_line")


def fillet_elements(element_id1: str, element_id2: str, radius: float,
                    pick_x: float, pick_y: float, pick_z: float = 0.0,
                    own_element_only: bool = True, reason: str = "") -> dict:
    """Create a fillet arc between two elements (sources not auto-trimmed)."""
    return _ok_or_raise(
        _bridge.call("FILLET_ELEMENTS", elementId1=element_id1, elementId2=element_id2,
                     radius=radius, pickX=pick_x, pickY=pick_y, pickZ=pick_z,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "fillet_elements")


def create_complex_string(element_ids: list[str], reason: str = "") -> dict:
    """Create a complex string from existing chainable element IDs."""
    return _ok_or_raise(
        _bridge.call("CREATE_COMPLEX_STRING", elementIds=",".join(element_ids), reason=reason),
        "create_complex_string")


def place_fence_block(x1: float, y1: float, x2: float, y2: float, z: float = 0.0,
                      view_num: int = 1, reason: str = "") -> dict:
    """Define a rectangular fence from corner points."""
    return _ok_or_raise(
        _bridge.call("PLACE_FENCE_BLOCK", x1=x1, y1=y1, x2=x2, y2=y2, z=z, viewNum=view_num, reason=reason),
        "place_fence_block")


def fence_undefine(reason: str = "") -> dict:
    """Clear the current fence definition."""
    return _ok_or_raise(_bridge.call("FENCE_UNDEFINE", reason=reason), "fence_undefine")


def fence_copy_contents(delta_x: float, delta_y: float, delta_z: float = 0.0, reason: str = "") -> dict:
    """Clone+Move every element inside the current fence."""
    return _ok_or_raise(
        _bridge.call("FENCE_COPY_CONTENTS", deltaX=delta_x, deltaY=delta_y, deltaZ=delta_z, reason=reason),
        "fence_copy_contents")


def fence_move_contents(delta_x: float, delta_y: float, delta_z: float = 0.0, reason: str = "") -> dict:
    """Move every element inside the current fence."""
    return _ok_or_raise(
        _bridge.call("FENCE_MOVE_CONTENTS", deltaX=delta_x, deltaY=delta_y, deltaZ=delta_z, reason=reason),
        "fence_move_contents")


def fence_delete_contents(reason: str = "") -> dict:
    """Delete every element inside the current fence (not undoable)."""
    return _ok_or_raise(_bridge.call("FENCE_DELETE_CONTENTS", reason=reason), "fence_delete_contents")


def select_element(element_id: str, clear_first: bool = True, reason: str = "") -> dict:
    """Add an element to the selection set (optionally clearing first)."""
    return _ok_or_raise(
        _bridge.call("SELECT_ELEMENT", elementId=element_id,
                     clearFirst=("Y" if clear_first else "N"), reason=reason),
        "select_element")


def clear_selection(reason: str = "") -> dict:
    """Clear the model selection set."""
    return _ok_or_raise(_bridge.call("CLEAR_SELECTION", reason=reason), "clear_selection")


def place_element_run(element_idx: int, vertices: list[list[float]], reason: str = "") -> dict:
    """Place a channelizing-device / removal-striping / barrier run.
    element_idx: 2=Channelizing Devices, 3=Removal Striping, 4=Temporary
    Barrier, 5=Barrier w/Warning Lights (1=Work Space — use place_workspace
    instead). vertices is an ordered list of [x, y, z] points."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(
        _bridge.call("PLACE_ELEMENT_RUN", elementIdx=element_idx, verticesTSV=verts_tsv, reason=reason),
        "place_element_run")


def place_channelizing_markers(vertices: list[list[float]], half_size_ft: float = 1.5,
                                reason: str = "") -> dict:
    """Place discrete small orange squares at each cone center (TWZCD_P,
    color 6, solid linestyle). Prefer this over place_element_run for
    sheet-compiled channelizing — polylines on TWZCD_P pick up a custom
    ByLevel linestyle that reads as a solid orange wash."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(
        _bridge.call("PLACE_CHANNELIZING_MARKERS", verticesTSV=verts_tsv,
                     halfSizeFt=half_size_ft, reason=reason),
        "place_channelizing_markers")


def place_cell(cell_name: str, pt_x: float, pt_y: float, pt_z: float = 0, angle_deg: float = 0,
               library_path: str = "", reason: str = "") -> dict:
    """Place a cell at (pt_x, pt_y). Default library is WZTC (ny_plan_wztc.cel).
    Pass library_path for other libs (e.g. ny_plan_striping.cel turn arrows)."""
    params = {
        "cellName": cell_name, "ptX": pt_x, "ptY": pt_y, "ptZ": pt_z,
        "angleDeg": angle_deg, "reason": reason,
    }
    lib = (library_path or "").strip()
    if lib:
        params["libraryPath"] = lib
    return _ok_or_raise(_bridge.call("PLACE_CELL", **params), "place_cell")


def place_cell_on_post(cell_name: str, pt_x: float, pt_y: float, dir_x: float, dir_y: float,
                        pt_z: float = 0, angle_deg: float = 0, reason: str = "") -> dict:
    """Place a cell on a 50 ft stem/post the same way a roadside sign is
    built (post-outward stem, cell's inward edge snapped to the stem's
    outer end) — for plan symbols like the Arrow Panel that should read
    like a sign instead of floating at a bare lateral offset (engineer
    ask 2026-08-10). (pt_x, pt_y) is the base/tick point the stem starts
    from; (dir_x, dir_y) is the outward unit direction the stem/cell sit
    along. Returns elementId (the cell) and stemElementId (the post
    line)."""
    resp = _ok_or_raise(
        _bridge.call("PLACE_CELL_ON_POST", cellName=cell_name, ptX=pt_x, ptY=pt_y, ptZ=pt_z,
                     dirX=dir_x, dirY=dir_y, angleDeg=angle_deg, reason=reason),
        "place_cell_on_post")
    resp["createdElementIds"] = [str(v) for v in
                                  (resp.get("elementId"), resp.get("stemElementId")) if v]
    return resp


def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Finish symbology after place_sign. Labels/text → SF_P white wt=3;
    stems → SF_P white; post cell TWZSGN_P → color 6 (orange). Face cells
    are intentionally LEFT ALONE (library SF_P/SFB_P + ByCell weights) —
    do NOT also call change_element_symbology on faces (forcing color 0/6
    or weight 3 bleaches or wrecks the legend; live 2026-08-03).
    element_ids from place_sign createdElementIds; applied count may be
    less than requested because faces are skipped on purpose."""
    resp = _ok_or_raise(
        _bridge.call("SET_SIGN_ATTRIBUTES", elementIds=",".join(element_ids), reason=reason),
        "set_sign_attributes")
    if _PLAN_SESSION.sheet_plan_active():
        _PLAN_SESSION.sign_attrs_applied = True
        _save_sheet_plan()
    return _attach_plan_next(resp) if isinstance(resp, dict) else resp


def handoff(kind: str, from_sta: Optional[float] = None, to_sta: Optional[float] = None,
            text: Optional[str] = None, notes: Optional[str] = None, reason: str = "") -> dict:
    """Queue a dimension or callout for MANUAL placement. kind is
    'dimension' or 'callout'. Unlike every place_* tool above, these have
    NO programmatic CadInputQueue precedent anywhere in this codebase (the
    "red list" — see the plan) and are never faked as complete. Returns
    status DEFERRED, not OK. After drawing everything else, summarize what's
    queued (list_deferred_handoffs) and tell the engineer the existing
    PlaceElements/PlaceCells form handles exactly those items with a few
    clicks."""
    params = {"kind": kind, "reason": reason}
    if from_sta is not None:
        params["fromSta"] = from_sta
    if to_sta is not None:
        params["toSta"] = to_sta
    if text is not None:
        params["text"] = text
    if notes is not None:
        params["notes"] = notes
    return _ok_or_raise(_bridge.call("HANDOFF", **params), "handoff")


# ======================================== Spec placement-plan compiler
# sheet_spec.compile_* produces absolute-coordinate primitives; these
# tools fetch alignment vertices, compile, gate, and optionally place.
# Prefer place_sheet_geometry over place_order_table_labels/dimensions/
# channelizing/workspace/symbol_cells for sheets that have a JSON spec.

def _shoulder_width_ft(shoulder_width: str) -> float:
    """Collapse a display band / dropdown value to a numeric ft for geometry.
    Prefer an explicit '12 ft' / '8 ft' number; band labels map to a
    representative width (same bands sheet_spec.shoulder_band uses)."""
    import re
    s = (shoulder_width or "").strip().lower().replace("–", "-").replace("≥", ">=")
    m = re.search(r"(\d+(?:\.\d+)?)\s*ft", s)
    if m:
        return float(m.group(1))
    if ">= 8" in s or ">=8" in s:
        return 8.0
    if "5" in s and "7" in s:
        return 6.0
    if "< 5" in s or "4 ft" in s or s.startswith("4"):
        return 4.0
    return 8.0


def _segments_for_align(align_idx: int):
    import alignment_geometry as ag
    rows = get_alignment_vertices(align_idx)
    if not rows:
        raise ValueError(
            f"no vertices for align_idx={align_idx} — commit_alignment or "
            f"adopt_alignment first")
    return ag.parse_vertices(rows)


def _filter_symbol_alts(symbol_prims: list[dict], arrow_panel_choice: str) -> list[dict]:
    """arrow_panel_choice: 'trailer' keeps arrowPanel, drops alt-group PVs;
    'vehicle' keeps the PV partner, drops arrowPanel."""
    choice = (arrow_panel_choice or "trailer").strip().lower()
    by_group: dict[str, list[dict]] = {}
    for p in symbol_prims:
        g = p.get("altGroup")
        if g:
            by_group.setdefault(g, []).append(p)
    drop_ids = set()
    for g, members in by_group.items():
        ap = [m for m in members if m.get("kind") == "arrowPanel"]
        pvs = [m for m in members if m.get("kind") == "protectiveVehicle"]
        if choice in ("trailer", "arrow", "arrowpanel", "ap"):
            for m in pvs:
                drop_ids.add(m["id"])
        else:
            for m in ap:
                drop_ids.add(m["id"])
    out = []
    for p in symbol_prims:
        if p.get("kind") in ("protectiveVehicle", "arrowPanel") and p.get("id") in drop_ids:
            continue
        if p.get("kind") == "vehicleMountedSign" and p.get("mountedOn") in drop_ids:
            continue
        out.append(p)
    return out


def compile_sheet_plan(sheet_num: str, speed: int, lane_width: int, shoulder_width: str,
                        area_type: str = "", closure_type: str = "",
                        exposure_condition: str = "",
                        protective_vehicle_gvw: int = 0,
                        align_idxs: Optional[list[int]] = None,
                        outward_sign: float = -1.0,
                        sheet_elements: str = "",
                        arrow_panel_choice: str = "trailer",
                        include_primitives: bool = False,
                        force: bool = False) -> dict:
    """Compile a sheet-faithful placement plan in absolute model coords
    (no drawing). Requires Data/sheet-specs/<sheet>.json and committed/
    adopted alignments for each align_idxs entry.

    Blank designer kwargs are filled from get_locked_designer_inputs when
    an order table was built this session; conflicts raise unless force=True.

    Returns gateFailures (empty = pass), primitive counts, and (when
    include_primitives=True) the plan dict that place_sheet_geometry
    executes. Prefer place_sheet_geometry(dry_run=True) for a one-shot
    preview — leave include_primitives=False for agent calls (coords are
    huge); place_sheet_geometry keeps the full plan in-process."""
    import sheet_spec

    merged = _merge_locked_designer_inputs(
        sheet_num, speed, lane_width, shoulder_width,
        area_type=area_type, closure_type=closure_type,
        exposure_condition=exposure_condition,
        protective_vehicle_gvw=protective_vehicle_gvw, force=force)
    sheet_num = merged["sheet_num"]
    speed = int(merged["speed"])
    lane_width = int(merged["lane_width"])
    shoulder_width = str(merged["shoulder_width"])
    area_type = merged["area_type"] or ""
    closure_type = merged["closure_type"] or ""
    exposure_condition = merged["exposure_condition"] or ""
    protective_vehicle_gvw = int(merged["protective_vehicle_gvw"] or 0)
    outward_sign, _, _ = _apply_locked_lateral(outward_sign, 40.0, True)

    spec = sheet_spec.load(sheet_num)
    if spec is None:
        raise ValueError(
            f"no sheet spec for {sheet_num!r} — cannot compile. "
            f"Use place_order_table_* heuristics only as a last resort, or "
            f"author Data/sheet-specs/{sheet_num}.json first.")

    roles = spec.get("tableRoles") or {}
    if roles.get("advanceWarningSpacing") and not area_type:
        plan_workflow.raise_plan_gate(
            f"sheet {sheet_num} needs area_type.",
            tool="compile_sheet_plan",
            current_step="compiler_placed",
            missing=["area_type"],
            accepted=["URBAN", "RURAL", "FREEWAY"],
            next_tool="place_sheet_geometry",
            next_step=(
                "Pass area_type from get_locked_designer_inputs (or omit and "
                "let auto-fill). Do not guess other parameters."
            ),
        )

    gvw = protective_vehicle_gvw if protective_vehicle_gvw and protective_vehicle_gvw > 0 else None
    try:
        resolved = sheet_spec.resolve(
            spec, speed, lane_width, shoulder_width,
            area_type or None, closure_type or None, exposure_condition or None,
            protective_vehicle_gvw=gvw)
    except sheet_spec.SpecError as e:
        raise ValueError(str(e)) from e

    if not sheet_elements:
        req = get_sheet_requirements(sheet_num)
        sheet_elements = req.get("elements") or ""

    if align_idxs:
        idxs = list(align_idxs)
    else:
        # Default to every alignment the sheet itself declares (not just
        # [1]) so two-alignment sheets (e.g. 619-311's upstream+downstream
        # split) get their work-area hatch compiled by default -- compile_hatch
        # needs both align 1 and align 2 present in segs_by_align, and
        # silently defaulting to [1] alone used to drop it with no error.
        idxs = sorted({a["alignIdx"] for a in spec.get("orderTable", {}).get("alignments", [])}) or [1]
    sh_ft = _shoulder_width_ft(shoulder_width)
    segs_by_align = {i: _segments_for_align(i) for i in idxs}

    plan_by_align: dict[str, list] = {}
    chan_by_align: dict[str, list] = {}
    sym_by_align: dict[str, list] = {}
    all_plan, all_chan, all_sym = [], [], []

    tip_hl = (
        float(_PLAN_SESSION.lateral_half_len)
        if _PLAN_SESSION.lateral_half_len is not None else None)
    for a in idxs:
        segs = segs_by_align[a]
        plan = sheet_spec.compile_plan(
            spec, resolved, a, segs, outward_sign=outward_sign,
            sheet_elements=sheet_elements, tip_half_len_ft=tip_hl)
        chan = sheet_spec.compile_channelizing(
            spec, resolved, a, segs, lane_width_ft=float(lane_width),
            shoulder_width_ft=sh_ft, outward_sign=outward_sign)
        sym = _filter_symbol_alts(
            sheet_spec.compile_symbols(
                spec, resolved, a, segs, outward_sign=outward_sign,
                lane_width_ft=float(lane_width), shoulder_width_ft=sh_ft,
                tip_half_len_ft=tip_hl),
            arrow_panel_choice)
        plan_by_align[str(a)] = plan
        chan_by_align[str(a)] = chan
        sym_by_align[str(a)] = sym
        all_plan.extend(plan)
        all_chan.extend(chan)
        all_sym.extend(sym)

    hatch: list = []
    if 1 in segs_by_align and 2 in segs_by_align:
        hatch = sheet_spec.compile_hatch(
            spec, resolved, segs_by_align[1], segs_by_align[2],
            lane_width_ft=float(lane_width), shoulder_width_ft=sh_ft,
            outward_sign=outward_sign,
            work_bay_vertices=_PLAN_SESSION.work_bay_vertices)
    elif 1 in idxs and 2 not in idxs:
        hatch = [{"kind": "note",
                  "text": "hatch skipped — need both align 1 and 2 committed/adopted"}]

    gate_align = idxs[0]
    gate = sheet_spec.run_rules_gate(
        spec, resolved, gate_align,
        plan_by_align.get(str(gate_align), []),
        chan_by_align.get(str(gate_align), []),
        sym_by_align.get(str(gate_align), []),
        hatch if isinstance(hatch, list) else None)

    def _count(prims):
        from collections import Counter
        return dict(Counter(p.get("kind", "?") for p in prims))

    out = {
        "status": "OK" if not gate else "GATE_FAILED",
        "sheet": sheet_num,
        "specDriven": True,
        "gateFailures": gate,
        "counts": {
            "plan": _count(all_plan),
            "channelizing": _count(all_chan),
            "symbols": _count(all_sym),
            "hatch": _count([h for h in hatch if isinstance(h, dict)]) if hatch else {},
        },
        "arrowPanelChoice": arrow_panel_choice,
        "planAlignIdxs": idxs,
    }
    if merged.get("filledFromLock"):
        out["filledFromLock"] = merged["filledFromLock"]
    if include_primitives:
        out["plan"] = {
            "alignIdxs": idxs,
            "planByAlign": plan_by_align,
            "channelizingByAlign": chan_by_align,
            "symbolsByAlign": sym_by_align,
            "hatch": hatch,
            "sheetElements": sheet_elements,
        }
    return out


def execute_compiled_plan(plan: dict, layers: Optional[list[str]] = None,
                           force: bool = False,
                           sheet_num: str = "") -> dict:
    """Place primitives from compile_sheet_plan / place_sheet_geometry.
    Internal helper — agent should call place_sheet_geometry, not this.
    layers defaults to dimensions,labels,channelizing,symbols,hatch
    (stations/signs stay on place_order_table_stations + place_sign).
    Refuses if plan['gateFailures'] is non-empty unless force=True.

    Captures createdElementIds into Bridge/placement-registry.jsonl."""
    if not plan or "plan" not in plan:
        raise ValueError("execute_compiled_plan needs the dict returned by compile_sheet_plan")

    inner = plan["plan"]
    if "planByAlign" not in inner:
        raise ValueError("compiled plan missing planByAlign — re-run compile_sheet_plan")

    gate = plan.get("gateFailures") or []
    if gate and not force:
        raise ValueError(
            f"rules gate failed ({len(gate)}): {gate[:3]}… Pass force=True to place anyway, "
            f"or fix inputs / alignments and recompile.")

    sheet = sheet_num or plan.get("sheet") or ""
    if not sheet and _PLAN_SESSION.designer_inputs:
        sheet = _PLAN_SESSION.designer_inputs.sheet_num

    want = set(layers or ["dimensions", "labels", "channelizing", "symbols", "hatch"])
    placed: list[dict] = []
    errors: list[str] = []

    def _register(resp, *, kind, primitive_id, bridge_op, align_idx, spec_ref, layer, detail,
                  geom_extra=None):
        ids = placement_registry.parse_created_ids(resp if isinstance(resp, dict) else {})
        req_id = ""
        if isinstance(resp, dict):
            req_id = str(resp.get("reqId") or resp.get("req_id") or "")
        entry = {"layer": layer, **detail, "createdElementIds": ids,
                 "primitiveId": primitive_id, "reqId": req_id}
        if isinstance(resp, dict):
            entry["status"] = resp.get("status", detail.get("status"))
        placed.append(entry)
        if ids:
            extra = dict(geom_extra or {})
            placement_registry.append_placement(
                sheet_num=sheet,
                align_idx=int(align_idx or 0),
                kind=kind,
                primitive_id=primitive_id or f"0:unknown:{kind}",
                bridge_op=bridge_op,
                element_ids=ids,
                spec_ref=spec_ref or {},
                req_id=req_id,
                extra=extra or None,
            )
        return entry

    # --- dimensions + labels (per align); skip station primitives (ticks
    # come from place_order_table_stations)
    if "dimensions" in want or "labels" in want:
        for a_str, prims in inner["planByAlign"].items():
            align_i = int(a_str) if str(a_str).isdigit() else 0
            for p in prims:
                try:
                    if p["kind"] == "dimension" and "dimensions" in want:
                        t1, t2, off = p["tip1"], p["tip2"], p["offset"]
                        if p.get("curved") and p.get("path") and len(p["path"]) >= 2:
                            r = place_path_hugging_dimension(
                                p["path"], p.get("text") or "", off,
                                reason=f"compiled curve dim {p.get('text','')}",
                                force_arc=True)
                            bridge_op = {
                                "CurvedPlanArc": "PLACE_CURVED_PLAN_DIMENSION",
                                "ArcSize": "PLACE_ARC_SIZE_DIMENSION",
                            }.get(r.get("dimType") or "", "PLACE_DIMENSION")
                        else:
                            sheet_txt = _format_ny_plan_dim_text(p.get("text") or "")
                            r = place_dimension(
                                t1[0], t1[1], t2[0], t2[1], off[0], off[1],
                                reason=f"compiled dim {p.get('text','')}",
                                override_text=sheet_txt)
                            bridge_op = "PLACE_DIMENSION"
                        mid = (0.5 * (float(t1[0]) + float(t2[0])),
                               0.5 * (float(t1[1]) + float(t2[1])))
                        _register(
                            r, kind="dimension",
                            primitive_id=p.get("primitiveId") or "",
                            bridge_op=bridge_op,
                            align_idx=align_i,
                            spec_ref=p.get("specRef"),
                            layer="dimension",
                            detail={"text": p.get("text"),
                                    "curved": bool(p.get("curved")),
                                    "partKind": (p.get("specRef") or {}).get("partKind"),
                                    "partsSumFt": (p.get("specRef") or {}).get("partsSumFt"),
                                    "sheetLengthFt": (p.get("specRef") or {}).get("sheetLengthFt")},
                            geom_extra={
                                "tip1": list(t1)[:3], "tip2": list(t2)[:3],
                                "offset": list(off)[:3],
                                "midX": mid[0], "midY": mid[1],
                                "text": p.get("text"),
                                "curved": bool(p.get("curved")),
                                "partKind": (p.get("specRef") or {}).get("partKind"),
                            },
                        )
                    elif p["kind"] == "label" and "labels" in want:
                        r = place_text_label(p["text"], p["x"], p["y"],
                                             reason="compiled Non-Sign label",
                                             angle_deg=float(p.get("angleDeg") or 0.0))
                        _register(
                            r, kind="label",
                            primitive_id=p.get("primitiveId") or "",
                            bridge_op="PLACE_TEXT_LABEL",
                            align_idx=align_i,
                            spec_ref=p.get("specRef"),
                            layer="label",
                            detail={"text": p.get("text")},
                            geom_extra={
                                "x": float(p["x"]), "y": float(p["y"]),
                                "text": p.get("text"),
                            },
                        )
                except Exception as e:
                    errors.append(f"{p.get('kind')}: {e}")

    # --- channelizing: discrete markers (representation.mode=markers)
    if "channelizing" in want:
        for a_str, prims in inner.get("channelizingByAlign", {}).items():
            align_i = int(a_str) if str(a_str).isdigit() else 0
            by_run: dict[str, list] = {}
            for p in prims:
                if p.get("kind") == "cone":
                    by_run.setdefault(p.get("run", "run"), []).append(p)
            for run_id, cones in by_run.items():
                cones_sorted = sorted(cones, key=lambda c: float(c.get("stationFt", 0)))
                if not cones_sorted:
                    continue
                rep = cones_sorted[0].get("representation") or {
                    "mode": "markers", "markerHalfSizeFt": 1.5}
                mode = str(rep.get("mode") or "markers").lower()
                if mode != "markers":
                    errors.append(
                        f"channelizing {run_id}: unsupported representation.mode="
                        f"{mode!r} (only 'markers' is implemented)")
                    continue
                half = float(rep.get("markerHalfSizeFt") or 1.5)
                verts = [[c["x"], c["y"], 0.0] for c in cones_sorted]
                try:
                    r = place_channelizing_markers(
                        verts, half_size_ft=half,
                        reason=f"compiled channelizing markers {run_id}")
                    prim_id = (cones_sorted[0].get("primitiveId")
                               or f"{align_i}:{run_id}:cone")
                    spec_ref = cones_sorted[0].get("specRef") or {
                        "zone": None, "run": run_id, "alignIdx": align_i}
                    _register(
                        r, kind="cone",
                        primitive_id=prim_id,
                        bridge_op="PLACE_CHANNELIZING_MARKERS",
                        align_idx=align_i,
                        spec_ref=spec_ref,
                        layer="channelizing_markers",
                        detail={"run": run_id, "cones": len(cones_sorted)},
                        geom_extra={
                            "x": float(cones_sorted[0]["x"]),
                            "y": float(cones_sorted[0]["y"]),
                            "run": run_id,
                            "coneCount": len(cones_sorted),
                            "stationFt": cones_sorted[0].get("stationFt"),
                        },
                    )
                except Exception as e:
                    errors.append(f"channelizing {run_id}: {e}")

    # --- symbols
    if "symbols" in want:
        placed_arrow_panel = False
        for a_str, prims in inner.get("symbolsByAlign", {}).items():
            align_i = int(a_str) if str(a_str).isdigit() else 0
            for p in prims:
                try:
                    if p["kind"] == "arrowPanel":
                        # One AP only (count=1). Skip Align2 / re-entry dupes
                        # and leftover stacked rebuilds (live 2026-08-10).
                        if placed_arrow_panel:
                            continue
                        placed_arrow_panel = True
                        r = place_cell_on_post(
                            p["cellName"], p["x"], p["y"], p["dirX"], p["dirY"],
                            angle_deg=p.get("angleDeg", 0.0),
                            reason=f"compiled {p['kind']} {p.get('id','')}")
                        _register(
                            r, kind=p["kind"],
                            primitive_id=p.get("primitiveId") or "",
                            bridge_op="PLACE_CELL_ON_POST",
                            align_idx=align_i,
                            spec_ref=p.get("specRef"),
                            layer=p["kind"],
                            detail={"id": p.get("id"),
                                    "requiredNote": p.get("requiredNote")},
                            geom_extra={
                                "x": float(p["x"]), "y": float(p["y"]),
                                "stationFt": p.get("stationFt"),
                                "id": p.get("id"),
                                "altGroup": p.get("altGroup"),
                            },
                        )
                    elif p["kind"] == "protectiveVehicle":
                        r = place_cell(p["cellName"], p["x"], p["y"], 0.0,
                                       p.get("angleDeg", 0.0),
                                       reason=f"compiled {p['kind']} {p.get('id','')}")
                        _register(
                            r, kind=p["kind"],
                            primitive_id=p.get("primitiveId") or "",
                            bridge_op="PLACE_CELL",
                            align_idx=align_i,
                            spec_ref=p.get("specRef"),
                            layer=p["kind"],
                            detail={"id": p.get("id"),
                                    "requiredNote": p.get("requiredNote")},
                            geom_extra={
                                "x": float(p["x"]), "y": float(p["y"]),
                                "stationFt": p.get("stationFt"),
                                "id": p.get("id"),
                                "altGroup": p.get("altGroup"),
                            },
                        )
                    elif p["kind"] == "label":
                        r = place_text_label(p["text"], p["x"], p["y"],
                                             reason="compiled symbol label",
                                             angle_deg=float(p.get("angleDeg") or 0.0))
                        _register(
                            r, kind="label",
                            primitive_id=p.get("primitiveId") or "",
                            bridge_op="PLACE_TEXT_LABEL",
                            align_idx=align_i,
                            spec_ref=p.get("specRef"),
                            layer="symbol_label",
                            detail={"text": p.get("text")},
                            geom_extra={
                                "x": float(p["x"]), "y": float(p["y"]),
                                "text": p.get("text"),
                            },
                        )
                    elif p["kind"] == "vehicleMountedSign":
                        r = handoff(
                            kind="callout",
                            notes=(f"{p.get('signCode')} vehicle-mounted on "
                                   f"{p.get('mountedOn')} at ({p['x']:.1f},{p['y']:.1f})"),
                            reason="compiled vehicle-mounted sign (not a roadside post)")
                        placed.append({
                            "layer": "vehicleMountedSign",
                            "signCode": p.get("signCode"),
                            "status": r.get("status") if isinstance(r, dict) else None,
                            "primitiveId": p.get("primitiveId"),
                            "createdElementIds": [],
                        })
                except Exception as e:
                    errors.append(f"symbol {p.get('kind')}: {e}")

    # --- hatch + transverse (place_workspace must NOT repeat first vertex)
    if "hatch" in want:
        for p in inner.get("hatch") or []:
            try:
                if p.get("kind") == "hatch":
                    verts = [[x, y, 0.0] for x, y in p["boundary"]]
                    r = place_workspace(verts, reason="compiled work-area hatch")
                    _register(
                        r, kind="hatch",
                        primitive_id=p.get("primitiveId") or "0:workArea:hatch",
                        bridge_op="PLACE_WORKSPACE",
                        align_idx=0,
                        spec_ref=p.get("specRef"),
                        layer="hatch",
                        detail={"workAreaLengthFt": p.get("workAreaLengthFt")},
                    )
                elif p.get("kind") == "transverseRun":
                    t1, t2 = p["tip1"], p["tip2"]
                    r = place_element_run(2, [[t1[0], t1[1], 0], [t2[0], t2[1], 0]],
                                          reason="compiled transverse channelizing")
                    _register(
                        r, kind="transverseRun",
                        primitive_id=p.get("primitiveId") or "",
                        bridge_op="PLACE_ELEMENT_RUN",
                        align_idx=0,
                        spec_ref=p.get("specRef"),
                        layer="transverseRun",
                        detail={},
                    )
            except Exception as e:
                errors.append(f"hatch: {e}")

    return {
        "status": "OK" if not errors else "PARTIAL",
        "placedCount": len(placed),
        "placed": placed[:40],
        "placedTruncated": max(0, len(placed) - 40),
        "placedWithIds": [
            {"primitiveId": p.get("primitiveId"), "kind": p.get("layer"),
             "elementIds": p.get("createdElementIds") or []}
            for p in placed if p.get("createdElementIds")
        ][:80],
        "errors": errors,
    }


def place_sheet_geometry(sheet_num: str, speed: int, lane_width: int, shoulder_width: str,
                          area_type: str = "", closure_type: str = "",
                          exposure_condition: str = "",
                          protective_vehicle_gvw: int = 0,
                          align_idxs: Optional[list[int]] = None,
                          outward_sign: float = -1.0,
                          sheet_elements: str = "",
                          arrow_panel_choice: str = "trailer",
                          dry_run: bool = False,
                          force: bool = False,
                          layers: Optional[list[str]] = None) -> dict:
    """Compile + (unless dry_run) place sheet-faithful dims/labels/
    channelizing/symbols/hatch from Data/sheet-specs/<sheet>.json.

    Prefer this over place_order_table_labels / place_order_table_dimensions /
    place_order_table_channelizing / place_order_table_workspace /
    place_sheet_symbol_cells when a sheet spec exists — those batch tools
    use generic heuristics; this path uses the placement-plan compiler.

    Still call separately: build_wztc_order_table, place_order_table_stations,
    place_sign for every isSign row. Alignments must already be
    committed or adopted.

    Blank designer kwargs (esp. area_type='') are auto-filled from the
    locked build_wztc_order_table inputs — do NOT re-ask the engineer.
    Conflicting values raise unless force=True.

    dry_run=True: compile + rules gate only (no drawing).
    force=True: place even if the rules gate reports failures; also allows
    intentional designer-input overrides.
    arrow_panel_choice: 'trailer' (default TWZAP_P) or 'vehicle' (OR PV)."""
    # Sheet-plan checklist: stations (and normally signs) before compiler.
    # force=True skips; general CAD never hits this (no order table).
    if _PLAN_SESSION.sheet_plan_active() and not force and not dry_run:
        done = plan_workflow.stage_done(_PLAN_SESSION)
        if not done["corridor_ready"]:
            plan_workflow.raise_plan_gate(
                "corridor not ready for place_sheet_geometry.",
                tool="place_sheet_geometry",
                current_step="corridor_ready",
                next_tool="assemble_corridor",
                next_step="assemble_corridor then place_order_table_stations",
            )
        if not done["stations_placed"]:
            st = plan_workflow.next_action(_PLAN_SESSION, done)
            plan_workflow.raise_plan_gate(
                "stations incomplete for place_sheet_geometry.",
                tool="place_sheet_geometry",
                current_step="stations_placed",
                missing=st.get("stationsNeeded") or [],
                next_tool="place_order_table_stations",
                next_step=st.get("nextStep") or "",
            )
        if not done["signs_placed"]:
            st = plan_workflow.next_action(_PLAN_SESSION, done)
            plan_workflow.raise_plan_gate(
                "order-table signs incomplete for place_sheet_geometry.",
                tool="place_sheet_geometry",
                current_step="signs_placed",
                missing=st.get("remainingSigns") or [],
                next_tool="place_sign",
                next_step=st.get("nextStep") or "",
            )
        if not done["sign_attrs_applied"]:
            plan_workflow.raise_plan_gate(
                "set_sign_attributes not applied yet for this sheet build.",
                tool="place_sheet_geometry",
                current_step="sign_attrs_applied",
                next_tool="set_sign_attributes",
                next_step="set_sign_attributes on createdElementIds from each place_sign",
            )
    # Same station/path preflight as place_order_table_stations — compiler
    # hatch/dims assume Align1/2 sta0 are work-area edges and path length
    # covers the walk.
    xv = None
    if _PLAN_SESSION.designer_inputs is not None and not dry_run:
        idxs = align_idxs if align_idxs else [1, 2]
        fails_acc: list[str] = []
        aligns_out: list[dict] = []
        for aidx in idxs:
            one = cross_validate_stations(align_idx=int(aidx), force=True)
            aligns_out.extend(one.get("alignments") or [])
            fails_acc.extend(one.get("failures") or [])
        xv = {"status": "OK" if not fails_acc else "FAIL",
              "alignments": aligns_out, "failures": fails_acc}
        if fails_acc and not force:
            raise ValueError(
                "place_sheet_geometry blocked by cross_validate_stations: "
                + "; ".join(fails_acc)
                + " Fix via assemble_corridor / redefine, or force=True."
            )
    compiled = compile_sheet_plan(
        sheet_num, speed, lane_width, shoulder_width,
        area_type=area_type, closure_type=closure_type,
        exposure_condition=exposure_condition,
        protective_vehicle_gvw=protective_vehicle_gvw,
        align_idxs=align_idxs, outward_sign=outward_sign,
        sheet_elements=sheet_elements,
        arrow_panel_choice=arrow_panel_choice,
        include_primitives=True, force=force)
    if dry_run:
        slim = {k: v for k, v in compiled.items() if k != "plan"}
        slim["note"] = "dry_run — nothing drawn; pass dry_run=False to place"
        return slim
    executed = execute_compiled_plan(
        compiled, layers=layers, force=force,
        sheet_num=str(compiled.get("sheet") or sheet_num))
    _PLAN_SESSION.sheet_geometry_placed = True
    _PLAN_SESSION.find_near_calls = 0  # allow ≤2 targeted QA lookups, then refuse
    gates = compiled.get("gateFailures") or []
    sheet_key = str(compiled.get("sheet") or sheet_num)
    reg_rows = placement_registry.resolve_latest_placements(sheet_num=sheet_key)
    model_rows = []
    try:
        path = list(_PLAN_SESSION.corridor_path or _PLAN_SESSION.work_bay_vertices or [])
        model_rows = _model_rows_for_path(path)
    except Exception:
        model_rows = []
    scorecard = sheet_scorecard.build_placement_scorecard(
        compiled, registry_rows=reg_rows, executed=executed, gate_failures=gates,
        model_rows=model_rows)
    _PLAN_SESSION.last_scorecard = scorecard
    _PLAN_SESSION.last_compiled = compiled
    _PLAN_SESSION.geometry_qa_passed = bool(scorecard.get("passed"))
    _PLAN_SESSION.visual_qa_passed = False
    _PLAN_SESSION.visual_qa_failures = []
    if not scorecard.get("passed"):
        _PLAN_SESSION.last_failed_phase = "place_sheet_geometry"
    _save_sheet_plan()
    out = {
        "status": executed["status"] if scorecard.get("passed") else "PARTIAL",
        "sheet": sheet_key,
        "gateFailures": gates,
        "counts": compiled.get("counts"),
        "placedCount": executed.get("placedCount"),
        "placed": executed.get("placed"),
        "placedWithIds": executed.get("placedWithIds"),
        "placedTruncated": executed.get("placedTruncated"),
        "errors": executed.get("errors"),
        "scorecard": {
            "passed": scorecard.get("passed"),
            "failures": (scorecard.get("failures") or [])[:20],
            "expectedByKind": (scorecard.get("expected") or {}).get("byKind"),
            "placedByKind": (scorecard.get("placed") or {}).get("byKind"),
            "missingPrimitiveIds": scorecard.get("missingPrimitiveIds") or [],
            "citationCount": len(scorecard.get("citations") or []),
        },
        "arrowPanelChoice": arrow_panel_choice,
        "deferred": [
            p for p in (executed.get("placed") or [])
            if str(p.get("status", "")).upper() in ("DEFERRED", "HANDOFF")
            or p.get("layer") == "vehicleMountedSign"
        ],
        "nextStep": (
            "run_visual_qa_captures" if scorecard.get("passed")
            else "Fix scorecard.failures then re-run place_sheet_geometry "
                 "(or force=True only if engineer accepts)"
        ),
    }
    if compiled.get("filledFromLock"):
        out["filledFromLock"] = compiled["filledFromLock"]
    if xv is not None:
        out["crossValidate"] = xv
    return _attach_plan_next(out)


def _sign_detail_for(align_idx: int, sign_num: str) -> dict:
    key = str(sign_num).strip().upper()
    for r in _PLAN_SESSION.locked_sign_details:
        if int(r.get("align_idx") or 0) == int(align_idx) and str(r.get("sign_num", "")).strip().upper() == key:
            return r
    return {"align_idx": align_idx, "sign_num": sign_num, "side": "One Side"}


def _place_locked_signs_from_stations(outward_sign: float = -1.0,
                                      half_len: float = 40.0) -> dict:
    """Place every locked order-table sign at the outward perp tip of its
    station row. Uses last_station_rows from place_order_table_stations."""
    from sheet_compile import _outward_unit, _post_angle_deg

    inputs = _PLAN_SESSION.designer_inputs
    if inputs is None:
        raise ValueError("no locked designer inputs")
    road_type = inputs.road_type or "Non-Freeway"
    placed: list[dict] = []
    attr_ids: list[str] = []
    errors: list[str] = []

    for align_idx in sorted(_PLAN_SESSION.required_aligns or {1, 2}):
        rows = _PLAN_SESSION.last_station_rows.get(int(align_idx)) or []
        if not rows:
            errors.append(f"align {align_idx}: no station rows cached — place_order_table_stations first")
            continue
        for row in rows:
            if str(row.get("isSign", "")).strip().upper() not in ("Y", "YES", "TRUE", "1"):
                continue
            sign_num = str(row.get("label") or "").strip()
            if not sign_num:
                errors.append(f"align {align_idx}: isSign row missing label")
                continue
            if (int(align_idx), sign_num.upper()) in _PLAN_SESSION.signs_placed_rows:
                continue
            try:
                pt_x = float(row["ptX"])
                pt_y = float(row["ptY"])
                pt_z = float(row.get("ptZ") or 0)
                tan_x = float(row["tanX"])
                tan_y = float(row["tanY"])
            except (KeyError, TypeError, ValueError) as e:
                errors.append(f"align {align_idx} {sign_num}: bad station coords ({e})")
                continue
            detail = _sign_detail_for(align_idx, sign_num)
            side = str(detail.get("side") or "One Side")
            # Align2 tan points away downstream (= +travel); Align1 tan points
            # away upstream (= −travel). _outward_unit(Align2_tan, outward_sign)
            # flips across the road — use Align1-equivalent basis (−Align2 tan)
            # so one-side tips stay on the closed shoulder locally. Do NOT use
            # a single world-locked closed_outward: on curved corridors that
            # freezes the WA-mid normal and disconnects assemblies around bends
            # (engineer QA 2026-08-13).
            if int(align_idx) == 2:
                basis_tx, basis_ty = -tan_x, -tan_y
            else:
                basis_tx, basis_ty = tan_x, tan_y
            out_x, out_y = _outward_unit(basis_tx, basis_ty, outward_sign)
            tip_x = pt_x + out_x * half_len
            tip_y = pt_y + out_y * half_len
            kwargs = dict(
                sign_num=sign_num, road_type=road_type, side=side,
                pt1x=tip_x, pt1y=tip_y, pt1z=pt_z, dir1x=out_x, dir1y=out_y,
                align_idx=int(align_idx), one_off=False,
                post_angle_deg=_post_angle_deg(int(align_idx), tan_x, tan_y),
            )
            if side.strip().lower() == "both sides":
                kwargs.update(
                    pt2x=pt_x - out_x * half_len,
                    pt2y=pt_y - out_y * half_len,
                    pt2z=pt_z,
                    dir2x=-out_x,
                    dir2y=-out_y,
                )
            try:
                resp = place_sign(**kwargs)
                ids = []
                if isinstance(resp, dict):
                    raw = resp.get("createdElementIds") or resp.get("elementIds") or ""
                    if isinstance(raw, str) and raw.strip():
                        ids = [x.strip() for x in raw.replace(";", ",").split(",") if x.strip()]
                    elif isinstance(raw, list):
                        ids = [str(x) for x in raw]
                    eid = resp.get("elementId")
                    if eid:
                        ids.append(str(eid))
                attr_ids.extend(ids)
                placed.append({"align_idx": align_idx, "sign_num": sign_num,
                               "status": (resp or {}).get("status", "OK"),
                               "elementIds": ids})
            except Exception as e:
                errors.append(f"align {align_idx} {sign_num}: {e}")

    attrs = None
    if attr_ids:
        try:
            attrs = set_sign_attributes(attr_ids, reason="run_sheet_build batch")
        except Exception as e:
            errors.append(f"set_sign_attributes: {e}")
    elif placed and not errors:
        # Signs placed but no IDs returned — mark attrs incomplete
        pass

    return {"placed": placed, "setSignAttributes": attrs, "errors": errors,
            "attrElementIds": attr_ids}


def run_sheet_build(upstream_edge: Optional[list[float]] = None,
                    downstream_edge: Optional[list[float]] = None,
                    outward_sign: float = -1.0,
                    half_len: float = 40.0,
                    arrow_panel_choice: str = "trailer",
                    include_visual_qa: bool = True,
                    clear_prior_stations: bool = False,
                    force: bool = False,
                    approach_length_ft: float = 0.0,
                    use_locked_lateral: bool = True,
                    path_vertices: Optional[list] = None) -> dict:
    """SHEET-PLAN ONLY executor: advance the locked checklist without the
    LLM choosing step order.

    After build_wztc_order_table, the agent only needs to collect the two
    WORK AREA edge points (ask_user_choice point-pick) and call this once:

      resolve_sheet_lateral(up, dn, closed_side=..., real_road_edge=...,
                            path_vertices=… optional)
      run_sheet_build(upstream_edge=[...], downstream_edge=[...],
                      path_vertices=… same polyline for curved roads)

    path_vertices: pass the closed-lane / first-travel outer polyline when
    the corridor is curved (same arg as assemble_corridor). Omit for
    straight chord corridors.

    It then runs (skipping stages already done):
      assemble_corridor → stations → signs+attrs → place_sheet_geometry
      → optional run_visual_qa_captures

    outward_sign / half_len: when resolve_sheet_lateral has locked
    PlanSession lateral_* and use_locked_lateral=True (default), those
    locked values win over the −1 / 40 defaults. Pass
    use_locked_lateral=False to force the kwargs.

    Outside a sheet plan: returns sheetPlanActive=False (general CAD
    stays freeform — do not use this for ad-hoc drawing).

    force=True: pass through to assemble/place_sheet_geometry gates.
    clear_prior_stations=True: wipe+re-place stations before signs."""
    if not _PLAN_SESSION.sheet_plan_active():
        return {
            "status": "OK",
            "sheetPlanActive": False,
            "note": (
                "No named sheet plan active. run_sheet_build is only for "
                "619 standard-sheet builds after build_wztc_order_table. "
                "For general CAD, call place_*/adjust_view yourself."
            ),
        }

    inputs = _PLAN_SESSION.designer_inputs
    assert inputs is not None
    highway_caution = _highway_caution_for_sheet(inputs.sheet_num)
    overlap = check_build_overlap(
        sheet_num=inputs.sheet_num,
        path_vertices=path_vertices,
        lateral_half_width=float(_PLAN_SESSION.lateral_half_len or half_len or 40.0),
        scan_model=True,
    )
    outward_sign, half_len, lat_meta = _apply_locked_lateral(
        outward_sign, half_len, use_locked_lateral)
    phases: list[dict] = []
    phases.append({"phase": "highway_kind", "result": highway_caution})
    phases.append({"phase": "overlap", "result": (overlap or {}).get("overlapCaution")})
    if lat_meta.get("usedLockedLateral"):
        phases.append({"phase": "locked_lateral", "result": lat_meta})

    # Stale plan: order_table_built but lockedSignRows wiped — rebuild table
    # so PLACE_SIGN is not skipped (live agent 2026-08-10).
    if not _PLAN_SESSION.locked_sign_rows:
        ot = build_wztc_order_table(
            speed=inputs.speed,
            road_type=inputs.road_type,
            lane_width=inputs.lane_width,
            shoulder_width=inputs.shoulder_width,
            sheet_num=inputs.sheet_num,
            area_type=inputs.area_type or "",
            closure_type=inputs.closure_type or "",
            exposure_condition=inputs.exposure_condition or "",
            protective_vehicle_gvw=inputs.protective_vehicle_gvw or 0,
        )
        phases.append({"phase": "rebuild_order_table", "result": {
            "status": ot.get("status"),
            "lockedSignCount": len(_PLAN_SESSION.locked_sign_rows),
            "note": "Auto-rebuilt order table — lockedSignRows was empty",
        }})
        if not _PLAN_SESSION.locked_sign_rows:
            # Sheet truly has no roadside signs, or build failed to lock.
            pass

    done = plan_workflow.stage_done(_PLAN_SESSION)

    # --- corridor ---
    if not done["corridor_ready"]:
        if upstream_edge is None or downstream_edge is None:
            plan_workflow.raise_plan_gate(
                "corridor not ready — need work-area edge point-picks.",
                tool="run_sheet_build",
                current_step="corridor_ready",
                missing=["upstream_edge", "downstream_edge"],
                next_tool="ask_user_choice",
                next_step=(
                    "ask_user_choice(allow_point_pick=True) for upstream then "
                    "downstream WORK AREA edges, then re-call "
                    "run_sheet_build(upstream_edge=..., downstream_edge=..., "
                    "path_vertices=… if curved)"
                ),
            )
        corr = assemble_corridor(
            upstream_edge, downstream_edge,
            approach_length_ft=approach_length_ft, force=force,
            path_vertices=path_vertices)
        phases.append({"phase": "assemble_corridor", "result": {
            k: corr.get(k) for k in (
                "status", "workAreaLengthFt", "approachLengthFt",
                "stationWalkMaxFt", "curved", "nextStep") if k in corr
        }})
        done = plan_workflow.stage_done(_PLAN_SESSION)
    else:
        phases.append({"phase": "assemble_corridor", "skipped": True,
                       "note": "corridor already ready"})

    # --- stations ---
    req = sorted(_PLAN_SESSION.required_aligns or {1, 2})
    station_results = []
    first = True
    for aidx in req:
        need = (aidx not in _PLAN_SESSION.stations_placed_aligns) or clear_prior_stations
        if not need:
            station_results.append({"align_idx": aidx, "skipped": True})
            continue
        st = place_order_table_stations(
            align_idx=aidx,
            reset_session=first,
            clear_prior=clear_prior_stations and (aidx in _PLAN_SESSION.stations_placed_aligns),
            force=force or clear_prior_stations,
        )
        first = False
        station_results.append({
            "align_idx": aidx,
            "status": st.get("status"),
            "rowCount": len(st.get("rows") or []),
        })
    phases.append({"phase": "stations", "aligns": station_results})
    done = plan_workflow.stage_done(_PLAN_SESSION)

    # --- signs + attrs ---
    if not done["signs_placed"] or not done["sign_attrs_applied"]:
        signs = _place_locked_signs_from_stations(outward_sign=outward_sign, half_len=half_len)
        phases.append({"phase": "signs", "result": {
            "placedCount": len(signs.get("placed") or []),
            "errors": signs.get("errors") or [],
            "attrsStatus": (signs.get("setSignAttributes") or {}).get("status"),
        }})
        if signs.get("errors") and not force:
            plan_workflow.raise_plan_gate(
                "run_sheet_build sign phase had errors",
                tool="run_sheet_build",
                current_step="signs_placed",
                missing=signs["errors"][:8],
                next_tool="place_sign",
                next_step="Fix listed sign errors or pass force=True to continue",
            )
    else:
        phases.append({"phase": "signs", "skipped": True})
    done = plan_workflow.stage_done(_PLAN_SESSION)

    # --- compiler (re-enter if scorecard failed; preserve earlier phases) ---
    geom = None
    replan = None
    if not done["compiler_placed"] or not done.get("geometry_qa_passed"):
        try:
            geom = place_sheet_geometry(
                sheet_num=inputs.sheet_num,
                speed=inputs.speed,
                lane_width=inputs.lane_width,
                shoulder_width=inputs.shoulder_width,
                area_type=inputs.area_type or "",
                closure_type=inputs.closure_type or "",
                exposure_condition=inputs.exposure_condition or "",
                protective_vehicle_gvw=inputs.protective_vehicle_gvw or 0,
                align_idxs=req,
                outward_sign=outward_sign,
                arrow_panel_choice=arrow_panel_choice,
                dry_run=False,
                force=force,
            )
        except Exception as e:
            replan = _replan_after_failure(
                "place_sheet_geometry", {"failures": [str(e)]})
            phases.append({"phase": "place_sheet_geometry", "error": str(e),
                           "replan": replan})
            out = {
                "status": "ERROR",
                "sheetPlanActive": True,
                "sheet": inputs.sheet_num,
                "phases": phases,
                "failedPhase": "place_sheet_geometry",
                "replan": replan,
                "planStatus": get_plan_status(),
            }
            return _attach_plan_next(out)
        sc = geom.get("scorecard") or {}
        phases.append({"phase": "place_sheet_geometry", "result": {
            "status": geom.get("status"),
            "gateFailures": geom.get("gateFailures"),
            "placedCount": geom.get("placedCount"),
            "counts": geom.get("counts"),
            "deferred": geom.get("deferred"),
            "scorecardPassed": sc.get("passed"),
            "scorecardFailures": sc.get("failures"),
        }})
        if not sc.get("passed") and not force:
            _append_guide_cleanup(phases)
            replan = _replan_after_failure("place_sheet_geometry", {
                "failures": sc.get("failures") or [],
                "gateFailures": geom.get("gateFailures") or [],
            })
            out = {
                "status": "ERROR",
                "sheetPlanActive": True,
                "sheet": inputs.sheet_num,
                "phases": phases,
                "failedPhase": "place_sheet_geometry",
                "replan": replan,
                "planStatus": get_plan_status(),
            }
            return _attach_plan_next(out)
    else:
        phases.append({"phase": "place_sheet_geometry", "skipped": True})

    # Drop white align + perp ticks before QA (straight and curved).
    _append_guide_cleanup(phases)

    ledger_rec = {}
    try:
        ledger_rec = _record_sheet_build_ledger(path_vertices) or {}
    except Exception:
        ledger_rec = {}

    # --- visual QA ---
    qa = None
    if include_visual_qa and not _PLAN_SESSION.visual_qa_passed:
        try:
            qa = run_visual_qa_captures(force=force)
            phases.append({"phase": "visual_qa", "result": {
                "status": qa.get("status"),
                "visualQaPassed": qa.get("visualQaPassed"),
                "frames": [c.get("frame") for c in (qa.get("captures") or [])],
                "checklist": qa.get("checklist"),
                "failures": qa.get("failures"),
            }})
            if str(qa.get("status", "")).upper() == "ERROR" and not force:
                replan = qa.get("replan") or _replan_after_failure(
                    "visual_qa", {"failures": qa.get("failures") or []})
                out = {
                    "status": "ERROR",
                    "sheetPlanActive": True,
                    "sheet": inputs.sheet_num,
                    "phases": phases,
                    "failedPhase": "visual_qa",
                    "replan": replan,
                    "planStatus": get_plan_status(),
                }
                return _attach_plan_next(out)
        except Exception as e:
            replan = _replan_after_failure("visual_qa", {"failures": [str(e)]})
            phases.append({"phase": "visual_qa", "error": str(e), "replan": replan})
            out = {
                "status": "ERROR",
                "sheetPlanActive": True,
                "sheet": inputs.sheet_num,
                "phases": phases,
                "failedPhase": "visual_qa",
                "replan": replan,
                "planStatus": get_plan_status(),
            }
            return _attach_plan_next(out)
    elif not include_visual_qa:
        phases.append({"phase": "visual_qa", "skipped": True,
                       "note": "include_visual_qa=False"})
    else:
        phases.append({"phase": "visual_qa", "skipped": True})

    out = {
        "status": "OK",
        "sheetPlanActive": True,
        "sheet": inputs.sheet_num,
        "phases": phases,
        "planStatus": get_plan_status(),
        "highwayCaution": highway_caution,
        "overlapCaution": (overlap or {}).get("overlapCaution"),
    }
    if _PLAN_SESSION.real_road_edge:
        out["realRoadNext"] = (
            "If force/clear wiped striping, place_two_way_highway (or keep "
            "existing road) AFTER this build. delete_construction_guides "
            "already ran when real_road_edge was locked — re-call only if "
            "new ticks appeared. G20-2 stays on closed-shoulder EOP with "
            "other one-side signs."
        )
    # Lift QA captures to the top level so chat_driver can attach vision
    # the same way as a direct run_visual_qa_captures call.
    if isinstance(qa, dict) and qa.get("captures"):
        out["captures"] = qa["captures"]
        out["visualQaPassed"] = qa.get("visualQaPassed")
        out["checklist"] = qa.get("checklist")
        out["visionAttachedByChatDriver"] = True
    _attach_build_guide_fields(inputs.sheet_num, out)
    if ledger_rec:
        out["ledgerBuildId"] = ledger_rec.get("buildId")
    return _attach_plan_next(out)


def _alignment_bbox_pts(align_idxs: list[int]) -> list[tuple[float, float]]:
    pts: list[tuple[float, float]] = []
    for a in align_idxs:
        try:
            rows = get_alignment_vertices(int(a))
        except Exception:
            continue
        for r in rows:
            for kx, ky in (("sx", "sy"), ("ex", "ey")):
                if kx in r and ky in r:
                    try:
                        pts.append((float(r[kx]), float(r[ky])))
                    except (TypeError, ValueError):
                        pass
    return pts


def _replan_after_failure(phase: str, detail: dict | None = None) -> dict:
    """Map a failed sheet-build phase to resumeFrom / nextTool / fixHints.
    Preserves successful earlier phases (does not wipe them)."""
    detail = detail or {}
    failures = list(detail.get("failures") or detail.get("errors") or [])
    gates = list(detail.get("gateFailures") or [])
    blob = " ".join(str(x) for x in (failures + gates)).lower()

    if phase in ("assemble_corridor", "corridor"):
        resume, tool = "assemble_corridor", "assemble_corridor"
        hints = ["Re-pick distinct upstream/downstream WORK AREA edges",
                 "Then run_sheet_build(upstream_edge=..., downstream_edge=...)"]
    elif phase == "stations":
        resume, tool = "place_order_table_stations", "place_order_table_stations"
        hints = ["place_order_table_stations for missing aligns",
                 "Or run_sheet_build() to resume from stations"]
    elif phase == "signs":
        resume, tool = "place_sign", "place_sign"
        hints = ["Fix listed sign errors; use locked order-table sign_num",
                 "set_sign_attributes on createdElementIds"]
    elif phase in ("place_sheet_geometry", "compiler", "geometry_qa"):
        if "corridor-topology" in blob or "topology" in blob:
            resume, tool = "assemble_corridor", "assemble_corridor"
            hints = ["corridor-topology failure — rebuild edges via assemble_corridor",
                     "Then re-run place_sheet_geometry / run_sheet_build"]
        else:
            resume, tool = "place_sheet_geometry", "place_sheet_geometry"
            hints = ["Read scorecard.failures / gateFailures",
                     "Fix inputs or delete_placements then re-place",
                     "Call reflect_sheet_build for citations"]
    elif phase in ("visual_qa", "visual_qa_captures"):
        resume, tool = "run_visual_qa_captures", "get_geometry_scorecard"
        hints = ["Scorecard must pass before visual_qa_passed can be True",
                 "get_geometry_scorecard → fix → place_sheet_geometry → "
                 "run_visual_qa_captures"]
    else:
        resume, tool = "get_plan_status", "get_plan_status"
        hints = ["call get_plan_status / reflect_sheet_build"]

    replan = {
        "failedPhase": phase,
        "resumeFrom": resume,
        "nextTool": tool,
        "nextStep": "; ".join(hints),
        "fixHints": hints,
        "preservedPhases": [
            s["id"] for s in plan_workflow.PLAN_STAGES
            if plan_workflow.stage_done(_PLAN_SESSION).get(s["id"])
        ],
        "detailSample": (failures or gates)[:8],
    }
    _PLAN_SESSION.last_failed_phase = phase
    _PLAN_SESSION.last_replan = replan
    _save_sheet_plan()
    return replan


def get_geometry_scorecard(sheet_num: str = "") -> dict:
    """Return the last post-placement scorecard, or rebuild from registry.

    Prefer calling after place_sheet_geometry. Does not draw."""
    sheet = sheet_num
    if not sheet and _PLAN_SESSION.designer_inputs:
        sheet = _PLAN_SESSION.designer_inputs.sheet_num
    if _PLAN_SESSION.last_scorecard is not None:
        sc = dict(_PLAN_SESSION.last_scorecard)
        sc["source"] = "session"
        sc["sheetNum"] = sheet
        return sc
    rows = placement_registry.resolve_latest_placements(sheet_num=sheet or "")
    sc = sheet_scorecard.build_placement_scorecard(
        {"plan": {}, "gateFailures": [], "counts": {}},
        registry_rows=rows)
    sc["source"] = "registry_only"
    sc["note"] = (
        "No compiled plan in session — coverage check is registry-only. "
        "Re-run place_sheet_geometry for a full scorecard."
    )
    sc["sheetNum"] = sheet
    return sc


def reflect_sheet_build(max_iterations: int = 1) -> dict:
    """Structured reflection for the active sheet build (evaluator step).

    Deterministic critique citing registry primitiveIds / reqIds / scorecard
    failures. Caps iterations; appends to Bridge/sheet-reflection.jsonl.
    Does NOT auto-fix geometry — returns revision_instructions for the
    agent (or run_sheet_build resumeFrom)."""
    if not _PLAN_SESSION.sheet_plan_active():
        return {
            "status": "OK",
            "sheetPlanActive": False,
            "note": "No sheet plan — reflection is for named 619 builds only.",
        }
    max_iterations = max(1, min(int(max_iterations or 1), 3))
    sheet = _PLAN_SESSION.designer_inputs.sheet_num if _PLAN_SESSION.designer_inputs else ""
    rows = placement_registry.resolve_latest_placements(sheet_num=sheet)
    scorecard = _PLAN_SESSION.last_scorecard or sheet_scorecard.build_placement_scorecard(
        {"plan": {}, "gateFailures": [], "counts": {}}, registry_rows=rows)
    pre = sheet_scorecard.visual_qa_prechecks(
        scorecard, registry_rows=rows,
        sheet_geometry_placed=_PLAN_SESSION.sheet_geometry_placed)

    issues: list[str] = []
    revision: list[str] = []
    for f in (scorecard.get("failures") or []):
        issues.append(str(f))
        revision.append(f"Fix: {f}")
    for f in pre:
        if f not in issues:
            issues.append(f)
            revision.append(f"Fix: {f}")
    if not _PLAN_SESSION.sign_attrs_applied and _PLAN_SESSION.locked_sign_rows:
        issues.append("reflection: sign_attrs_applied is False")
        revision.append("set_sign_attributes on place_sign createdElementIds")
    if _PLAN_SESSION.last_failed_phase:
        issues.append(f"reflection: last_failed_phase={_PLAN_SESSION.last_failed_phase}")
        if _PLAN_SESSION.last_replan:
            revision.append(
                f"Resume from {_PLAN_SESSION.last_replan.get('resumeFrom')} "
                f"via {_PLAN_SESSION.last_replan.get('nextTool')}"
            )

    satisfactory = (
        not issues
        and bool(scorecard.get("passed"))
        and bool(_PLAN_SESSION.geometry_qa_passed)
    )
    citations = list(scorecard.get("citations") or [])[:24]
    if not citations:
        for r in rows[:24]:
            citations.append({
                "primitiveId": r.get("primitiveId"),
                "kind": r.get("kind"),
                "elementIds": r.get("elementIds") or [],
                "reqId": r.get("reqId") or "",
                "specRef": r.get("specRef") or {},
            })

    artifact = {
        "ts": _iso_now(),
        "sheetNum": sheet,
        "satisfactory": satisfactory,
        "issues": issues[:40],
        "revision_instructions": revision[:40],
        "citations": citations,
        "planStatus": {
            "geometryQaPassed": _PLAN_SESSION.geometry_qa_passed,
            "visualQaPassed": _PLAN_SESSION.visual_qa_passed,
            "compilerPlaced": _PLAN_SESSION.sheet_geometry_placed,
            "lastFailedPhase": _PLAN_SESSION.last_failed_phase or None,
        },
        "maxIterations": max_iterations,
    }
    _PLAN_SESSION.reflection_log.append(artifact)
    if len(_PLAN_SESSION.reflection_log) > 20:
        _PLAN_SESSION.reflection_log = _PLAN_SESSION.reflection_log[-20:]
    refl_path = _BRIDGE_DIR / "sheet-reflection.jsonl"
    try:
        refl_path.parent.mkdir(parents=True, exist_ok=True)
        with refl_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(artifact, separators=(",", ":")) + "\n")
    except OSError:
        pass

    out = {
        "status": "OK",
        "sheetPlanActive": True,
        "satisfactory": satisfactory,
        "issues": issues[:40],
        "revision_instructions": revision[:40],
        "citations": citations,
        "scorecardPassed": bool(scorecard.get("passed")),
        "lastReplan": _PLAN_SESSION.last_replan,
        "artifactPath": str(refl_path),
        "note": (
            "Reflection complete (deterministic). If satisfactory=False, follow "
            "revision_instructions and cite primitiveIds when deleting/fixing. "
            "Do not declare FINAL until visual_qa_passed after a passing scorecard."
            if not satisfactory else
            "Reflection satisfactory — run_visual_qa_captures if not yet passed, then FINAL."
        ),
    }
    return _attach_plan_next(out)


def run_visual_qa_captures(view_num: int = 1, force: bool = False) -> dict:
    """SHEET-PLAN ONLY: scripted visual QA captures + hard prechecks.

    Captures four frames from locked alignment geometry. Marks
    visual_qa_passed ONLY when the geometry scorecard passes and the
    placement registry has compiler artifacts (unless force=True).

    Outside a sheet plan: returns sheetPlanActive=False."""
    if not _PLAN_SESSION.sheet_plan_active():
        return {
            "status": "OK",
            "sheetPlanActive": False,
            "note": (
                "No sheet plan active — run_visual_qa_captures is a sheet-build "
                "tool. For general CAD use adjust_view + view_drawing freely."
            ),
        }
    if not _PLAN_SESSION.sheet_geometry_placed:
        plan_workflow.raise_plan_gate(
            "place_sheet_geometry has not succeeded yet.",
            tool="run_visual_qa_captures",
            current_step="compiler_placed",
            next_tool="place_sheet_geometry",
            next_step="Finish compiler path first, then run_visual_qa_captures",
        )

    sheet = (_PLAN_SESSION.designer_inputs.sheet_num
             if _PLAN_SESSION.designer_inputs else "")
    reg_rows = placement_registry.resolve_latest_placements(sheet_num=sheet)
    scorecard = _PLAN_SESSION.last_scorecard
    pre_fails = sheet_scorecard.visual_qa_prechecks(
        scorecard, registry_rows=reg_rows,
        sheet_geometry_placed=_PLAN_SESSION.sheet_geometry_placed,
        compiled=_PLAN_SESSION.last_compiled)
    if pre_fails and not force:
        _PLAN_SESSION.visual_qa_passed = False
        _PLAN_SESSION.visual_qa_failures = list(pre_fails)
        replan = _replan_after_failure("visual_qa", {"failures": pre_fails})
        return _attach_plan_next({
            "status": "ERROR",
            "sheetPlanActive": True,
            "visualQaPassed": False,
            "failures": pre_fails,
            "replan": replan,
            "note": (
                "visual_qa_passed NOT set — scorecard/registry prechecks failed. "
                "Fix failures, then re-call. Pass force=True only if the engineer "
                "accepts captures without a passing scorecard."
            ),
        })

    req = sorted(_PLAN_SESSION.required_aligns or _PLAN_SESSION.aligns_ready or {1, 2})
    pts = _alignment_bbox_pts(req)
    if len(pts) < 2:
        plan_workflow.raise_plan_gate(
            "could not read alignment vertices for framing.",
            tool="run_visual_qa_captures",
            current_step="visual_qa_passed",
            next_step="adopt/commit alignments, or pass force adjust_view manually",
        )

    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    cx = (min_x + max_x) / 2.0
    cy = (min_y + max_y) / 2.0
    full_w = max(max_x - min_x, 50.0) * 1.25
    full_h = max(max_y - min_y, 50.0) * 1.25
    span = max(max_x - min_x, 1.0)
    third = span / 3.0
    frames = [
        ("full_corridor", cx, cy, full_w, full_h),
        ("upstream", min_x + third * 0.5, cy, max(third * 1.4, 200.0), max(full_h * 0.7, 150.0)),
        ("work_area", cx, cy, max(third * 1.4, 200.0), max(full_h * 0.7, 150.0)),
        ("downstream", max_x - third * 0.5, cy, max(third * 1.4, 200.0), max(full_h * 0.7, 150.0)),
    ]

    captures: list[dict] = []
    checklist = [
        "Dims: length text above each tip-to-tip span",
        "Labels: feature names below where sheet calls for them",
        "PV at roll-ahead / protective-vehicle bay per sheet; AP per sheet; no AP/PV overlap",
        "Channelizing stops where sheet shows; hatch in work area only",
        "Ignore pre-existing site geometry far from the corridor",
        "Registry citations: use reflect_sheet_build / get_placements for IDs",
        "Chat agent: frames are attached as vision + panel SCREENSHOT — review them before FINAL",
    ]
    _PLAN_SESSION._qa_capture_active = True
    cap_dir = _BRIDGE_DIR / "captures"
    cap_dir.mkdir(parents=True, exist_ok=True)
    try:
        for name, fx, fy, fw, fh in frames:
            adjust_view(center_x=fx, center_y=fy, width=fw, height=fh,
                        view_num=view_num, force=True)
            cap = capture_view()
            src = Path(str(cap.get("path") or ""))
            durable = cap_dir / f"qa_{sheet or 'sheet'}_{name}.png"
            try:
                if src.is_file():
                    import shutil
                    shutil.copy2(src, durable)
                    path_out = str(durable)
                else:
                    path_out = str(src) if src else ""
            except OSError:
                path_out = str(src) if src else ""
            captures.append({
                "frame": name,
                "path": path_out,
                "centerX": fx, "centerY": fy, "width": fw, "height": fh,
            })
    finally:
        _PLAN_SESSION._qa_capture_active = False

    _PLAN_SESSION.visual_qa_passed = True
    _PLAN_SESSION.visual_qa_failures = list(pre_fails) if force else []
    if _PLAN_SESSION.last_failed_phase == "visual_qa":
        _PLAN_SESSION.last_failed_phase = ""
    _save_sheet_plan()
    out = {
        "status": "OK",
        "sheetPlanActive": True,
        "visualQaPassed": True,
        "forced": bool(force and pre_fails),
        "precheckFailuresIgnored": list(pre_fails) if force else [],
        "captures": captures,
        "checklist": checklist,
        "scorecardPassed": bool(scorecard.get("passed")) if scorecard else None,
        "visionAttachedByChatDriver": True,
        "note": (
            "Scripted visual QA complete (scorecard prechecks passed). "
            "Four frames are on disk under Bridge/captures/qa_*.png. "
            "In the in-MicroStation chat agent, chat_driver attaches those "
            "frames as vision + panel SCREENSHOT — review them against the "
            "checklist, fix only critical defects, then FINAL. "
            "Call reflect_sheet_build if you need registry citations. "
            "Do NOT call a non-existent capture_view tool; use view_drawing "
            "only for an extra ad-hoc look outside these frames."
        ),
    }
    return _attach_plan_next(out)


def begin_sheet_sandbox(upstream_edge: Optional[list[float]] = None,
                        downstream_edge: Optional[list[float]] = None,
                        offset_y_ft: float = 2000.0) -> dict:
    """Start a KEEP/REVERT sandbox on an offset Y band (does not wipe the
    kept corridor). Pass work-area edges or reuse PlanSession.work_area_edges.

    Next: run_sheet_build_sandbox() or run_sheet_build with returned edges.
    Then keep_sheet_sandbox() or revert_sheet_sandbox()."""
    if not _PLAN_SESSION.sheet_plan_active():
        return {
            "status": "ERROR",
            "note": "No sheet plan — build_wztc_order_table first.",
        }
    edges = _PLAN_SESSION.work_area_edges or {}
    up = upstream_edge or edges.get("upstream") or edges.get("upstream_edge")
    dn = downstream_edge or edges.get("downstream") or edges.get("downstream_edge")
    if not up or not dn:
        return {
            "status": "ERROR",
            "note": (
                "Need upstream_edge and downstream_edge (or prior "
                "work_area_edges on the plan). Point-pick edges first."
            ),
        }
    sheet = (_PLAN_SESSION.designer_inputs.sheet_num
             if _PLAN_SESSION.designer_inputs else "")
    out = sheet_sandbox.begin_sandbox(
        upstream_edge=list(up),
        downstream_edge=list(dn),
        offset_y_ft=float(offset_y_ft),
        sheet_num=sheet,
    )
    _PLAN_SESSION.sandbox = out.get("sandbox")
    _PLAN_SESSION.work_area_edges = {
        "upstream": out["upstream_edge"],
        "downstream": out["downstream_edge"],
        "sandbox": True,
        "bandId": (out.get("sandbox") or {}).get("bandId"),
    }
    # Fresh corridor on the sandbox band — reset placement checklist bits
    # that belong to the prior band without wiping designer inputs.
    _PLAN_SESSION.stations_placed_aligns = set()
    _PLAN_SESSION.signs_placed_rows = set()
    _PLAN_SESSION.sign_attrs_applied = False
    _PLAN_SESSION.sheet_geometry_placed = False
    _PLAN_SESSION.geometry_qa_passed = False
    _PLAN_SESSION.visual_qa_passed = False
    _PLAN_SESSION.aligns_ready = set()
    _PLAN_SESSION.last_station_rows = {}
    _save_sheet_plan()
    return _attach_plan_next(out)


def get_sheet_sandbox() -> dict:
    """Return active sandbox band state (Bridge/sandbox-state.json)."""
    out = sheet_sandbox.get_sandbox()
    out["sessionSandbox"] = _PLAN_SESSION.sandbox
    return out


def run_sheet_build_sandbox(offset_y_ft: float = 2000.0,
                            include_visual_qa: bool = True,
                            force: bool = False) -> dict:
    """begin_sheet_sandbox (if needed) + run_sheet_build on sandbox edges.

    Cheap try path: score with get_geometry_scorecard; KEEP or REVERT."""
    st = sheet_sandbox.get_sandbox()
    if not st.get("active"):
        edges = _PLAN_SESSION.work_area_edges or {}
        up = edges.get("upstream") or edges.get("upstream_edge")
        dn = edges.get("downstream") or edges.get("downstream_edge")
        if not up or not dn:
            return {
                "status": "ERROR",
                "note": (
                    "No active sandbox and no work_area_edges. Call "
                    "begin_sheet_sandbox(upstream_edge, downstream_edge) first."
                ),
            }
        # If edges already sandbox-flagged, use them; else begin.
        if not edges.get("sandbox"):
            begun = begin_sheet_sandbox(
                upstream_edge=list(up), downstream_edge=list(dn),
                offset_y_ft=offset_y_ft)
            if begun.get("status") == "ERROR":
                return begun
            up = begun["upstream_edge"]
            dn = begun["downstream_edge"]
        else:
            up, dn = edges["upstream"], edges["downstream"]
    else:
        sb = st["sandbox"]
        up = sb["sandboxUpstream"]
        dn = sb["sandboxDownstream"]
    result = run_sheet_build(
        upstream_edge=up, downstream_edge=dn,
        include_visual_qa=include_visual_qa, force=force,
        clear_prior_stations=True,
    )
    result["sandbox"] = sheet_sandbox.get_sandbox().get("sandbox")
    result["nextStepHint"] = (
        "Scorecard in result / get_geometry_scorecard. "
        "KEEP: keep_sheet_sandbox(). REVERT: revert_sheet_sandbox()."
    )
    return result


def keep_sheet_sandbox() -> dict:
    """Mark the active sandbox band as KEPT (prior reference band untouched)."""
    out = sheet_sandbox.keep_sandbox()
    if out.get("status") == "OK":
        _PLAN_SESSION.sandbox = out.get("sandbox")
        _save_sheet_plan()
    return out


def revert_sheet_sandbox() -> dict:
    """REVERT sandbox try: clear plan elements on the sandbox corridor and
    soft-delete registry heads created after sandbox start. Does not restore
    DGN of the reference band (it was never cleared)."""
    st = sheet_sandbox.get_sandbox()
    if not st.get("active") and not (st.get("sandbox") or {}).get("bandId"):
        return {"status": "ERROR", "note": "No sandbox band to revert."}
    sb = st.get("sandbox") or {}
    # Wipe sandbox placements (journal-owned). keep_alignments=False removes
    # the sandbox corridor lines too so a later try can redefine.
    cleared = clear_plan_elements(keep_alignments=False)
    # Soft-delete registry rows newer than checkpoint length heuristic:
    # delete all current heads for this sheet (sandbox-only build).
    sheet = sb.get("sheetNum") or (
        _PLAN_SESSION.designer_inputs.sheet_num
        if _PLAN_SESSION.designer_inputs else "")
    heads = placement_registry.resolve_latest_placements(sheet_num=sheet)
    pids = [str(r.get("primitiveId")) for r in heads if r.get("primitiveId")]
    deleted = 0
    if pids:
        deleted = placement_registry.mark_deleted(set(pids))
    marked = sheet_sandbox.mark_reverted({
        "cleared": cleared,
        "softDeletedPrimitiveIds": len(pids),
        "softDeletedCount": deleted,
    })
    _PLAN_SESSION.sandbox = marked
    _PLAN_SESSION.stations_placed_aligns = set()
    _PLAN_SESSION.signs_placed_rows = set()
    _PLAN_SESSION.sheet_geometry_placed = False
    _PLAN_SESSION.geometry_qa_passed = False
    _PLAN_SESSION.visual_qa_passed = False
    _PLAN_SESSION.aligns_ready = set()
    _save_sheet_plan()
    return {
        "status": "OK",
        "sandbox": marked,
        "cleared": cleared,
        "softDeleted": deleted,
        "note": (
            "Sandbox REVERTED — sandbox corridor/plan elements cleared. "
            "Reference band (pre-offset) was not touched. "
            "Call begin_sheet_sandbox again to retry."
        ),
    }


# ================================================================ Session

def undo_last_op() -> dict:
    """Undo the most recent undoable op. Draw ops: deletes exactly the
    elements they created (createdElementIds/elementId). M6 mutations:
    re-applies priorDeltaX/Y (move), priorLevel (level change), or
    priorText (text edit) from the journal. DELETE_ELEMENT is NOT
    undoable — no snapshot to restore; its journal row is skipped.
    Does NOT use MicroStation's own undo stack. Safe to call repeatedly:
    once an op is undone it's marked so a second call won't re-target it."""
    return _ok_or_raise(_bridge.call("UNDO_LAST_OP"), "undo_last_op")


_CLEAR_SKIP_OPS = frozenset({
    "CLEAR_PLAN_ELEMENTS", "DELETE_ELEMENT", "UNDO_LAST_OP",
    "BUILD_WZTC_ORDER_TABLE", "COMPUTE_SPACING", "GET_JOURNAL", "HANDOFF",
})
_CLEAR_ALIGN_OPS = frozenset({
    "DEFINE_ALIGNMENT_SEGMENT", "COMMIT_ALIGNMENT", "ADOPT_ALIGNMENT_ELEMENT",
})


def harvest_journal_create_ids(text: str, *, keep_alignments: bool = True,
                               align_idx: int = 0) -> set[str]:
    """Parse createdElementIds= from a wztc-journal.tsv body.

    Same REQ/RESP adjacency rules as ExecClearPlanElements: reqId reuse
    across process restarts must bind a RESP to the most recent REQ, not
    a global last-wins map. Used to recover IDs after journal rotation
    moves ownership proof into Bridge/archive/.
    """
    cur_op: dict[str, str] = {}
    cur_align: dict[str, int] = {}
    cur_undone: dict[str, bool] = {}
    ids: set[str] = set()
    for ln in (text or "").splitlines():
        parts = ln.split("\t")
        if len(parts) < 4:
            continue
        kind = parts[1].strip().upper()
        req = parts[2].strip()
        if kind == "REQ":
            cur_op[req] = parts[3].strip().upper()
            cur_undone[req] = False
            cur_align.pop(req, None)
            for p in parts[4:]:
                if p.startswith("alignIdx="):
                    raw = p.split("=", 1)[1].strip()
                    if raw.isdigit():
                        cur_align[req] = int(raw)
                    break
            continue
        if kind == "UNDONE":
            cur_undone[req] = True
            continue
        if kind != "RESP":
            continue
        if cur_undone.get(req):
            continue
        if parts[3].strip().upper() != "OK":
            continue
        op = cur_op.get(req, "")
        if keep_alignments and op in _CLEAR_ALIGN_OPS:
            continue
        if op in _CLEAR_SKIP_OPS:
            continue
        if align_idx and align_idx > 0:
            if cur_align.get(req) != int(align_idx):
                continue
        for p in parts:
            if p.startswith("createdElementIds="):
                csv = p.split("=", 1)[1]
                for one in csv.split(","):
                    one = one.strip()
                    if one:
                        ids.add(one)
    return ids


def clear_plan_elements(keep_alignments: bool = True, align_idx: int = 0) -> dict:
    """Delete journal-owned plan elements — the idempotent-rebuild wipe.

    Default keep_alignments=True leaves DEFINE_ALIGNMENT_SEGMENT /
    COMMIT_ALIGNMENT / ADOPT_ALIGNMENT_ELEMENT geometry alone so a rebuild
    reuses the same corridor. Pass False only when the engineer asked to
    wipe the corridor too.

    align_idx (>0): scope the wipe to create-ops journaled with that
    alignIdx= only (Upstream=1, Downstream=2). place_order_table_stations(
    clear_prior=True) uses this so rebuilding Downstream does not delete
    Upstream ticks/signs. align_idx=0 (default) clears the whole plan.
    Pass align_idx on place_sign so signs are included in scoped clears.

    After VBA CLEAR_PLAN (live journal only), also deletes IDs recorded in
    rotated Bridge/archive/wztc-journal-*.tsv and the placement registry.
    Journal rotation was leaving stacked dims/labels across rebuilds
    (engineer QA 2026-08-13).

    Does NOT fence-delete by proximity (that can catch engineer-drawn
    elements). Safe when nothing has been placed (deleted=0).

    Call this BEFORE re-placing stations/labels/dims/symbols/workspace/
    channelizing when iterating on a plan."""
    kwargs = {"keepAlignments": "Y" if keep_alignments else "N"}
    if align_idx and align_idx > 0:
        kwargs["alignIdx"] = align_idx
    resp = _ok_or_raise(
        _bridge.call("CLEAR_PLAN_ELEMENTS", **kwargs),
        "clear_plan_elements")
    if align_idx and align_idx > 0:
        _PLAN_SESSION.stations_placed_aligns.discard(align_idx)
        _PLAN_SESSION.signs_placed_rows = {
            (a, c) for a, c in _PLAN_SESSION.signs_placed_rows if a != align_idx
        }
        _PLAN_SESSION.last_station_rows.pop(int(align_idx), None)
        _PLAN_SESSION.sheet_geometry_placed = False
        _PLAN_SESSION.geometry_qa_passed = False
        _PLAN_SESSION.visual_qa_passed = False
        _PLAN_SESSION.sign_attrs_applied = False
    else:
        _PLAN_SESSION.placed_workspace = False
        _PLAN_SESSION.stations_placed_aligns = set()
        _PLAN_SESSION.signs_placed_rows = set()
        _PLAN_SESSION.sign_attrs_applied = False
        _PLAN_SESSION.sheet_geometry_placed = False
        _PLAN_SESSION.geometry_qa_passed = False
        _PLAN_SESSION.visual_qa_passed = False
        _PLAN_SESSION.last_station_rows = {}
        if not keep_alignments:
            _PLAN_SESSION.aligns_ready = set()
            _PLAN_SESSION.work_area_edges = None
            _PLAN_SESSION.work_bay_vertices = None
    # order_table_built stays True — SharedState still holds the table;
    # rebuild does not need to rebuild the table unless inputs change.
    if not (align_idx and align_idx > 0):
        placement_registry.clear_registry()
    _save_sheet_plan()
    return resp


_GUIDE_OPS = frozenset({
    "PLACE_ORDER_TABLE_STATIONS",
    "PLACE_PERP_LINE",
    "DEFINE_ALIGNMENT_SEGMENT",
    "COMMIT_ALIGNMENT",
})


def _append_guide_cleanup(phases: list) -> None:
    """Remove alignment lines + perp ticks. Always after a sheet draw."""
    try:
        guides = delete_construction_guides()
        phases.append({"phase": "delete_construction_guides", "result": {
            "status": guides.get("status"),
            "deleted": guides.get("deleted"),
            "candidateIds": guides.get("candidateIds"),
            "note": guides.get("note"),
        }})
    except Exception as e:
        phases.append({
            "phase": "delete_construction_guides",
            "error": str(e),
            "note": "Guide cleanup failed — call delete_construction_guides manually",
        })


def delete_construction_guides() -> dict:
    """Delete ONLY alignment centerlines and order-table perp tick lines.

    Parses Bridge/wztc-journal.tsv for create-ops in _GUIDE_OPS and deletes
    their createdElementIds. Does NOT delete signs, channelizing, hatch,
    dims, labels, AP/PV, or road striping. Use after a real-road sheet
    build when ticks/align lines are no longer needed for QA.
    """
    from pathlib import Path

    journal = Path(__file__).resolve().parent.parent / "Bridge" / "wztc-journal.tsv"
    if not journal.exists():
        return {"status": "OK", "deleted": 0, "note": "no journal file"}

    cur_op: dict[str, str] = {}
    cur_undone: dict[str, bool] = {}
    ids: set[str] = set()
    try:
        lines = journal.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError as e:
        return {"status": "ERROR", "deleted": 0, "note": str(e)}

    for ln in lines:
        parts = ln.split("\t")
        if len(parts) < 4:
            continue
        kind = parts[1].strip().upper()
        req = parts[2].strip()
        if kind == "REQ":
            cur_op[req] = parts[3].strip().upper()
            cur_undone[req] = False
            continue
        if kind == "UNDONE":
            cur_undone[req] = True
            continue
        if kind != "RESP":
            continue
        if cur_undone.get(req):
            continue
        if parts[3].strip().upper() != "OK":
            continue
        op = cur_op.get(req, "")
        if op not in _GUIDE_OPS:
            continue
        for p in parts:
            if p.startswith("createdElementIds="):
                csv = p.split("=", 1)[1]
                for one in csv.split(","):
                    one = one.strip()
                    if one:
                        ids.add(one)

    deleted = 0
    errors: list[str] = []
    for eid in sorted(ids, key=lambda x: float(x) if x.replace(".", "", 1).isdigit() else 0):
        try:
            r = delete_element(eid, reason="construction guide (align/tick) cleanup")
            if str(r.get("status", "")).upper() == "OK":
                deleted += int(r.get("deleted") or 1)
            else:
                errors.append(f"{eid}:{r.get('note')}")
        except Exception as e:
            errors.append(f"{eid}:{e}")

    return {
        "status": "OK" if not errors else "PARTIAL",
        "deleted": deleted,
        "candidateIds": len(ids),
        "ops": sorted(_GUIDE_OPS),
        "errors": errors[:12],
        "note": "Removed alignment lines + perp ticks only; sheet plan geometry kept.",
    }


def get_placements(sheet_num: str = "", kind: str = "", zone: str = "",
                   run: str = "", align_idx: int = 0) -> list[dict]:
    """List agent-placed primitives from Bridge/placement-registry.jsonl.

    Filter by sheet_num / kind (dimension, label, cone, sign, …) / zone /
    run / align_idx. Empty filters match all. Use this instead of fishing
    the journal for 'delete the channelizers' style edits."""
    rows = placement_registry.load_placements(
        sheet_num=sheet_num, kind=kind, zone=zone, run=run, align_idx=align_idx)
    if not rows:
        return [{
            "note": (
                "No matching placement-registry records. Registry only tracks "
                "agent-placed geometry from place_sheet_geometry / place_sign."
            )
        }]
    return rows


def delete_placements(kind: str = "", zone: str = "", run: str = "",
                      align_idx: int = 0, sheet_num: str = "",
                      reason: str = "") -> dict:
    """Delete DGN elements for matching placement-registry records.

    Resolves elementIds from the registry, calls delete_element per id,
    then removes those registry lines. Stale IDs (hand-edited/deleted)
    are reported under failed — no auto-rebind."""
    rows = placement_registry.load_placements(
        sheet_num=sheet_num, kind=kind, zone=zone, run=run, align_idx=align_idx)
    if not rows or (len(rows) == 1 and rows[0].get("note")):
        return {"status": "OK", "deleted": 0, "note": "no matching placements"}
    deleted: list[str] = []
    failed: list[dict] = []
    gone_prims: set[str] = set()
    for rec in rows:
        pid = str(rec.get("primitiveId") or "")
        ids = [str(x) for x in (rec.get("elementIds") or [])]
        all_ok = True
        for eid in ids:
            try:
                delete_element(eid, own_element_only=True,
                               reason=reason or f"delete_placements {pid}")
                deleted.append(eid)
            except Exception as e:
                all_ok = False
                failed.append({"elementId": eid, "primitiveId": pid, "error": str(e)})
        if all_ok and pid:
            gone_prims.add(pid)
    removed = placement_registry.mark_deleted(gone_prims)
    return {
        "status": "OK" if not failed else "PARTIAL",
        "deletedCount": len(deleted),
        "deleted": deleted[:80],
        "failed": failed[:40],
        "registryRemoved": removed,
    }

def get_journal(limit: int = 20) -> list[str]:
    """Return the last `limit` raw journal lines — every op run this
    session, its full parameters (including any reason= passed), and its
    result. This is the PE audit trail: use it to answer "why is that sign
    there" or to review recent ops before handing off a sheet.

    `limit` is clamped to MAX_JOURNAL_LINES — journal lines are verbose and
    a limit=150 call was measured stuffing ~20K chars into one tool result.
    Do not use the journal to locate an element the engineer can click;
    prefer ask_user_choice(allow_point_pick=True)."""
    clamped = max(1, min(int(limit), MAX_JOURNAL_LINES))
    resp = _ok_or_raise(_bridge.call("GET_JOURNAL", limit=clamped), "get_journal")
    lines = [row.get("line", "") for row in resp.get("rows", [])]
    if int(limit) > MAX_JOURNAL_LINES:
        lines.append(
            f"[truncated: requested limit={limit}, returned last {MAX_JOURNAL_LINES}. "
            "Ask a narrower question or use ask_user_choice(allow_point_pick=True) "
            "instead of paging the whole journal.]"
        )
    return lines


def list_deferred_handoffs() -> list[dict]:
    """List every dimension/callout queued by handoff() this session that
    still needs a few manual clicks through the existing interactive forms."""
    resp = _ok_or_raise(_bridge.call("LIST_DEFERRED_HANDOFFS"), "list_deferred_handoffs")
    return resp.get("rows", [])


# ============================================================ Registry / Edit (M6)

def list_registry_commands(safety_status: str = "", opname_contains: str = "") -> list[dict]:
    """List MicroStation command recipes in Data/command-registry.tsv.
    Optional safety_status filter (e.g. 'verified-headless-safe',
    'needs-testing', 'interactive-only-use-handoff'). Only
    verified-headless-safe rows can be executed via run_registry_command;
    interactive-only rows point at handoff() instead.

    opname_contains narrows to opNames containing this substring
    (case-insensitive) -- e.g. 'ZOOM', 'PAN', 'LEVEL'. Strongly
    recommended: this registry has ~1800 rows (~1600
    verified-headless-safe), and returning them all costs real tokens --
    an unfiltered call was measured live at ~240K input tokens (~$0.75)
    for a single turn. If you have any idea what the command name might
    contain, pass it here rather than listing everything. If the
    (post-filter) result set still exceeds MAX_LISTED_ROWS, only the
    first that many come back, with a note appended -- narrow further
    with opname_contains rather than assuming you've seen every match."""
    params = {}
    if safety_status:
        params["safetyStatus"] = safety_status
    resp = _ok_or_raise(_bridge.call("LIST_REGISTRY_COMMANDS", **params), "list_registry_commands")
    rows = resp.get("rows", [])

    if opname_contains:
        needle = opname_contains.strip().upper()
        rows = [r for r in rows if needle in str(r.get("opName", "")).upper()]

    total = len(rows)
    if total > MAX_LISTED_ROWS:
        rows = rows[:MAX_LISTED_ROWS]
        rows.append({
            "note": f"{total} rows matched -- showing first {MAX_LISTED_ROWS}. "
                    "Narrow further with opname_contains instead of assuming this is everything."
        })
    return rows


def describe_registry_command(op_name: str) -> dict:
    """Return the full registry row for one opName — safetyStatus,
    recipeLines, requiredParams, notes, sourceRefs, etc."""
    return _ok_or_raise(
        _bridge.call("DESCRIBE_REGISTRY_COMMAND", opName=op_name),
        "describe_registry_command")


def run_registry_command(op_name: str, params: Optional[dict] = None, reason: str = "") -> dict:
    """Run a verified-headless-safe keyin_recipe from the command
    registry. Refuses needs-testing / interactive-only / unsafe-blocked
    rows with a clear ERROR (not a silent no-op). Pass recipe params in
    `params` matching requiredParams (e.g. {"level": "Default"} for
    ACTIVE_LEVEL, {"color": "0"} for ACTIVE_COLOR). direct_api rows
    (MOVE_ELEMENT etc.) must use their dedicated tools below — this
    entry point will refuse them."""
    call_params = {"opName": op_name, "reason": reason}
    if params:
        call_params.update(params)
    return _ok_or_raise(_bridge.call("RUN_REGISTRY_COMMAND", **call_params), "run_registry_command")


def move_element(element_id: str, delta_x: float, delta_y: float, delta_z: float = 0,
                  own_element_only: bool = True, reason: str = "") -> dict:
    """Move an element by delta_x/delta_y(/delta_z) design units (ft).
    By default own_element_only=True: element_id must appear in this
    session's journal as something the agent created. Response includes
    priorDeltaX/Y for undo_last_op."""
    return _ok_or_raise(
        _bridge.call("MOVE_ELEMENT", elementId=element_id, deltaX=delta_x, deltaY=delta_y,
                     deltaZ=delta_z, ownElementOnly=("Y" if own_element_only else "N"),
                     reason=reason),
        "move_element")


def copy_element(element_id: str, delta_x: float, delta_y: float, delta_z: float = 0,
                  own_element_only: bool = True, reason: str = "") -> dict:
    """Copy an element by ID (Clone + Move). Returns newElementId /
    createdElementIds. own_element_only defaults True (journal gate).

    When the engineer picked a pre-existing site/base element (not one you
    created this session), you MUST pass own_element_only=False or the
    copy is refused. Resolve geometry first with get_elements_range([id])
    so deltas are computed from real bbox coords, not guessed."""
    return _ok_or_raise(
        _bridge.call("COPY_ELEMENT", elementId=element_id, deltaX=delta_x, deltaY=delta_y,
                     deltaZ=delta_z, ownElementOnly=("Y" if own_element_only else "N"),
                     reason=reason),
        "copy_element")


def rotate_element(element_id: str, origin_x: float, origin_y: float, angle_deg: float,
                    origin_z: float = 0, own_element_only: bool = True, reason: str = "") -> dict:
    """Rotate an element about (origin_x, origin_y) by angle_deg (Z axis).
    Response includes priorAngleDeg for undo_last_op."""
    return _ok_or_raise(
        _bridge.call("ROTATE_ELEMENT", elementId=element_id, originX=origin_x, originY=origin_y,
                     originZ=origin_z, angleDeg=angle_deg,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "rotate_element")


def scale_element(element_id: str, origin_x: float, origin_y: float, scale_factor: float,
                   origin_z: float = 0, own_element_only: bool = True, reason: str = "") -> dict:
    """Uniform-scale an element about a point (Element.ScaleUniform).
    Response includes priorScaleFactor (1/factor) for undo."""
    return _ok_or_raise(
        _bridge.call("SCALE_ELEMENT", elementId=element_id, originX=origin_x, originY=origin_y,
                     originZ=origin_z, scaleFactor=scale_factor,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "scale_element")


def mirror_element(element_id: str, x1: float, y1: float, x2: float, y2: float,
                    z1: float = 0, z2: float = 0, own_element_only: bool = True,
                    reason: str = "") -> dict:
    """Mirror an element about the axis through (x1,y1)-(x2,y2).
    Re-run the same mirror to undo."""
    return _ok_or_raise(
        _bridge.call("MIRROR_ELEMENT", elementId=element_id, x1=x1, y1=y1, x2=x2, y2=y2,
                     z1=z1, z2=z2, ownElementOnly=("Y" if own_element_only else "N"),
                     reason=reason),
        "mirror_element")


def array_element(element_id: str, count: int, spacing_x: float, spacing_y: float,
                   own_element_only: bool = True, reason: str = "") -> dict:
    """Create `count` copies offset by i*(spacing_x, spacing_y).
    Returns newElementIds / createdElementIds."""
    return _ok_or_raise(
        _bridge.call("ARRAY_ELEMENT", elementId=element_id, count=count,
                     spacingX=spacing_x, spacingY=spacing_y,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "array_element")


def change_element_level(element_id: str, level: str, own_element_only: bool = True,
                          reason: str = "") -> dict:
    """Change an element's level by ID. own_element_only defaults True
    (journal-gated). Response includes priorLevel for undo_last_op."""
    return _ok_or_raise(
        _bridge.call("CHANGE_ELEMENT_LEVEL", elementId=element_id, level=level,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "change_element_level")


def edit_text_element(element_id: str, new_text: str, own_element_only: bool = True,
                       reason: str = "") -> dict:
    """Replace text on a TextElement / TextNodeElement by ID.
    own_element_only defaults True. Response includes priorText for
    undo_last_op."""
    return _ok_or_raise(
        _bridge.call("EDIT_TEXT_ELEMENT", elementId=element_id, newText=new_text,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "edit_text_element")


def delete_element(element_id: str, own_element_only: bool = True, reason: str = "") -> dict:
    """Delete an element by ID. own_element_only defaults True.
    NOT undoable via undo_last_op (no snapshot to restore) — the
    response says so plainly (notUndoable=Y). Prefer undo_last_op on
    the placing op when you still can."""
    return _ok_or_raise(
        _bridge.call("DELETE_ELEMENT", elementId=element_id,
                     ownElementOnly=("Y" if own_element_only else "N"), reason=reason),
        "delete_element")
