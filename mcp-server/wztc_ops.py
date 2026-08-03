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

from typing import Optional

import sheet_spec
import view_capture

_bridge = None

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

def find_elements_near(x: float, y: float, radius: float, type_filter: str = "") -> list[dict]:
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
    over a fishing expedition."""
    resp = _ok_or_raise(
        _bridge.call("FIND_ELEMENTS_NEAR", x=x, y=y, radius=radius, typeFilter=type_filter),
        "find_elements_near")
    return _cap_spatial_rows(resp.get("rows", []), "find_elements_near", radius)


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


def list_levels(name_contains: str = "") -> list[dict]:
    """List levels in the active design file matching name_contains
    (case-insensitive substring, e.g. 'TWZ', 'Traffic', 'SF_P').
    name_contains is REQUIRED — this file can have thousands of levels
    (measured live at 3046); an unfiltered dump costs real tokens and
    still won't surface the level you want if it isn't in the first page.
    Results are hard-capped at MAX_LISTED_ROWS matches. Returns name,
    number, isDisplayed."""
    needle = (name_contains or "").strip()
    if not needle:
        return [{
            "status": "ERROR",
            "note": "list_levels requires name_contains (e.g. 'TWZ', 'SFB', "
                    "'Traffic'). Refusing unfiltered listing — this DGN can "
                    "have thousands of levels.",
        }]
    resp = _ok_or_raise(_bridge.call("LIST_LEVELS"), "list_levels")
    rows = resp.get("rows", [])
    upper = needle.upper()
    rows = [r for r in rows if upper in str(r.get("name", "")).upper()]
    total = len(rows)
    if total > MAX_LISTED_ROWS:
        rows = rows[:MAX_LISTED_ROWS]
        rows.append({
            "note": (
                f"{total} levels matched name_contains={name_contains!r} -- "
                f"showing first {MAX_LISTED_ROWS}. Tighten the filter."
            )
        })
    return rows


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
                 view_num: int = 1) -> dict:
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

    zoom_out_percent: e.g. 40 zooms OUT so ~40% more area becomes visible
    (new width/height = current * 1.40). Negative zooms IN (e.g. -40 =
    40% less area, current * 0.60). Must be > -100. 0 = no zoom change.
    pan_x / pan_y: shift the view center by this many design units (the
    file's working units, typically feet) in the model's X/Y direction.
    0 = no pan. Positive pan_x moves the visible center east/right,
    positive pan_y moves it north/up (standard model-space convention).

    Takes ~2 seconds to settle before returning (MicroStation's repaint
    isn't synchronous with the property write) -- call capture_view or
    the chat agent's view_drawing afterward to see the result."""
    state = view_capture.get_view_state(view_num=view_num)

    scale = 1.0 + (zoom_out_percent / 100.0)
    if scale <= 0:
        return {"status": "ERROR",
                "note": f"zoom_out_percent={zoom_out_percent} would produce a non-positive "
                        "scale factor -- must be greater than -100."}

    new_width = state["width"] * scale
    new_height = state["height"] * scale
    new_center_x = state["centerX"] + pan_x
    new_center_y = state["centerY"] + pan_y

    view_capture.navigate_view(new_center_x, new_center_y, new_width, new_height,
                                z=state["centerZ"], view_num=view_num)

    return {
        "status": "OK",
        "previousWidth": state["width"], "previousHeight": state["height"],
        "newWidth": new_width, "newHeight": new_height,
        "centerX": new_center_x, "centerY": new_center_y,
    }


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
    than guessing."""
    resp = _bridge.call("GET_SHEET_REQUIREMENTS", sheetNum=sheet_num)
    if resp["status"] == "ERROR":
        return {"found": False, "note": resp.get("note", "")}
    resp["found"] = True
    return resp


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

# In-process plan-session flags (chat_driver process lifetime). Soft
# memory so place_perp_line can refuse the incomplete sketch pattern that
# shipped live 2026-08-02 (workspace + alignment + one sign + one tick,
# declared "done" with no order table). Cleared by reset_plan_session_flags
# (exit_mode) or rebuilt when build_wztc_order_table runs.
_PLAN_SESSION: dict = {
    "placed_workspace": False,
    "order_table_built": False,
    "stations_placed_aligns": set(),
}


def reset_plan_session_flags() -> None:
    """Drop plan-flow memory (call from exit_mode so a later general/wztc
    task doesn't inherit a prior plan's gate state)."""
    _PLAN_SESSION["placed_workspace"] = False
    _PLAN_SESSION["order_table_built"] = False
    _PLAN_SESSION["stations_placed_aligns"] = set()


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
        if _PLAN_SESSION["order_table_built"] and align_idx not in _PLAN_SESSION["stations_placed_aligns"]:
            raise ValueError(
                f"Order table exists but place_order_table_stations has not been called "
                f"for align_idx={align_idx} yet. Call place_order_table_stations instead "
                f"(not place_perp_line item-by-item). Pass one_off=True only if the "
                f"engineer explicitly asked for a single ad-hoc tick outside the order table."
            )
        if _PLAN_SESSION["placed_workspace"] and not _PLAN_SESSION["order_table_built"]:
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
               reason: str = "") -> dict:
    """Place a sign assembly (post + edge-connected stem + face + label).

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
    """
    if side.strip().lower() == "both sides":
        missing = [n for n, v in
                   [("pt2x", pt2x), ("pt2y", pt2y), ("dir2x", dir2x), ("dir2y", dir2y)] if v is None]
        if missing:
            raise ValueError(f"side='Both Sides' requires {missing}")
    return _ok_or_raise(
        _bridge.call("PLACE_SIGN", signNum=sign_num, roadType=road_type, side=side,
                     pt1X=pt1x, pt1Y=pt1y, pt1Z=pt1z, dir1X=dir1x, dir1Y=dir1y,
                     pt2X=pt2x, pt2Y=pt2y, pt2Z=pt2z, dir2X=dir2x, dir2Y=dir2y,
                     reason=reason),
        "place_sign")


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
    _PLAN_SESSION["placed_workspace"] = True
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
                            exposure_condition: str = "") -> dict:
    """Headless equivalent of WZTCDesigner.frm's Submit & Draw — builds the
    full per-alignment order table and writes the same SharedState the manual
    form writes. Never estimate spacing yourself; it comes from here.

    If Data/sheet-specs/<sheet_num>.json exists, THE SHEET DRIVES THE TABLE:
    the station sequence, every spacing, the sign order and each sign's
    SignLibrary key all come from the sheet, and sign_rows/area_type are only
    needed to disambiguate. This is the correct path — the generic fallback
    below emits the same 7 upstream rows for every sheet, including stations
    (Vehicle Space, temporary barrier, box/corr beam) that 619-311 does not
    have, and interpolates shoulder taper values Table 311-02 doesn't print.
    Pass area_type ("URBAN"/"RURAL") whenever a spec exists — the sheet's
    advance-sign spacing and sign legends both depend on it.

    Without a spec, falls back to WZTCRules defaults and marks the result
    specDriven=False, which you should relay to the engineer rather than
    presenting the table as sheet-faithful.

    sign_rows (optional when a spec exists): list of dicts, each
    {"align_idx": 1|2, "sign_num": SignLibrary key, "side": "One Side"|
    "Both Sides", "spacing_ft": optional, "size": optional}.

    Returns the order table (rows: alignIdx, alignName, rowNum, type, label,
    spacing, size, side) — show it to the engineer before drawing."""
    sign_rows = list(sign_rows or [])
    spec_rows_tsv = ""
    overrides_tsv = ""
    spec_info: dict = {"specDriven": False}

    spec = sheet_spec.load(sheet_num) if sheet_num else None
    if spec is not None:
        if not area_type:
            raise ValueError(
                f"sheet {sheet_num} has a spec whose advance-sign spacing and sign "
                f"legends depend on area type; pass area_type='URBAN' or 'RURAL' "
                f"(Table {sheet_num.split('-')[1]}-03).")
        resolved = sheet_spec.resolve(
            spec, speed, lane_width, shoulder_width, area_type,
            closure_type or None, exposure_condition or None)
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
        overrides_tsv = "|".join([
            f"bufferSpace={resolved['bufferFt']}",
            f"mergingTaper={resolved['laneTaper']['ft']}",
            f"shoulderTapers={resolved['shoulderTaper']['ft']}",
            f"rollAhead={resolved['rollAheadFt']['min']}",
            f"laneTaperSkips={resolved['laneTaper']['skipLines']}",
            f"shoulderTaperSkips={resolved['shoulderTaper']['skipLines']}",
            f"laneTaperDevices={resolved['laneTaper']['devices']}",
            f"shoulderTaperDevices={resolved['shoulderTaper']['devices']}",
        ])
        spec_info = {
            "specDriven": True,
            "sheet": spec["sheet"]["number"],
            "shoulderBandUsed": resolved["shoulderBand"],
            "signLegends": resolved["legend"],
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
    _PLAN_SESSION["order_table_built"] = True
    _PLAN_SESSION["stations_placed_aligns"] = set()
    return resp


def find_reference_linework(level_name_contains: str, include_references: bool = False,
                            ref_name_contains: str = "") -> list[dict]:
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
    place_workspace with no re-encoding."""
    resp = _ok_or_raise(
        _bridge.call("FIND_REFERENCE_LINEWORK", levelNameContains=level_name_contains,
                     includeReferences="Y" if include_references else "N",
                     refNameContains=ref_name_contains),
        "find_reference_linework")
    return resp.get("rows", [])


def define_alignment_segment(align_idx: int, vertices: list[list[float]], reason: str = "") -> dict:
    """Create straight alignment line segments from vertices (Default
    level/color 0/weight 0) and record them as one drawing session for
    align_idx — the same bookkeeping AlignDraw's interactive clicking
    produces. vertices come from find_reference_linework's verticesTSV
    (parsed back to [[x,y,z],...]) or from repeated ask_user_choice
    point-picks when no usable reference geometry exists. Call this one
    or more times per alignment, then commit_alignment once when done.
    align_idx convention: 1=Upstream, 2=Downstream (matches
    build_wztc_order_table)."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(
        _bridge.call("DEFINE_ALIGNMENT_SEGMENT", alignIdx=align_idx, verticesTSV=verts_tsv, reason=reason),
        "define_alignment_segment")


def commit_alignment(align_idx: int) -> dict:
    """Group every segment recorded by define_alignment_segment for
    align_idx into a graphic group, marking that alignment ready for
    place_order_table_stations. Call once per alignment after all its
    define_alignment_segment calls."""
    resp = _ok_or_raise(_bridge.call("COMMIT_ALIGNMENT", alignIdx=align_idx), "commit_alignment")
    if not _PLAN_SESSION["order_table_built"]:
        resp["nextStep"] = (
            "build_wztc_order_table (show engineer), then place_order_table_stations — "
            "do not place_perp_line/place_sign by hand for plan stations"
        )
    elif align_idx not in _PLAN_SESSION["stations_placed_aligns"]:
        resp["nextStep"] = f"place_order_table_stations(align_idx={align_idx}, reset_session=...)"
    return resp


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
    clear_prior=True calls clear_plan_elements() first (keeps alignments).
    Use this when rebuilding — without it, a second place stacks ticks /
    cells / channelizing on top of the previous run (the non-idempotent
    failure mode). If stations were already placed for this align_idx
    this session and clear_prior/force are both False, this refuses.
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
    already = align_idx in _PLAN_SESSION["stations_placed_aligns"]
    if already and not clear_prior and not force:
        raise ValueError(
            f"stations already placed for align_idx={align_idx} this session. "
            f"Call clear_plan_elements() (or pass clear_prior=True) before "
            f"rebuilding — otherwise ticks/cells/channelizing stack on the "
            f"previous run. Pass force=True only for intentional additive placement."
        )
    cleared = None
    if clear_prior:
        cleared = clear_plan_elements(keep_alignments=True)
    resp = _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_STATIONS", alignIdx=align_idx,
                     resetSession="Y" if reset_session else "N"),
        "place_order_table_stations")
    _PLAN_SESSION["stations_placed_aligns"].add(align_idx)
    if cleared is not None:
        resp["clearedPrior"] = cleared
    return resp


def place_order_table_labels(align_idx: int, outward_sign: float = -1.0,
                             text_extra_along: float = 20.0,
                             sheet_elements: str = "") -> dict:
    """Name labels BELOW tip-to-tip dims (X-centered). sheet_elements from
    get_sheet_requirements gates optional tapers (Must include ShoulderTaper
    when the official sheet shows it). Core Roll Ahead / Vehicle Space /
    Buffer always. Dims are separate — place_order_table_dimensions does
    every tick and is not sheet-gated."""
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_LABELS", alignIdx=align_idx,
                     outwardSign=outward_sign, textExtraAlong=text_extra_along,
                     sheetElements=sheet_elements),
        "place_order_table_labels")


def place_order_table_dimensions(align_idx: int, outward_sign: float = -1.0,
                                 offset_dist: float = 15.0,
                                 sheet_elements: str = "") -> dict:
    """Real ny_Plan Linear Size dims tip-to-tip between EVERY consecutive
    tick (including Sign spacings). Length above the dim line.
    sheet_elements is not used for gating (API compat only)."""
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_DIMENSIONS", alignIdx=align_idx,
                     outwardSign=outward_sign, offsetDist=offset_dist,
                     sheetElements=sheet_elements),
        "place_order_table_dimensions")


def place_sheet_symbol_cells(align_idx: int, sheet_elements: str,
                             outward_sign: float = -1.0) -> dict:
    """ProtectiveVehicle→TWZWVA_P in Vehicle Space bay; ArrowPanel→TWZAP_P
    at Shoulder Taper tip (sheet callout; fallback Merging Taper)."""
    return _ok_or_raise(
        _bridge.call("PLACE_SHEET_SYMBOL_CELLS", alignIdx=align_idx,
                     sheetElements=sheet_elements, outwardSign=outward_sign),
        "place_sheet_symbol_cells")


def place_order_table_workspace(align_idx: int, outward_sign: float = -1.0,
                                lane_width: float = 12.0) -> dict:
    """Hatched work-space box in the closed lane from path start through
    Vehicle Space end (sheet work bay — not freeform vertices)."""
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_WORKSPACE", alignIdx=align_idx,
                     outwardSign=outward_sign, laneWidth=lane_width),
        "place_order_table_workspace")


def place_order_table_channelizing(align_idx: int, outward_sign: float = -1.0,
                                   lane_width: float = 12.0) -> dict:
    """Sheet-bounded channelizing: shoulder/merging taper diagonals +
    longitudinal closed-lane run from taper toe to path start. Does not
    use freeform AccuDraw-length vertices."""
    return _ok_or_raise(
        _bridge.call("PLACE_ORDER_TABLE_CHANNELIZING", alignIdx=align_idx,
                     outwardSign=outward_sign, laneWidth=lane_width),
        "place_order_table_channelizing")


def place_dimension(x1: float, y1: float, x2: float, y2: float,
                    ox: float, oy: float, z: float = 0.0,
                    style_name: str = "ny_Plan", reason: str = "") -> dict:
    """Place a real Linear Size DimensionElement (msdDimTypeSizeArrow)
    between (x1,y1)-(x2,y2); dim-line offset toward (ox,oy). Uses DGN
    DimensionStyles (default ny_Plan). Prefer place_order_table_dimensions
    for full-plan spacing annotation."""
    return _ok_or_raise(
        _bridge.call("PLACE_DIMENSION", x1=x1, y1=y1, x2=x2, y2=y2,
                     ox=ox, oy=oy, z=z, styleName=style_name, reason=reason),
        "place_dimension")


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


def place_text_label(text: str, x: float, y: float, z: float = 0.0, reason: str = "") -> dict:
    """Place a single-line text label via TEXTEDITOR PLACE + INSERT_TEXT."""
    return _ok_or_raise(
        _bridge.call("PLACE_TEXT_LABEL", text=text, x=x, y=y, z=z, reason=reason),
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


def place_cell(cell_name: str, pt_x: float, pt_y: float, pt_z: float = 0, angle_deg: float = 0,
               reason: str = "") -> dict:
    """Place a single cell from the WZTC symbol library at (pt_x, pt_y, pt_z).
    cell_name must be an exact library cell (e.g. 'TWZIA_P') — call
    list_cells / attach_cell_library first if unsure. place_cell itself
    re-attaches the default WZTC .cel before placing."""
    return _ok_or_raise(
        _bridge.call("PLACE_CELL", cellName=cell_name, ptX=pt_x, ptY=pt_y, ptZ=pt_z,
                     angleDeg=angle_deg, reason=reason),
        "place_cell")


def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Finish symbology after place_sign. Labels/text → SF_P white wt=3;
    stems → SF_P white; post cell TWZSGN_P → color 6 (orange). Face cells
    are intentionally LEFT ALONE (library SF_P/SFB_P + ByCell weights) —
    do NOT also call change_element_symbology on faces (forcing color 0/6
    or weight 3 bleaches or wrecks the legend; live 2026-08-03).
    element_ids from place_sign createdElementIds; applied count may be
    less than requested because faces are skipped on purpose."""
    return _ok_or_raise(
        _bridge.call("SET_SIGN_ATTRIBUTES", elementIds=",".join(element_ids), reason=reason),
        "set_sign_attributes")


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


def clear_plan_elements(keep_alignments: bool = True) -> dict:
    """Delete every element this session's journal recorded under
    createdElementIds= that still exists — the idempotent-rebuild wipe.

    Default keep_alignments=True leaves DEFINE_ALIGNMENT_SEGMENT /
    COMMIT_ALIGNMENT / ADOPT_ALIGNMENT_ELEMENT geometry alone so a rebuild
    reuses the same corridor. Pass False only when the engineer asked to
    wipe the corridor too.

    Does NOT fence-delete by proximity (that can catch engineer-drawn
    elements). Safe when nothing has been placed (deleted=0). Resets the
    Python-side plan session flags so place_order_table_stations can run
    again without the re-place gate firing.

    Call this BEFORE re-placing stations/labels/dims/symbols/workspace/
    channelizing when iterating on a plan. place_order_table_stations
    also accepts clear_prior=True to do the same in one step."""
    resp = _ok_or_raise(
        _bridge.call("CLEAR_PLAN_ELEMENTS",
                     keepAlignments="Y" if keep_alignments else "N"),
        "clear_plan_elements")
    # Stations / workspace flags must drop so a rebuild is allowed.
    _PLAN_SESSION["placed_workspace"] = False
    _PLAN_SESSION["stations_placed_aligns"] = set()
    # order_table_built stays True — SharedState still holds the table;
    # rebuild does not need to rebuild the table unless inputs change.
    return resp


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
    createdElementIds. own_element_only defaults True."""
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
