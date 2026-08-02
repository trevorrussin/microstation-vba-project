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

import view_capture

_bridge = None


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
    types. Returns EVERY candidate with its distance and range, not just
    the nearest — matching is by bounding-box center, so a point near the
    end of a long line matches its midpoint, and multiple close candidates
    are a real ambiguity signal, not noise to collapse to one answer."""
    resp = _ok_or_raise(
        _bridge.call("FIND_ELEMENTS_NEAR", x=x, y=y, radius=radius, typeFilter=type_filter),
        "find_elements_near")
    return resp.get("rows", [])


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


def list_levels() -> list[dict]:
    """List every level defined in the active design file."""
    resp = _ok_or_raise(_bridge.call("LIST_LEVELS"), "list_levels")
    return resp.get("rows", [])


def describe_drawing_state() -> dict:
    """Inspect the active model before making any edits: 2D/3D, master/sub
    units and resolution, annotation scale (signs/cells are auto-multiplied
    by this — see the 2026-08-02 sign-scale fix), active level/color/line
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
    reason about."""
    resp = _ok_or_raise(_bridge.call("CLASSIFY_SITE_FEATURES", x=x, y=y, radius=radius), "classify_site_features")
    return resp.get("rows", [])


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

def place_perp_line(align_idx: int, sta: float, half_len: float = 40, reason: str = "") -> dict:
    """Place a perpendicular reference tick line (2*half_len ft long,
    default 80ft) at a station along a committed alignment."""
    return _ok_or_raise(
        _bridge.call("PLACE_PERP_LINE", alignIdx=align_idx, sta=sta, halfLen=half_len, reason=reason),
        "place_perp_line")


def place_sign(sign_num: str, road_type: str, side: str,
               pt1x: float, pt1y: float, pt1z: float, dir1x: float, dir1y: float,
               pt2x: Optional[float] = None, pt2y: Optional[float] = None, pt2z: Optional[float] = None,
               dir2x: Optional[float] = None, dir2y: Optional[float] = None,
               reason: str = "") -> dict:
    """Place a sign face + post + text label at a resolved point/direction —
    typically from station_to_point, offset along the perpendicular to dodge
    an obstruction found via find_elements_near/classify_site_features.
    sign_num MUST be a SignLibrary.bas key (e.g. 'W20-01RA'), not a raw
    sheet code — run it through resolve_sign_code first if it came from
    get_sheet_requirements.
    side is 'One Side' or 'Both Sides'; pt2/dir2/pt2z are required only for
    'Both Sides' (a connecting arc is drawn between the two). This tool only
    executes — it never decides where the sign belongs; resolve the point
    first, then call this. Pass reason whenever the point was adjusted from
    the default (e.g. "shifted 4 ft off perp — utility pole at 3+20")."""
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
    """Place the work space boundary shape + associative hatch.
    vertices is an ordered list of [x, y, z] points — do not repeat the
    first point to close it. Hatch uses CreateHatchPattern1 + SetPattern
    (Element API), not CadInputQueue HATCH ICON."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(_bridge.call("PLACE_WORKSPACE", verticesTSV=verts_tsv, reason=reason), "place_workspace")


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
                             line_style_index: int | None = None, own_element_only: bool = True,
                             reason: str = "") -> dict:
    """Set element color and/or line weight (and optional linestyle index)."""
    params = {"elementId": element_id, "ownElementOnly": ("Y" if own_element_only else "N"), "reason": reason}
    if color is not None:
        params["color"] = color
    if weight is not None:
        params["weight"] = weight
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
    """Place a single cell from the WZTC symbol library at (pt_x, pt_y, pt_z)."""
    return _ok_or_raise(
        _bridge.call("PLACE_CELL", cellName=cell_name, ptX=pt_x, ptY=pt_y, ptZ=pt_z,
                     angleDeg=angle_deg, reason=reason),
        "place_cell")


def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Apply standard sign display attributes (level=SF_P, color=240,
    weight=3) to already-placed elements. element_ids are numeric IDs
    returned by place_sign's createdElementIds or by find_elements_near —
    never guessed. Note: fillColor/elementClass=CONSTRUCTION from the
    original CHANGE ATTRIBUTES sequence aren't replicated (no confirmed
    VBA property path) — flagged, not silently dropped."""
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


def get_journal(limit: int = 50) -> list[str]:
    """Return the last `limit` raw journal lines — every op run this
    session, its full parameters (including any reason= passed), and its
    result. This is the PE audit trail: use it to answer "why is that sign
    there" or to review a whole session before handing off a sheet."""
    resp = _ok_or_raise(_bridge.call("GET_JOURNAL", limit=limit), "get_journal")
    return [row.get("line", "") for row in resp.get("rows", [])]


def list_deferred_handoffs() -> list[dict]:
    """List every dimension/callout queued by handoff() this session that
    still needs a few manual clicks through the existing interactive forms."""
    resp = _ok_or_raise(_bridge.call("LIST_DEFERRED_HANDOFFS"), "list_deferred_handoffs")
    return resp.get("rows", [])


# ============================================================ Registry / Edit (M6)

def list_registry_commands(safety_status: str = "") -> list[dict]:
    """List MicroStation command recipes in Data/command-registry.tsv.
    Optional safety_status filter (e.g. 'verified-headless-safe',
    'needs-testing', 'interactive-only-use-handoff'). Only
    verified-headless-safe rows can be executed via run_registry_command;
    interactive-only rows point at handoff() instead."""
    params = {}
    if safety_status:
        params["safetyStatus"] = safety_status
    resp = _ok_or_raise(_bridge.call("LIST_REGISTRY_COMMANDS", **params), "list_registry_commands")
    return resp.get("rows", [])


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
