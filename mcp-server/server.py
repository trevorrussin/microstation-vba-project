"""
WZTC Designer agent MCP server (plan Layer 5).

Exposes the Query/Compute/Draw/Session tool groups from
parallel-zooming-star.md as real MCP tools over WZTCBridge.bas, so an
agent can query site conditions, get a deterministic spacing/sign number,
draw an element, and record why — without the engineer clicking through
the 8-step wizard.

Engineering-judgment boundary (CLAUDE.md, and the plan's core design
rule): this server never computes a spacing, taper length, or sign size
itself. compute_spacing / get_sheet_requirements wrap WZTCRules.bas /
WZTCSheetRegistry.bas so those numbers stay deterministic and PE-auditable.
The calling agent decides *what* to place and *how to respond to a site
condition* (an obstruction, a driveway); it must never invent a number
that belongs in one of those two tools.

Scope note: the Edit vocabulary (move_element, change_level, edit_text) is
M6 in the plan and is NOT exposed here yet — WZTCBridge.bas has no dispatch
case for it. delete_element also isn't exposed as a standalone tool for the
same reason; undo_last_op uses the equivalent VBA primitive internally.

Run: python server.py  (stdio transport, for `claude mcp add`)
"""
from __future__ import annotations

from typing import Optional

from mcp.server.mcpserver import MCPServer

from bridge_client import bridge

mcp = MCPServer("wztc-designer")


def _ok_or_raise(resp: dict, context: str) -> dict:
    if resp["status"] == "ERROR":
        raise RuntimeError(f"{context} failed: {resp.get('note', resp)}")
    return resp


# ================================================================ Query

@mcp.tool()
def find_elements_near(x: float, y: float, radius: float, type_filter: str = "") -> list[dict]:
    """Find drawn elements within radius (ft) of (x, y) in the active model.
    type_filter narrows by kind (e.g. 'CELL'); empty string matches all
    types. Returns EVERY candidate with its distance and range, not just
    the nearest — matching is by bounding-box center, so a point near the
    end of a long line matches its midpoint, and multiple close candidates
    are a real ambiguity signal, not noise to collapse to one answer."""
    resp = _ok_or_raise(
        bridge.call("FIND_ELEMENTS_NEAR", x=x, y=y, radius=radius, typeFilter=type_filter),
        "find_elements_near")
    return resp.get("rows", [])


@mcp.tool()
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
    return _ok_or_raise(bridge.call("STATION_TO_POINT", alignIdx=align_idx, sta=sta), "station_to_point")


@mcp.tool()
def get_alignment_stationing(align_idx: int) -> list[dict]:
    """Return the full stationing breakdown for a committed alignment."""
    resp = _ok_or_raise(bridge.call("GET_ALIGNMENT_STATIONING", alignIdx=align_idx), "get_alignment_stationing")
    return resp.get("rows", [])


@mcp.tool()
def list_levels() -> list[dict]:
    """List every level defined in the active design file."""
    resp = _ok_or_raise(bridge.call("LIST_LEVELS"), "list_levels")
    return resp.get("rows", [])


@mcp.tool()
def classify_site_features(x: float, y: float, radius: float) -> list[dict]:
    """Classify elements near (x, y) by matching level/cell name against
    known WZTC feature names. Site data quality is mixed by design — an
    element that doesn't match a known name/level still comes back
    (kind='unclassified') with its raw geometry rather than being dropped,
    since an unnamed obstruction is still an obstruction the agent must
    reason about."""
    resp = _ok_or_raise(bridge.call("CLASSIFY_SITE_FEATURES", x=x, y=y, radius=radius), "classify_site_features")
    return resp.get("rows", [])


# ============================================================== Compute

@mcp.tool()
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
        bridge.call("COMPUTE_SPACING", speed=speed, laneWidth=lane_width,
                    shoulderWidth=shoulder_width, roadType=road_type),
        "compute_spacing")


@mcp.tool()
def get_sheet_requirements(sheet_num: str) -> dict:
    """Look up required signs/elements for a 619-series standard sheet
    (e.g. '619-302') from the seeded sheet registry (Data/sheet-registry.tsv
    — currently 6 of 91 sheets). A 'found: false' result for an unseeded
    sheet is the correct, honest answer — fall back to asking the engineer
    for the sign/element list for that sheet rather than guessing one."""
    resp = bridge.call("GET_SHEET_REQUIREMENTS", sheetNum=sheet_num)
    if resp["status"] == "ERROR":
        return {"found": False, "note": resp.get("note", "")}
    resp["found"] = True
    return resp


# ================================================================== Draw
# Every draw tool takes an optional `reason`. It rides through untouched to
# WZTCBridge's journal (Bridge/wztc-journal.tsv) alongside every other
# param — pass it whenever a placement isn't the default/expected one (an
# obstruction dodge, a non-standard station) so a PE reviewing the journal
# later can see *why*, not just *what*.

@mcp.tool()
def place_perp_line(align_idx: int, sta: float, half_len: float = 40, reason: str = "") -> dict:
    """Place a perpendicular reference tick line (2*half_len ft long,
    default 80ft) at a station along a committed alignment."""
    return _ok_or_raise(
        bridge.call("PLACE_PERP_LINE", alignIdx=align_idx, sta=sta, halfLen=half_len, reason=reason),
        "place_perp_line")


@mcp.tool()
def place_sign(sign_num: str, road_type: str, side: str,
               pt1x: float, pt1y: float, pt1z: float, dir1x: float, dir1y: float,
               pt2x: Optional[float] = None, pt2y: Optional[float] = None, pt2z: Optional[float] = None,
               dir2x: Optional[float] = None, dir2y: Optional[float] = None,
               reason: str = "") -> dict:
    """Place a sign face + post + text label at a resolved point/direction —
    typically from station_to_point, offset along the perpendicular to dodge
    an obstruction found via find_elements_near/classify_site_features.
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
        bridge.call("PLACE_SIGN", signNum=sign_num, roadType=road_type, side=side,
                    pt1X=pt1x, pt1Y=pt1y, pt1Z=pt1z, dir1X=dir1x, dir1Y=dir1y,
                    pt2X=pt2x, pt2Y=pt2y, pt2Z=pt2z, dir2X=dir2x, dir2Y=dir2y,
                    reason=reason),
        "place_sign")


@mcp.tool()
def place_workspace(vertices: list[list[float]], reason: str = "") -> dict:
    """Place the work space boundary shape + hatch. vertices is an ordered
    list of [x, y, z] points describing the boundary — do not repeat the
    first point to close it, that's handled automatically. Interior hatch
    seed point is computed to handle non-convex (e.g. L-shaped) boundaries
    correctly, not a naive centroid."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(bridge.call("PLACE_WORKSPACE", verticesTSV=verts_tsv, reason=reason), "place_workspace")


@mcp.tool()
def place_element_run(element_idx: int, vertices: list[list[float]], reason: str = "") -> dict:
    """Place a channelizing-device / removal-striping / barrier run.
    element_idx: 2=Channelizing Devices, 3=Removal Striping, 4=Temporary
    Barrier, 5=Barrier w/Warning Lights (1=Work Space — use place_workspace
    instead). vertices is an ordered list of [x, y, z] points."""
    verts_tsv = "|".join(f"{p[0]},{p[1]},{p[2] if len(p) > 2 else 0}" for p in vertices)
    return _ok_or_raise(
        bridge.call("PLACE_ELEMENT_RUN", elementIdx=element_idx, verticesTSV=verts_tsv, reason=reason),
        "place_element_run")


@mcp.tool()
def place_cell(cell_name: str, pt_x: float, pt_y: float, pt_z: float = 0, angle_deg: float = 0,
               reason: str = "") -> dict:
    """Place a single cell from the WZTC symbol library at (pt_x, pt_y, pt_z)."""
    return _ok_or_raise(
        bridge.call("PLACE_CELL", cellName=cell_name, ptX=pt_x, ptY=pt_y, ptZ=pt_z,
                    angleDeg=angle_deg, reason=reason),
        "place_cell")


@mcp.tool()
def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Apply standard sign display attributes (level=SF_P, color=240,
    weight=3) to already-placed elements. element_ids are numeric IDs
    returned by place_sign's createdElementIds or by find_elements_near —
    never guessed. Note: fillColor/elementClass=CONSTRUCTION from the
    original CHANGE ATTRIBUTES sequence aren't replicated (no confirmed
    VBA property path) — flagged, not silently dropped."""
    return _ok_or_raise(
        bridge.call("SET_SIGN_ATTRIBUTES", elementIds=",".join(element_ids), reason=reason),
        "set_sign_attributes")


@mcp.tool()
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
    return _ok_or_raise(bridge.call("HANDOFF", **params), "handoff")


# ================================================================ Session

@mcp.tool()
def undo_last_op() -> dict:
    """Undo the most recent draw op by deleting exactly the elements it
    created (tracked in the journal via createdElementIds/elementId), NOT
    MicroStation's own undo stack — that stack's grouping behavior across a
    multi-element op (e.g. place_sign's post+face+text+arc) has not been
    verified in the IDE, so this sidesteps it with a deterministic,
    independently-testable delete instead. Safe to call repeatedly: once an
    op is undone it's marked so a second call won't re-target it."""
    return _ok_or_raise(bridge.call("UNDO_LAST_OP"), "undo_last_op")


@mcp.tool()
def get_journal(limit: int = 50) -> list[str]:
    """Return the last `limit` raw journal lines — every op run this
    session, its full parameters (including any reason= passed), and its
    result. This is the PE audit trail: use it to answer "why is that sign
    there" or to review a whole session before handing off a sheet."""
    resp = _ok_or_raise(bridge.call("GET_JOURNAL", limit=limit), "get_journal")
    return [row.get("line", "") for row in resp.get("rows", [])]


@mcp.tool()
def list_deferred_handoffs() -> list[dict]:
    """List every dimension/callout queued by handoff() this session that
    still needs a few manual clicks through the existing interactive forms."""
    resp = _ok_or_raise(bridge.call("LIST_DEFERRED_HANDOFFS"), "list_deferred_handoffs")
    return resp.get("rows", [])


if __name__ == "__main__":
    mcp.run()
