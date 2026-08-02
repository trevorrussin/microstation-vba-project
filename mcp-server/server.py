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

M6 adds the command-registry tool group (list/describe/run_registry_command)
plus move_element / change_element_level / edit_text_element / delete_element.
TEST_REGISTRY_COMMAND is intentionally NOT exposed here — promotion-only,
manual IDE path.

M7 (Stage 4) moved every op's actual implementation into wztc_ops.py, shared
with chat_driver.py's agent loop — everything below is now a one-line
wrapper. Also adds search_reference_manual, over a local FTS5 index
(Data/manual-index.sqlite, built by ingest_manuals.py) of the three
NYSDOT/MUTCD reference PDFs — this one doesn't touch WZTCBridge at all
(pure local SQLite), so it isn't routed through wztc_ops's bridge plumbing.

Run: python server.py  (stdio transport, for `claude mcp add`)
"""
from __future__ import annotations

from typing import Optional

from mcp.server.mcpserver import Image, MCPServer

import manual_search
import wztc_ops
from bridge_client import bridge

wztc_ops.set_bridge(bridge)

mcp = MCPServer("wztc-designer")


# ================================================================ Query

@mcp.tool()
def find_elements_near(x: float, y: float, radius: float, type_filter: str = "") -> list[dict]:
    """Find drawn elements within radius (ft) of (x, y) in the active model.
    type_filter narrows by kind (e.g. 'CELL'); empty string matches all
    types. Returns EVERY candidate with its distance and range, not just
    the nearest — matching is by bounding-box center, so a point near the
    end of a long line matches its midpoint, and multiple close candidates
    are a real ambiguity signal, not noise to collapse to one answer."""
    return wztc_ops.find_elements_near(x, y, radius, type_filter)


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
    return wztc_ops.station_to_point(align_idx, sta)


@mcp.tool()
def get_alignment_stationing(align_idx: int) -> list[dict]:
    """Return the full stationing breakdown for a committed alignment."""
    return wztc_ops.get_alignment_stationing(align_idx)


@mcp.tool()
def list_levels() -> list[dict]:
    """List every level defined in the active design file."""
    return wztc_ops.list_levels()


@mcp.tool()
def classify_site_features(x: float, y: float, radius: float) -> list[dict]:
    """Classify elements near (x, y) by matching level/cell name against
    known WZTC feature names. Site data quality is mixed by design — an
    element that doesn't match a known name/level still comes back
    (kind='unclassified') with its raw geometry rather than being dropped,
    since an unnamed obstruction is still an obstruction the agent must
    reason about."""
    return wztc_ops.classify_site_features(x, y, radius)


# =========================================================== Observation

@mcp.tool()
def capture_view() -> Image:
    """Screenshot the live MicroStation window and return the actual
    image — lets the caller visually verify spacing/layout/sign placement
    instead of only reasoning from coordinates returned by the query
    tools. OS-level capture, not a WZTCBridge op — works regardless of
    what's on top of the MicroStation window, but MicroStation must be
    open. See wztc_ops.capture_view / view_capture.py."""
    result = wztc_ops.capture_view()
    return Image(path=result["path"])


@mcp.tool()
def capture_window(title_substring: str) -> Image:
    """Screenshot any visible top-level window whose title contains
    title_substring -- e.g. "WZTC Agent Chat" for the in-MicroStation chat
    panel, which is a separate OS window from MicroStation's main frame
    (capture_view targets the main frame only). See wztc_ops.capture_window."""
    result = wztc_ops.capture_window(title_substring)
    return Image(path=result["path"])


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
    return wztc_ops.compute_spacing(speed, lane_width, shoulder_width, road_type)


@mcp.tool()
def get_sheet_requirements(sheet_num: str) -> dict:
    """Look up required signs/elements for a 619-series standard sheet
    (e.g. '619-302') from the seeded sheet registry (Data/sheet-registry.tsv
    — currently 6 of 91 sheets). A 'found: false' result for an unseeded
    sheet is the correct, honest answer — fall back to asking the engineer
    for the sign/element list for that sheet rather than guessing one."""
    return wztc_ops.get_sheet_requirements(sheet_num)


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
    return wztc_ops.place_perp_line(align_idx, sta, half_len, reason)


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
    return wztc_ops.place_sign(sign_num, road_type, side, pt1x, pt1y, pt1z, dir1x, dir1y,
                                pt2x, pt2y, pt2z, dir2x, dir2y, reason)


@mcp.tool()
def place_workspace(vertices: list[list[float]], reason: str = "") -> dict:
    """Place the work space boundary shape + hatch. vertices is an ordered
    list of [x, y, z] points describing the boundary — do not repeat the
    first point to close it, that's handled automatically. Interior hatch
    seed point is computed to handle non-convex (e.g. L-shaped) boundaries
    correctly, not a naive centroid."""
    return wztc_ops.place_workspace(vertices, reason)


@mcp.tool()
def place_element_run(element_idx: int, vertices: list[list[float]], reason: str = "") -> dict:
    """Place a channelizing-device / removal-striping / barrier run.
    element_idx: 2=Channelizing Devices, 3=Removal Striping, 4=Temporary
    Barrier, 5=Barrier w/Warning Lights (1=Work Space — use place_workspace
    instead). vertices is an ordered list of [x, y, z] points."""
    return wztc_ops.place_element_run(element_idx, vertices, reason)


@mcp.tool()
def place_cell(cell_name: str, pt_x: float, pt_y: float, pt_z: float = 0, angle_deg: float = 0,
               reason: str = "") -> dict:
    """Place a single cell from the WZTC symbol library at (pt_x, pt_y, pt_z)."""
    return wztc_ops.place_cell(cell_name, pt_x, pt_y, pt_z, angle_deg, reason)


@mcp.tool()
def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Apply standard sign display attributes (level=SF_P, color=240,
    weight=3) to already-placed elements. element_ids are numeric IDs
    returned by place_sign's createdElementIds or by find_elements_near —
    never guessed. Note: fillColor/elementClass=CONSTRUCTION from the
    original CHANGE ATTRIBUTES sequence aren't replicated (no confirmed
    VBA property path) — flagged, not silently dropped."""
    return wztc_ops.set_sign_attributes(element_ids, reason)


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
    return wztc_ops.handoff(kind, from_sta, to_sta, text, notes, reason)


# ================================================================ Session

@mcp.tool()
def undo_last_op() -> dict:
    """Undo the most recent undoable op. Draw ops: deletes exactly the
    elements they created (createdElementIds/elementId). M6 mutations:
    re-applies priorDeltaX/Y (move), priorLevel (level change), or
    priorText (text edit) from the journal. DELETE_ELEMENT is NOT
    undoable — no snapshot to restore; its journal row is skipped.
    Does NOT use MicroStation's own undo stack. Safe to call repeatedly:
    once an op is undone it's marked so a second call won't re-target it."""
    return wztc_ops.undo_last_op()


@mcp.tool()
def get_journal(limit: int = 50) -> list[str]:
    """Return the last `limit` raw journal lines — every op run this
    session, its full parameters (including any reason= passed), and its
    result. This is the PE audit trail: use it to answer "why is that sign
    there" or to review a whole session before handing off a sheet."""
    return wztc_ops.get_journal(limit)


@mcp.tool()
def list_deferred_handoffs() -> list[dict]:
    """List every dimension/callout queued by handoff() this session that
    still needs a few manual clicks through the existing interactive forms."""
    return wztc_ops.list_deferred_handoffs()


# ============================================================ Registry / Edit (M6)

@mcp.tool()
def list_registry_commands(safety_status: str = "") -> list[dict]:
    """List MicroStation command recipes in Data/command-registry.tsv.
    Optional safety_status filter (e.g. 'verified-headless-safe',
    'needs-testing', 'interactive-only-use-handoff'). Only
    verified-headless-safe rows can be executed via run_registry_command;
    interactive-only rows point at handoff() instead."""
    return wztc_ops.list_registry_commands(safety_status)


@mcp.tool()
def describe_registry_command(op_name: str) -> dict:
    """Return the full registry row for one opName — safetyStatus,
    recipeLines, requiredParams, notes, sourceRefs, etc."""
    return wztc_ops.describe_registry_command(op_name)


@mcp.tool()
def run_registry_command(op_name: str, params: Optional[dict] = None, reason: str = "") -> dict:
    """Run a verified-headless-safe keyin_recipe from the command
    registry. Refuses needs-testing / interactive-only / unsafe-blocked
    rows with a clear ERROR (not a silent no-op). Pass recipe params in
    `params` matching requiredParams (e.g. {"level": "Default"} for
    ACTIVE_LEVEL, {"color": "0"} for ACTIVE_COLOR). direct_api rows
    (MOVE_ELEMENT etc.) must use their dedicated tools below — this
    entry point will refuse them."""
    return wztc_ops.run_registry_command(op_name, params, reason)


@mcp.tool()
def move_element(element_id: str, delta_x: float, delta_y: float, delta_z: float = 0,
                 own_element_only: bool = True, reason: str = "") -> dict:
    """Move an element by delta_x/delta_y(/delta_z) design units (ft).
    By default own_element_only=True: element_id must appear in this
    session's journal as something the agent created. Response includes
    priorDeltaX/Y for undo_last_op."""
    return wztc_ops.move_element(element_id, delta_x, delta_y, delta_z, own_element_only, reason)


@mcp.tool()
def change_element_level(element_id: str, level: str, own_element_only: bool = True,
                         reason: str = "") -> dict:
    """Change an element's level by ID. own_element_only defaults True
    (journal-gated). Response includes priorLevel for undo_last_op."""
    return wztc_ops.change_element_level(element_id, level, own_element_only, reason)


@mcp.tool()
def edit_text_element(element_id: str, new_text: str, own_element_only: bool = True,
                      reason: str = "") -> dict:
    """Replace text on a TextElement / TextNodeElement by ID.
    own_element_only defaults True. Response includes priorText for
    undo_last_op."""
    return wztc_ops.edit_text_element(element_id, new_text, own_element_only, reason)


@mcp.tool()
def delete_element(element_id: str, own_element_only: bool = True, reason: str = "") -> dict:
    """Delete an element by ID. own_element_only defaults True.
    NOT undoable via undo_last_op (no snapshot to restore) — the
    response says so plainly (notUndoable=Y). Prefer undo_last_op on
    the placing op when you still can."""
    return wztc_ops.delete_element(element_id, own_element_only, reason)


# ==================================================== Reference Manuals (M7)

@mcp.tool()
def search_reference_manual(query: str, source: str = "", max_results: int = 10) -> list[dict]:
    """Full-text search over the three NYSDOT/MUTCD reference manuals —
    MUTCD Part 6 (Temporary Traffic Control), the NYS MUTCD Supplement, and
    the NYSDOT standard detail sheets — indexed from Project Documentation/
    via ingest_manuals.py. Ground engineer-facing answers about MUTCD/NYSDOT
    requirements in these excerpts rather than recollection; include the
    page citation (page_start/page_end) in the answer so the engineer can
    verify against the actual manual. source optionally narrows to one of
    'part6' | 'supplement' | 'stdsht' (empty searches all three). An empty
    result means either no match, or the index hasn't been built yet (run
    ingest_manuals.py) — not necessarily "nothing exists on this topic"."""
    return manual_search.search(query, source=source, max_results=max_results)


if __name__ == "__main__":
    mcp.run()
