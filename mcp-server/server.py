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
def find_elements_near(x: float, y: float, radius: float, type_filter: str = "",
                       force: bool = False) -> list[dict]:
    """Find drawn elements within radius (ft) of (x, y) in the active model.
    type_filter narrows by kind (e.g. 'CELL'); empty string matches all
    types. Mid sheet-plan: wide radius / repeated calls are refused unless
    force=True — prefer view_drawing then FINAL after place_sheet_geometry."""
    return wztc_ops.find_elements_near(x, y, radius, type_filter, force=force)


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
def get_alignment_vertices(align_idx: int) -> list[dict]:
    """Return a committed alignment's raw path segments (straight or arc)
    in master units — fetch once, then interpolate station->XY locally
    instead of one bridge round trip per point."""
    return wztc_ops.get_alignment_vertices(align_idx)


@mcp.tool()
def get_locked_designer_inputs() -> dict:
    """Return the designer inputs (speed/road_type/lane_width/
    shoulder_width/area_type/sheet_num/...) locked in by the most recent
    successful build_wztc_order_table call this session, or
    {"locked": False} if none yet. Call this instead of re-deriving or
    re-asking the engineer for values already established earlier in the
    same build — including after a turn that hit MAX_TOOL_ITERATIONS and
    had to continue in a fresh turn."""
    return wztc_ops.get_locked_designer_inputs()


@mcp.tool()
def get_plan_status() -> dict:
    """Named 619 sheet-build checklist only. Outside a sheet plan returns
    sheetPlanActive=False — general CAD is not gated. When active, includes
    persistedPath/updatedAt from Bridge/sheet-plan.json."""
    return wztc_ops.get_plan_status()


@mcp.tool()
def get_placements(sheet_num: str = "", kind: str = "", zone: str = "",
                   run: str = "", align_idx: int = 0) -> list[dict]:
    """List agent-placed primitives from the placement registry (kind/zone/run).
    Prefer this over fishing get_journal for delete/edit by feature."""
    return wztc_ops.get_placements(
        sheet_num=sheet_num, kind=kind, zone=zone, run=run, align_idx=align_idx)


@mcp.tool()
def delete_placements(kind: str = "", zone: str = "", run: str = "",
                      align_idx: int = 0, sheet_num: str = "",
                      reason: str = "") -> dict:
    """Delete DGN elements for matching placement-registry records
    (e.g. kind='cone', run='laneTaperRun')."""
    return wztc_ops.delete_placements(
        kind=kind, zone=zone, run=run, align_idx=align_idx,
        sheet_num=sheet_num, reason=reason)


@mcp.tool()
def get_geometry_scorecard(sheet_num: str = "") -> dict:
    """Post-placement scorecard: compile expectations vs placement registry
    plus live Tier-1 stacked-duplicate hash against the model."""
    return wztc_ops.get_geometry_scorecard(sheet_num=sheet_num)


@mcp.tool()
def check_build_overlap(sheet_num: str = "", origin: list | None = None,
                        path_vertices: list | None = None,
                        lateral_half_width: float = 0.0,
                        sta0: float = 0.0, sta1: float = 0.0,
                        scan_model: bool = True) -> dict:
    """Caution-not-block overlap check (ledger + Tier 1 stacks + Tier 2
    station/offset). Do not compose find_elements_near. blocking=False."""
    return wztc_ops.check_build_overlap(
        sheet_num=sheet_num, origin=origin, path_vertices=path_vertices,
        lateral_half_width=lateral_half_width, sta0=sta0, sta1=sta1,
        scan_model=scan_model)


@mcp.tool()
def get_elements_in_range_box(low_x: float, low_y: float,
                              high_x: float, high_y: float,
                              max_rows: int = 1500) -> dict:
    """Elements whose Range intersects a world AABB (not center-in-box)."""
    return wztc_ops.get_elements_in_range_box(low_x, low_y, high_x, high_y, max_rows)


@mcp.tool()
def reflect_sheet_build(max_iterations: int = 1) -> dict:
    """Deterministic reflection for the active sheet build — cites registry
    primitiveIds / reqIds and scorecard failures. Call before FINAL."""
    return wztc_ops.reflect_sheet_build(max_iterations=max_iterations)


@mcp.tool()
def list_levels(name_contains: str = "") -> list[dict]:
    """List levels matching name_contains (required substring, e.g. 'TWZ').
    Refuses unfiltered listings — DGNs can have thousands of levels."""
    return wztc_ops.list_levels(name_contains)


@mcp.tool()
def list_colors() -> list[dict]:
    """Return every color-table index + RGB for the active DGN."""
    return wztc_ops.list_colors()


@mcp.tool()
def resolve_color(name: str = "", red: int | None = None,
                  green: int | None = None, blue: int | None = None) -> dict:
    """Map a named color or RGB triple to the closest index in this DGN's
    color table. Call before change_element_symbology when the engineer
    names a color — never guess an index."""
    return wztc_ops.resolve_color(name, red, green, blue)


@mcp.tool()
def list_line_styles(name_contains: str = "") -> list[dict]:
    """List line styles matching name_contains (required). Prefer
    resolve_line_style when you know the name."""
    return wztc_ops.list_line_styles(name_contains)


@mcp.tool()
def resolve_line_style(name: str = "") -> dict:
    """Map a line-style name/alias to the exact Name key for this DGN.
    Pass returned name to change_element_symbology(line_style_name=...)."""
    return wztc_ops.resolve_line_style(name)


@mcp.tool()
def cell_library_status() -> dict:
    """Whether a cell library is attached and its path."""
    return wztc_ops.cell_library_status()


@mcp.tool()
def attach_cell_library(lib_path: str = "") -> dict:
    """Attach a .cel library (empty path = default WZTC ny_plan_wztc.cel)."""
    return wztc_ops.attach_cell_library(lib_path)


@mcp.tool()
def list_cells(name_contains: str = "", include_shared: bool = False) -> list[dict]:
    """List cells in the attached library (filter by name/description)."""
    return wztc_ops.list_cells(name_contains, include_shared)


@mcp.tool()
def list_cell_libraries(name_contains: str = "", lib_dir: str = "") -> dict:
    """List .cel libraries in the NY plan cell folder (utility, striping,
    roadway, wztc, …). Filter with name_contains e.g. 'utility'."""
    return wztc_ops.list_cell_libraries(name_contains, lib_dir)


@mcp.tool()
def find_cell(query: str, lib_dir: str = "", library_path: str = "",
              max_results: int = 25) -> dict:
    """Search cell name+description across NY plan .cel libraries for a
    plain-language query (e.g. 'gas meter', 'ARROW LEFT'). Returns
    cellName + libraryPath for place_cell(..., library_path=...)."""
    return wztc_ops.find_cell(query, lib_dir, library_path, max_results)


@mcp.tool()
def list_fonts(name_contains: str = "") -> list[dict]:
    """List fonts in the active DGN. Optional name_contains filter."""
    return wztc_ops.list_fonts(name_contains)


@mcp.tool()
def resolve_font(name: str = "") -> dict:
    """Map a font name to Name + ID for this DGN."""
    return wztc_ops.resolve_font(name)


@mcp.tool()
def list_text_styles(name_contains: str = "") -> list[dict]:
    """List text styles (name, height, width, font). Optional filter."""
    return wztc_ops.list_text_styles(name_contains)


@mcp.tool()
def resolve_text_style(name: str = "") -> dict:
    """Map a text-style name to height/width/font for this DGN."""
    return wztc_ops.resolve_text_style(name)


@mcp.tool()
def describe_drawing_state() -> dict:
    """Inspect the active model before making any edits: 2D/3D, units and
    resolution, annotation scale, active level/symbology, active ACS, open
    views, reference attachments, and current selection. Call at the start
    of a session and whenever unsure what drawing/scale you're working in."""
    return wztc_ops.describe_drawing_state()


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
    image — MCP / Claude Code clients only.

    The in-MicroStation chat_driver does NOT expose this tool; that agent
    uses view_drawing (ad-hoc) or run_visual_qa_captures / run_sheet_build
    (scripted frames attached as vision + panel SCREENSHOT). OS-level
    capture, not a WZTCBridge op — MicroStation must be open. See
    wztc_ops.capture_view / view_capture.py."""
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


@mcp.tool()
def adjust_view(zoom_out_percent: float = 0, pan_x: float = 0, pan_y: float = 0,
                 view_num: int = 1,
                 center_x: float | None = None, center_y: float | None = None,
                 width: float | None = None, height: float | None = None,
                 force: bool = False) -> dict:
    """Zoom/pan the MicroStation view via COM. After place_sheet_geometry
    during a sheet build, prefer run_visual_qa_captures (force=True to
    override). General CAD is unaffected."""
    return wztc_ops.adjust_view(zoom_out_percent=zoom_out_percent, pan_x=pan_x, pan_y=pan_y,
                                 view_num=view_num, center_x=center_x, center_y=center_y,
                                 width=width, height=height, force=force)


@mcp.tool()
def get_elements_range(element_ids: list[str] | str) -> dict:
    """Return the combined bbox of one or more element IDs. Use this
    whenever you already have elementId(s) instead of find_elements_near."""
    return wztc_ops.get_elements_range(element_ids)


@mcp.tool()
def get_element_vertices(element_id: str) -> dict:
    """Densified vertices for a picked line / line-string / arc / complex
    chain. Use after an element pick. Not a bounding box."""
    return wztc_ops.get_element_vertices(element_id)


@mcp.tool()
def propose_corridor_source() -> dict:
    """Phase B: ask which roadway to build along (last placed / click /
    level / points). Call after designer inputs."""
    return wztc_ops.propose_corridor_source()


@mcp.tool()
def lock_corridor_path(source: str, element_id: str = "",
                       vertices: list | None = None, reverse: bool = False,
                       edge_role: str = "first_travel_outer",
                       level_name_contains: str = "") -> dict:
    """Lock first-travel-outer path_vertices from the Phase B answer.
    source: last_placed | element | level | points. edge_role centerline
    offsets to the outer edge."""
    return wztc_ops.lock_corridor_path(
        source, element_id=element_id, vertices=vertices, reverse=reverse,
        edge_role=edge_role, level_name_contains=level_name_contains)


@mcp.tool()
def propose_work_area_on_path() -> dict:
    """Phase C: how to place the work bay along the locked road."""
    return wztc_ops.propose_work_area_on_path()


@mcp.tool()
def snap_work_area_to_path(mode: str, p1: list | None = None,
                           p2: list | None = None,
                           start_sta: float | None = None,
                           length_ft: float | None = None,
                           mid: list | None = None) -> dict:
    """Snap work-bay ends onto the locked corridor. mode: ends |
    station_length | mid_length. Returns upstream_edge / downstream_edge
    for resolve_sheet_lateral."""
    return wztc_ops.snap_work_area_to_path(
        mode, p1=p1, p2=p2, start_sta=start_sta, length_ft=length_ft, mid=mid)


@mcp.tool()
def focus_view_on_elements(element_ids: list[str] | str, margin: float = 1.3,
                            view_num: int = 1, min_width: float = 50.0,
                            min_height: float = 50.0) -> dict:
    """Frame the view on the bbox of the given element ID(s). Degenerate
    (zero-area) ranges -- e.g. a horizontal line -- get at least
    min_width x min_height so the view is still usable."""
    return wztc_ops.focus_view_on_elements(element_ids, margin=margin, view_num=view_num,
                                            min_width=min_width, min_height=min_height)


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
    (e.g. '619-302') from Data/sheet-registry.tsv (all 91 DesignerRef
    sheets; some stubs have empty signs when not in the 2026 Book 3 PDF).
    Check notes for stub/catalog rows. A 'found: false' result means the
    sheet number is unknown to the registry — ask the engineer rather
    than guessing. Sign codes in the `signs` field are as printed on the
    sheet (e.g. 'W20-1') — pass each through resolve_sign_code before
    calling place_sign, don't assume it's already a valid library key.
    When Data/sheet-specs/<sheet>.build.md exists, response includes
    buildGuidePath + buildGuide (durable tips) — follow those on builds."""
    return wztc_ops.get_sheet_requirements(sheet_num)


@mcp.tool()
def get_required_designer_inputs(sheet_num: str = "") -> dict:
    """Table-driven ask list from Data/sheet-specs/<sheet>.json inputs[].
    Call before ask_user_choice on a named 619 sheet. Use the returned
    options; do not invent speed/area_type; do not offer out-of-domain
    values (619-311 has no 60 mph). Skip locked; apply derived and cite."""
    return wztc_ops.get_required_designer_inputs(sheet_num)


@mcp.tool()
def get_sheet_build_guide(sheet_num: str) -> dict:
    """Load the durable live-build playbook for a named 619 sheet
    (Data/sheet-specs/<sheet>.build.md). Machine prefs stay in the JSON;
    this markdown holds tips, QA checklist, and gotchas. Call when
    get_sheet_requirements attached a buildGuide, or mid-build via
    get_plan_status.buildGuidePath."""
    return wztc_ops.get_sheet_build_guide(sheet_num)


@mcp.tool()
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
    return wztc_ops.resolve_sign_code(code)


# ================================================================== Draw
# Every draw tool takes an optional `reason`. It rides through untouched to
# WZTCBridge's journal (Bridge/wztc-journal.tsv) alongside every other
# param — pass it whenever a placement isn't the default/expected one (an
# obstruction dodge, a non-standard station) so a PE reviewing the journal
# later can see *why*, not just *what*.

@mcp.tool()
def place_perp_line(align_idx: int, sta: float, half_len: float = 40,
                    reason: str = "", one_off: bool = False) -> dict:
    """Place a SINGLE perpendicular reference tick line (2*half_len ft
    long, default 80ft) at a station along a committed alignment. For a
    full-plan run, prefer place_order_table_stations instead — it places
    every order-table item's tick line in ONE call. Use this one only
    for a genuinely one-off tick line outside the order-table flow, and
    pass one_off=True — without that flag the tool refuses when the
    session already looks like a plan (workspace / order table)."""
    return wztc_ops.place_perp_line(align_idx, sta, half_len, reason, one_off)


@mcp.tool()
def place_sign(sign_num: str, road_type: str, side: str,
               pt1x: float, pt1y: float, pt1z: float, dir1x: float, dir1y: float,
               pt2x: Optional[float] = None, pt2y: Optional[float] = None, pt2z: Optional[float] = None,
               dir2x: Optional[float] = None, dir2y: Optional[float] = None,
               reason: str = "", align_idx: int = 0, one_off: bool = False,
               post_angle_deg: Optional[float] = None) -> dict:
    """Place a sign assembly (post + edge-connected stem + face + label).

    pt1 is the ATTACHMENT on the perp tick — typically the OUTWARD TIP of
    the 80ft tick (station + outward_perp * half_len), NOT the alignment
    station and NOT the face center. dir1 is that same unit outward perp
    — never the alignment tangent (live miss: assembly built along the
    road). From an order-table row: outward = rotate tan 90deg toward the
    side; tip = (ptX,ptY)+outward*half_len; place_sign(..., pt1=tip,
    dir1=outward). Stem/post/face edge geometry is handled in VBA.
    sign_num MUST be a SignLibrary.bas key — resolve_sign_code first if it
    came from get_sheet_requirements. side is 'One Side' or 'Both Sides';
    pt2/dir2 required only for Both Sides. Pass reason when the tip was
    adjusted from the default (e.g. obstruction dodge). Pass align_idx
    (1=Upstream, 2=Downstream) so scoped clear_prior can wipe that sign.
    If build_wztc_order_table already ran, sign_num must match one of its
    resolved sign_rows — this refuses a hand-picked/guessed legend variant
    that bypasses the order table's own Table-driven resolution. Pass
    one_off=True only for a genuine ad-hoc sign outside the order table.
    post_angle_deg rotates the TWZSGN_P post with travel tangent; omit to
    keep the post at view angle. Faces stay view-horizontal either way."""
    return wztc_ops.place_sign(sign_num, road_type, side, pt1x, pt1y, pt1z, dir1x, dir1y,
                                pt2x, pt2y, pt2z, dir2x, dir2y, reason, align_idx, one_off,
                                post_angle_deg)


@mcp.tool()
def place_workspace(vertices: list[list[float]], reason: str = "") -> dict:
    """Place the work space boundary (unfilled) + hatch stripes.
    Verify returned elementId with find_elements_near before continuing."""
    return wztc_ops.place_workspace(vertices, reason)


@mcp.tool()
def build_wztc_order_table(speed: int, road_type: str, lane_width: int, shoulder_width: str,
                            sign_rows: list[dict] | None = None,
                            category: str = "", sheet_num: str = "",
                            area_type: str = "", closure_type: str = "",
                            exposure_condition: str = "",
                            protective_vehicle_gvw: int = 0) -> dict:
    """Build the full per-alignment order table (headless Submit & Draw).

    Pass sheet_num when a Data/sheet-specs/<sheet>.json exists — the sheet
    drives stations, spacings, and SignLibrary keys (sign_rows optional).
    Pass area_type ("URBAN"/"RURAL"/"FREEWAY") only when that sheet's
    tableRoles include advanceWarningSpacing; omit it for sheets like
    619-301. Pass protective_vehicle_gvw (lbs) when roll-ahead is GVW-keyed
    (e.g. 619-301); default 22000 is used if omitted.

    Without a spec it falls back to generic defaults (specDriven=False).

    sign_rows: [{"align_idx": 1|2, "sign_num": SignLibrary key, "side":
    "One Side"|"Both Sides", "spacing_ft": optional, "size": optional}]."""
    return wztc_ops.build_wztc_order_table(speed, road_type, lane_width, shoulder_width,
                                            sign_rows, category, sheet_num,
                                            area_type, closure_type, exposure_condition,
                                            protective_vehicle_gvw)


@mcp.tool()
def find_reference_linework(level_name_contains: str, include_references: bool = False,
                            ref_name_contains: str = "", force: bool = False) -> list[dict]:
    """Locate connected line/line-string chains on a level. After an order
    table is built, vague Default/RDEFAULT fishing is refused — prefer
    assemble_corridor. force=True only if the engineer named that level."""
    return wztc_ops.find_reference_linework(
        level_name_contains, include_references, ref_name_contains, force=force)


@mcp.tool()
def define_alignment_segment(align_idx: int, vertices: list[list[float]],
                             reason: str = "", force: bool = False) -> dict:
    """Create straight alignment line segments from vertices and record
    them as a drawing session for align_idx (1=Upstream, 2=Downstream).
    Mid sheet-plan prefer assemble_corridor; freestyle define is refused
    unless force=True."""
    return wztc_ops.define_alignment_segment(align_idx, vertices, reason, force=force)


@mcp.tool()
def commit_alignment(align_idx: int, force: bool = False) -> dict:
    """Group every segment recorded by define_alignment_segment for
    align_idx into a graphic group, ready for place_order_table_stations.
    Once both align 1 and align 2 are ready (and a sheet spec is locked
    for this build), runs the corridor-topology check immediately and
    raises if it fails. force=True to proceed anyway."""
    return wztc_ops.commit_alignment(align_idx, force)


@mcp.tool()
def adopt_alignment(align_idx: int, element_id: str, force: bool = False) -> dict:
    """Re-bind SharedState for align_idx to an EXISTING LINE element —
    no redraw. Use after VBA hot-reload / IDE Reset wiped session state,
    or when the engineer picks an existing centerline. align_idx:
    1=Upstream, 2=Downstream. element_id must be a LINE. Once both
    alignments are ready (and a sheet spec is locked), runs the
    corridor-topology check immediately and raises if it fails. force=True
    to proceed anyway."""
    return wztc_ops.adopt_alignment(align_idx, element_id, force)


@mcp.tool()
def assemble_corridor(upstream_edge: list[float], downstream_edge: list[float],
                      approach_length_ft: float = 0.0,
                      force: bool = False,
                      path_vertices: list | None = None) -> dict:
    """Build Upstream+Downstream alignments from the two work-area edge
    points. Prefer over freestyle define_alignment_segment pairs.
    Requires build_wztc_order_table first. approach_length_ft=0 auto-sizes
    from station_walk + slack. force=True wipes an existing corridor.
    path_vertices: optional closed-lane / first-travel outer polyline for
    curved corridors (Align1/2 + hatch follow the path; signs stay
    view-horizontal)."""
    return wztc_ops.assemble_corridor(
        upstream_edge, downstream_edge, approach_length_ft, force,
        path_vertices=path_vertices)


@mcp.tool()
def cross_validate_stations(align_idx: int = 0, tol_ft: float = 0.5,
                            force: bool = False) -> dict:
    """Compare VBA order-table stations vs Python station_walk and check
    path length covers the walk. align_idx=0 checks all. Auto-run by
    place_order_table_stations / place_sheet_geometry."""
    return wztc_ops.cross_validate_stations(align_idx, tol_ft, force)


@mcp.tool()
def place_order_table_stations(align_idx: int, reset_session: bool = False,
                                clear_prior: bool = False,
                                force: bool = False) -> dict:
    """Batched replacement for PlacePerp.frm's interactive walk — places
    perp tick lines at EVERY row in align_idx's order table in one call.
    ALWAYS call this, not repeated place_perp_line calls, once an
    alignment is committed as part of a full-plan run — calling
    place_perp_line once per item defeats the whole point of batching
    and costs real money for no benefit.
    Requires build_wztc_order_table and commit_alignment for this
    align_idx first. reset_session=True for the first alignment in a
    fresh plan run, False for subsequent alignments.
    clear_prior=True wipes journal-owned plan elements first (keeps
    alignments) — REQUIRED when rebuilding, otherwise geometry stacks.
    If stations were already placed for this align this session and
    clear_prior/force are both False, this refuses.
    Returns one row per item; for isSign=Y rows, resolve_sign_code then
    place_sign at the OUTWARD PERP TIP (station + outward*half_len) with
    dir1=outward — never at (ptX,ptY) with dir=tangent. Follow with
    place_order_table_labels, place_order_table_dimensions, and
    place_sheet_symbol_cells for a sheet-faithful plan."""
    return wztc_ops.place_order_table_stations(align_idx, reset_session,
                                                clear_prior, force)


@mcp.tool()
def place_order_table_labels(align_idx: int, outward_sign: float = -1.0,
                             text_extra_along: float = 20.0,
                             sheet_elements: str = "", force: bool = False) -> dict:
    """Sheet-gated Non-Sign labels, X-centered on matching dim midpoints.
    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry). force=True to override."""
    return wztc_ops.place_order_table_labels(align_idx, outward_sign,
                                             text_extra_along, sheet_elements, force)


@mcp.tool()
def place_order_table_dimensions(align_idx: int, outward_sign: float = -1.0,
                                 offset_dist: float = 15.0,
                                 sheet_elements: str = "", force: bool = False) -> dict:
    """Tip-to-tick ny_Plan Linear Size dims (sheet-gated; single text).
    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry). force=True to override."""
    return wztc_ops.place_order_table_dimensions(align_idx, outward_sign,
                                                 offset_dist, sheet_elements, force)


@mcp.tool()
def place_sheet_symbol_cells(align_idx: int, sheet_elements: str,
                             outward_sign: float = -1.0, force: bool = False) -> dict:
    """TWZWVA_P in Vehicle Space bay; TWZAP_P at Shoulder Taper tip.
    Generic heuristic (fixed offset, not lane/shoulder-width-derived), no
    rules-gate validation — refuses when a sheet spec exists for this build
    (prefer place_sheet_geometry). force=True to override."""
    return wztc_ops.place_sheet_symbol_cells(align_idx, sheet_elements, outward_sign, force)


@mcp.tool()
def place_order_table_workspace(align_idx: int, outward_sign: float = -1.0,
                                lane_width: float = 12.0, force: bool = False) -> dict:
    """Hatched work-space box: path start through Vehicle Space, closed lane.
    Generic heuristic, no rules-gate validation — refuses when a sheet spec
    exists for this build (prefer place_sheet_geometry). force=True to override."""
    return wztc_ops.place_order_table_workspace(align_idx, outward_sign, lane_width, force)


@mcp.tool()
def place_order_table_channelizing(align_idx: int, outward_sign: float = -1.0,
                                   lane_width: float = 12.0, force: bool = False) -> dict:
    """Sheet-bounded taper + closed-lane channelizing (not freeform length).
    Prefer place_sheet_geometry when a sheet JSON exists — this generic
    heuristic has no rules-gate validation and refuses when a spec exists
    for this build. force=True to override."""
    return wztc_ops.place_order_table_channelizing(align_idx, outward_sign, lane_width, force)


@mcp.tool()
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
    """Compile sheet-faithful dims/labels/channelizing/symbols/hatch
    (no drawing). Blank designer kwargs fill from locked order-table
    inputs. Prefer place_sheet_geometry(dry_run=True)."""
    return wztc_ops.compile_sheet_plan(
        sheet_num, speed, lane_width, shoulder_width,
        area_type=area_type, closure_type=closure_type,
        exposure_condition=exposure_condition,
        protective_vehicle_gvw=protective_vehicle_gvw,
        align_idxs=align_idxs, outward_sign=outward_sign,
        sheet_elements=sheet_elements,
        arrow_panel_choice=arrow_panel_choice,
        include_primitives=include_primitives, force=force)


@mcp.tool()
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
    """Compile + place sheet-faithful dims/labels/channelizing/symbols/
    hatch from Data/sheet-specs/<sheet>.json. Prefer over
    place_order_table_labels/dimensions/channelizing/workspace/
    place_sheet_symbol_cells when a sheet JSON exists. Still call
    build_wztc_order_table, place_order_table_stations, and place_sign
    separately. dry_run=True = compile + rules gate only."""
    return wztc_ops.place_sheet_geometry(
        sheet_num, speed, lane_width, shoulder_width,
        area_type=area_type, closure_type=closure_type,
        exposure_condition=exposure_condition,
        protective_vehicle_gvw=protective_vehicle_gvw,
        align_idxs=align_idxs, outward_sign=outward_sign,
        sheet_elements=sheet_elements,
        arrow_panel_choice=arrow_panel_choice,
        dry_run=dry_run, force=force, layers=layers)


@mcp.tool()
def delete_construction_guides() -> dict:
    """After a real-road sheet build: delete alignment centerlines and
    perp tick lines only. Leaves signs/cones/hatch/dims/AP/PV/striping."""
    return wztc_ops.delete_construction_guides()


@mcp.tool()
def resolve_sheet_lateral(upstream_edge: list[float],
                          downstream_edge: list[float],
                          closed_side: str,
                          lane_width_ft: float = 0.0,
                          shoulder_width_ft: float = 0.0,
                          real_road_edge: bool = True,
                          yellow_gap_ft: float = 2.0,
                          opposing_lanes: int = 2,
                          path_vertices: list | None = None) -> dict:
    """Lock outward_sign + half_len from travel (up→dn) and closed_side
    (right|left). Call before run_sheet_build on real-road / right-lane
    sheets. real_road_edge uses lane+shoulder for tip-at-EOP half_len.
    Also locks closed_outward so Align2 G20-2 tips on the same closed
    shoulder as Align1 advance signs. path_vertices: same curved polyline
    as assemble_corridor / run_sheet_build when the road bends."""
    return wztc_ops.resolve_sheet_lateral(
        upstream_edge, downstream_edge, closed_side,
        lane_width_ft=lane_width_ft, shoulder_width_ft=shoulder_width_ft,
        real_road_edge=real_road_edge, yellow_gap_ft=yellow_gap_ft,
        opposing_lanes=opposing_lanes, path_vertices=path_vertices)


@mcp.tool()
def run_sheet_build(upstream_edge: list[float] | None = None,
                    downstream_edge: list[float] | None = None,
                    outward_sign: float = -1.0,
                    half_len: float = 40.0,
                    arrow_panel_choice: str = "trailer",
                    include_visual_qa: bool = True,
                    clear_prior_stations: bool = False,
                    force: bool = False,
                    approach_length_ft: float = 0.0,
                    use_locked_lateral: bool = True,
                    path_vertices: list | None = None) -> dict:
    """Sheet-build only executor: assemble→stations→signs→compiler→QA.
    Outside a sheet plan returns sheetPlanActive=False.
    Prefer resolve_sheet_lateral first; locked outward_sign/half_len apply
    when use_locked_lateral=True. path_vertices: curved closed-lane /
    first-travel outer polyline (passed through to assemble_corridor)."""
    return wztc_ops.run_sheet_build(
        upstream_edge=upstream_edge, downstream_edge=downstream_edge,
        outward_sign=outward_sign, half_len=half_len,
        arrow_panel_choice=arrow_panel_choice,
        include_visual_qa=include_visual_qa,
        clear_prior_stations=clear_prior_stations, force=force,
        approach_length_ft=approach_length_ft,
        use_locked_lateral=use_locked_lateral,
        path_vertices=path_vertices)


@mcp.tool()
def run_visual_qa_captures(view_num: int = 1) -> dict:
    """Sheet-build only: scripted corridor/upstream/work-area/downstream
    captures after place_sheet_geometry. No-op outside a sheet plan."""
    return wztc_ops.run_visual_qa_captures(view_num=view_num)


@mcp.tool()
def begin_sheet_sandbox(upstream_edge: list[float] | None = None,
                        downstream_edge: list[float] | None = None,
                        offset_y_ft: float = 2000.0) -> dict:
    """Offset-Y KEEP/REVERT sandbox — does not wipe the kept corridor."""
    return wztc_ops.begin_sheet_sandbox(
        upstream_edge=upstream_edge, downstream_edge=downstream_edge,
        offset_y_ft=offset_y_ft)


@mcp.tool()
def get_sheet_sandbox() -> dict:
    """Active sandbox band state."""
    return wztc_ops.get_sheet_sandbox()


@mcp.tool()
def run_sheet_build_sandbox(offset_y_ft: float = 2000.0,
                            include_visual_qa: bool = True,
                            force: bool = False) -> dict:
    """Try a full sheet build on the sandbox band; then keep or revert."""
    return wztc_ops.run_sheet_build_sandbox(
        offset_y_ft=offset_y_ft, include_visual_qa=include_visual_qa,
        force=force)


@mcp.tool()
def keep_sheet_sandbox() -> dict:
    """KEEP the sandbox try (reference band untouched)."""
    return wztc_ops.keep_sheet_sandbox()


@mcp.tool()
def revert_sheet_sandbox() -> dict:
    """REVERT the sandbox try — clear sandbox placements only."""
    return wztc_ops.revert_sheet_sandbox()


@mcp.tool()
def place_dimension(x1: float, y1: float, x2: float, y2: float,
                    ox: float, oy: float, z: float = 0.0,
                    style_name: str = "ny_Plan", reason: str = "") -> dict:
    """One real Linear Size DimensionElement (prefer order-table batch)."""
    return wztc_ops.place_dimension(x1, y1, x2, y2, ox, oy, z, style_name, reason)


@mcp.tool()
def hatch_element(element_id: str, spacing: float = 10.0, angle_deg: float = 45.0,
                  own_element_only: bool = True, reason: str = "") -> dict:
    """Apply associative hatch to an existing closed shape by element ID.
    Does not create a new element. spacing in master units; angle_deg in degrees."""
    return wztc_ops.hatch_element(element_id, spacing, angle_deg, own_element_only, reason)


@mcp.tool()
def place_arc(x1: float, y1: float, x2: float, y2: float, x3: float, y3: float,
              z: float = 0.0, reason: str = "") -> dict:
    """Place a 3-point arc (placeArcModeEx=3). Point order: start, end, bulge."""
    return wztc_ops.place_arc(x1, y1, x2, y2, x3, y3, z, reason)


@mcp.tool()
def place_text_label(text: str, x: float, y: float, z: float = 0.0,
                     reason: str = "", angle_deg: float = 0.0) -> dict:
    """Place a single-line text label. angle_deg rotates about Z."""
    return wztc_ops.place_text_label(text, x, y, z, reason, angle_deg)


@mcp.tool()
def place_circle(cx: float, cy: float, radius: float, z: float = 0.0, reason: str = "") -> dict:
    """Place a circle (equal-radius ellipse Element API)."""
    return wztc_ops.place_circle(cx, cy, radius, z, reason)


@mcp.tool()
def place_ellipse(cx: float, cy: float, primary_radius: float, secondary_radius: float,
                  angle_deg: float = 0.0, z: float = 0.0, reason: str = "") -> dict:
    """Place an ellipse."""
    return wztc_ops.place_ellipse(cx, cy, primary_radius, secondary_radius, angle_deg, z, reason)


@mcp.tool()
def place_block(x1: float, y1: float, x2: float, y2: float, z: float = 0.0, reason: str = "") -> dict:
    """Place an axis-aligned rectangle."""
    return wztc_ops.place_block(x1, y1, x2, y2, z, reason)


@mcp.tool()
def place_polyline(vertices: list[list[float]], reason: str = "") -> dict:
    """Place an open polyline from [[x,y,z?], ...]."""
    return wztc_ops.place_polyline(vertices, reason)


@mcp.tool()
def place_lane_highway(lanes: int, x1: float = 0.0, y1: float = 0.0,
                       x2: float = 0.0, y2: float = 0.0,
                       lane_width_ft: float = 12.0, shoulder_width_ft: float = 0.0,
                       dash_ft: float = 10.0, gap_ft: float = 30.0,
                       side: str = "right", reason: str = "",
                       vertices: list | None = None) -> dict:
    """Draw an N-lane one-way highway strip (general CAD). Two solid outer
    travel edges + (lanes-1) dashed separators. Optional shoulder_width_ft
    adds solid white EOP outside both outers. Dashes 10/30 real gaps.
    Pass vertices=[[x,y],…] for curved/S first-travel-outer path (overrides
    x1..y2). Ask for missing lanes/width/endpoints/side/path."""
    return wztc_ops.place_lane_highway(
        lanes, x1, y1, x2, y2, lane_width_ft, shoulder_width_ft,
        dash_ft, gap_ft, side, reason, vertices)


@mcp.tool()
def place_two_way_highway(lanes: int, x1: float = 0.0, y1: float = 0.0,
                          x2: float = 0.0, y2: float = 0.0,
                          lane_width_ft: float = 12.0, yellow_gap_ft: float = 2.0,
                          shoulder_width_ft: float = 0.0,
                          dash_ft: float = 10.0, gap_ft: float = 30.0,
                          side: str = "right", reason: str = "",
                          vertices: list | None = None) -> dict:
    """Draw even-N undivided two-way (double solid yellow center). Optional
    shoulder_width_ft. Yellow via resolve_color. Pass vertices=[[x,y],…] for
    curved/S first-travel-outer path. Ask for missing inputs."""
    return wztc_ops.place_two_way_highway(
        lanes, x1, y1, x2, y2, lane_width_ft, yellow_gap_ft,
        shoulder_width_ft, dash_ft, gap_ft, side, reason, vertices)


@mcp.tool()
def place_divided_highway(lanes_per_direction: int, x1: float = 0.0,
                          y1: float = 0.0, x2: float = 0.0, y2: float = 0.0,
                          median_width_ft: float = 0.0,
                          lane_width_ft: float = 12.0,
                          shoulder_width_ft: float = 0.0,
                          dash_ft: float = 10.0, gap_ft: float = 30.0,
                          side: str = "right", reason: str = "",
                          vertices: list | None = None) -> dict:
    """Draw divided multilane/freeway dual carriageway (619-302-style).
    Each dir: white outer, (N-1) dashed, yellow median edge; median_width_ft
    empty gap between yellows (required — ask). Optional shoulders.
    Pass vertices=[[x,y],…] for curved/S first-travel-outer path."""
    return wztc_ops.place_divided_highway(
        lanes_per_direction, x1, y1, x2, y2, median_width_ft,
        lane_width_ft, shoulder_width_ft, dash_ft, gap_ft, side, reason,
        vertices)


@mcp.tool()
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
    TWLT bounded by two dashed yellow lines. Pass vertices=[[x,y],…] for
    curved/S path. Do not use two_way for TWLT."""
    return wztc_ops.place_twlt_highway(
        lanes_per_direction, x1, y1, x2, y2, twlt_width_ft,
        lane_width_ft, shoulder_width_ft, dash_ft, gap_ft, side, reason,
        vertices)


@mcp.tool()
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
    """Draw a + or T intersection (MUTCD box rules). Edges meet the box;
    yellow/dashes stop at stop bar; defaults: crosswalks + stop bars +
    striping turn arrows (ny_plan_striping.cel). SAL/SAR + SLONLY only when
    approach lanes_in > through lanes_out (primary_lanes_out /
    secondary_lanes_out); equal → SAS only. Dotted center when
    has_turning_lanes, TWLT, or dedicated > 0."""
    return wztc_ops.place_orthogonal_intersection(
        junction_x, junction_y, primary_road_type, secondary_road_type,
        primary_length_ft, secondary_stub_ft, primary_bearing_deg,
        junction, tee_side, primary_lanes, secondary_lanes,
        primary_lanes_per_direction, secondary_lanes_per_direction,
        lane_width_ft, yellow_gap_ft,
        primary_median_width_ft, secondary_median_width_ft,
        primary_twlt_width_ft, secondary_twlt_width_ft,
        primary_shoulder_width_ft, secondary_shoulder_width_ft,
        dash_ft, gap_ft, side, crosswalks, stop_bars,
        has_turning_lanes, turn_arrows,
        primary_lanes_out, secondary_lanes_out, reason)


@mcp.tool()
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
    """Draw mainline one-way + diverging ramp at a gore nose (Family 5
    sketch, general CAD). Mainline first edge (x1,y1)->(x2,y2) or
    vertices=[[x,y],…] for a curved mainline; nose at gore_station_ft on
    ramp-side edge; ramp_angle_deg toward side. Ask for missing
    angle/station/lengths."""
    return wztc_ops.place_ramp_gore(
        x1, y1, x2, y2, mainline_lanes, ramp_angle_deg,
        gore_station_ft, ramp_length_ft, ramp_lanes, side,
        gore_mark_ft, lane_width_ft, shoulder_width_ft,
        dash_ft, gap_ft, reason, vertices)


@mcp.tool()
def place_polygon(cx: float, cy: float, radius: float, sides: int, z: float = 0.0, reason: str = "") -> dict:
    """Place a regular n-gon."""
    return wztc_ops.place_polygon(cx, cy, radius, sides, z, reason)


@mcp.tool()
def change_element_symbology(element_id: str, color: int | None = None, weight: int | None = None,
                             line_style_index: int | None = None, line_style_name: str = "",
                             own_element_only: bool = True, reason: str = "") -> dict:
    """Change element color/weight/line style. Prefer line_style_name from
    resolve_line_style over line_style_index."""
    return wztc_ops.change_element_symbology(
        element_id, color, weight, line_style_index, line_style_name,
        own_element_only, reason)


@mcp.tool()
def copy_parallel(element_id: str, distance: float, own_element_only: bool = True, reason: str = "") -> dict:
    """Perpendicular offset-copy of a LINE."""
    return wztc_ops.copy_parallel(element_id, distance, own_element_only, reason)


@mcp.tool()
def crosshatch_element(element_id: str, spacing: float = 10.0, angle_deg: float = 45.0,
                       own_element_only: bool = True, reason: str = "") -> dict:
    """Apply crosshatch to a closed element."""
    return wztc_ops.crosshatch_element(element_id, spacing, angle_deg, own_element_only, reason)


@mcp.tool()
def remove_hatch(element_id: str, own_element_only: bool = True, reason: str = "") -> dict:
    """Remove associative hatch from a closed element."""
    return wztc_ops.remove_hatch(element_id, own_element_only, reason)


@mcp.tool()
def break_line(element_id: str, x: float, y: float, z: float = 0.0,
               own_element_only: bool = True, reason: str = "") -> dict:
    """Break a line into two at (x,y)."""
    return wztc_ops.break_line(element_id, x, y, z, own_element_only, reason)


@mcp.tool()
def extend_line(element_id: str, new_length: float, own_element_only: bool = True, reason: str = "") -> dict:
    """Set line length from start (extend or shorten)."""
    return wztc_ops.extend_line(element_id, new_length, own_element_only, reason)


@mcp.tool()
def fillet_elements(element_id1: str, element_id2: str, radius: float,
                    pick_x: float, pick_y: float, pick_z: float = 0.0,
                    own_element_only: bool = True, reason: str = "") -> dict:
    """Create fillet arc between two elements (no auto-trim)."""
    return wztc_ops.fillet_elements(element_id1, element_id2, radius, pick_x, pick_y, pick_z, own_element_only, reason)


@mcp.tool()
def create_complex_string(element_ids: list[str], reason: str = "") -> dict:
    """Create a complex string from element IDs."""
    return wztc_ops.create_complex_string(element_ids, reason)


@mcp.tool()
def place_fence_block(x1: float, y1: float, x2: float, y2: float, z: float = 0.0,
                      view_num: int = 1, reason: str = "") -> dict:
    """Define a rectangular fence."""
    return wztc_ops.place_fence_block(x1, y1, x2, y2, z, view_num, reason)


@mcp.tool()
def fence_undefine(reason: str = "") -> dict:
    """Clear the fence."""
    return wztc_ops.fence_undefine(reason)


@mcp.tool()
def fence_copy_contents(delta_x: float, delta_y: float, delta_z: float = 0.0, reason: str = "") -> dict:
    """Copy elements inside the fence by delta."""
    return wztc_ops.fence_copy_contents(delta_x, delta_y, delta_z, reason)


@mcp.tool()
def fence_move_contents(delta_x: float, delta_y: float, delta_z: float = 0.0, reason: str = "") -> dict:
    """Move elements inside the fence by delta."""
    return wztc_ops.fence_move_contents(delta_x, delta_y, delta_z, reason)


@mcp.tool()
def fence_delete_contents(reason: str = "") -> dict:
    """Delete elements inside the fence (not undoable)."""
    return wztc_ops.fence_delete_contents(reason)


@mcp.tool()
def select_element(element_id: str, clear_first: bool = True, reason: str = "") -> dict:
    """Select an element by ID."""
    return wztc_ops.select_element(element_id, clear_first, reason)


@mcp.tool()
def clear_selection(reason: str = "") -> dict:
    """Clear the selection set."""
    return wztc_ops.clear_selection(reason)


@mcp.tool()
def place_element_run(element_idx: int, vertices: list[list[float]], reason: str = "") -> dict:
    """Place a channelizing-device / removal-striping / barrier run.
    element_idx: 2=Channelizing Devices, 3=Removal Striping, 4=Temporary
    Barrier, 5=Barrier w/Warning Lights (1=Work Space — use place_workspace
    instead). vertices is an ordered list of [x, y, z] points."""
    return wztc_ops.place_element_run(element_idx, vertices, reason)


@mcp.tool()
def place_cell(cell_name: str, pt_x: float, pt_y: float, pt_z: float = 0, angle_deg: float = 0,
               library_path: str = "", reason: str = "") -> dict:
    """Place a cell at (pt_x, pt_y). Default library is WZTC; pass
    library_path for ny_plan_striping.cel etc."""
    return wztc_ops.place_cell(cell_name, pt_x, pt_y, pt_z, angle_deg,
                               library_path, reason)


@mcp.tool()
def place_cell_on_post(cell_name: str, pt_x: float, pt_y: float, dir_x: float, dir_y: float,
                        pt_z: float = 0, angle_deg: float = 0, reason: str = "") -> dict:
    """Place a cell on a 50 ft stem/post the same way a roadside sign is
    built, instead of a bare lateral offset. (pt_x, pt_y) is the base/tick
    point; (dir_x, dir_y) is the outward unit direction."""
    return wztc_ops.place_cell_on_post(cell_name, pt_x, pt_y, dir_x, dir_y, pt_z, angle_deg, reason)


@mcp.tool()
def set_sign_attributes(element_ids: list[str], reason: str = "") -> dict:
    """Finish symbology after place_sign: white labels/stems, orange post.
    Face cells stay library symbology — do not also recolor faces."""
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
def clear_plan_elements(keep_alignments: bool = True, align_idx: int = 0) -> dict:
    """Idempotent-rebuild wipe: delete journal-owned create-ops. Default
    keep_alignments=True leaves the corridor alone. Pass align_idx=1|2 to
    clear only that alignment (so Downstream rebuild does not wipe
    Upstream). align_idx=0 clears the whole plan. Call BEFORE re-placing
    when iterating — otherwise geometry stacks. place_order_table_stations(
    clear_prior=True) scopes automatically to its align_idx."""
    return wztc_ops.clear_plan_elements(keep_alignments, align_idx)

@mcp.tool()
def get_journal(limit: int = 50) -> list[str]:
    """Return the last `limit` raw journal lines — every op run this
    session, its full parameters (including any reason= passed), and its
    result. This is the PE audit trail: use it to answer "why is that sign
    there" or to review a whole session before handing off a sheet.

    Default is intentionally wider than wztc_ops.get_journal's own default
    (20) -- MCP clients reviewing a session tend to want more context per
    call than the chat agent's own internal uses of this function do."""
    return wztc_ops.get_journal(limit)


@mcp.tool()
def list_deferred_handoffs() -> list[dict]:
    """List every dimension/callout queued by handoff() this session that
    still needs a few manual clicks through the existing interactive forms."""
    return wztc_ops.list_deferred_handoffs()


# ============================================================ Registry / Edit (M6)

@mcp.tool()
def list_registry_commands(safety_status: str = "", opname_contains: str = "") -> list[dict]:
    """List MicroStation command recipes in Data/command-registry.tsv.
    Optional safety_status filter (e.g. 'verified-headless-safe',
    'needs-testing', 'interactive-only-use-handoff'). Only
    verified-headless-safe rows can be executed via run_registry_command;
    interactive-only rows point at handoff() instead.

    opname_contains narrows to opNames containing this substring
    (case-insensitive) -- e.g. 'ZOOM', 'PAN', 'LEVEL'. Strongly
    recommended: this registry has ~1800 rows, and returning them all
    costs real tokens -- an unfiltered call was measured live at ~240K
    input tokens (~$0.75) for a single turn. If you have any idea what
    the command name might contain, pass it here rather than listing
    everything."""
    return wztc_ops.list_registry_commands(safety_status, opname_contains)


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
def copy_element(element_id: str, delta_x: float, delta_y: float, delta_z: float = 0,
                 own_element_only: bool = True, reason: str = "") -> dict:
    """Copy an element by ID (Clone + Move). Returns newElementId.
    For engineer-picked pre-existing elements pass own_element_only=False.
    Resolve geometry with get_elements_range first."""
    return wztc_ops.copy_element(element_id, delta_x, delta_y, delta_z, own_element_only, reason)


@mcp.tool()
def rotate_element(element_id: str, origin_x: float, origin_y: float, angle_deg: float,
                   origin_z: float = 0, own_element_only: bool = True, reason: str = "") -> dict:
    """Rotate an element about a point by angle_deg (Z axis)."""
    return wztc_ops.rotate_element(element_id, origin_x, origin_y, angle_deg,
                                   origin_z, own_element_only, reason)


@mcp.tool()
def scale_element(element_id: str, origin_x: float, origin_y: float, scale_factor: float,
                  origin_z: float = 0, own_element_only: bool = True, reason: str = "") -> dict:
    """Uniform-scale an element about a point."""
    return wztc_ops.scale_element(element_id, origin_x, origin_y, scale_factor,
                                  origin_z, own_element_only, reason)


@mcp.tool()
def mirror_element(element_id: str, x1: float, y1: float, x2: float, y2: float,
                   z1: float = 0, z2: float = 0, own_element_only: bool = True,
                   reason: str = "") -> dict:
    """Mirror an element about the axis through (x1,y1)-(x2,y2)."""
    return wztc_ops.mirror_element(element_id, x1, y1, x2, y2, z1, z2, own_element_only, reason)


@mcp.tool()
def array_element(element_id: str, count: int, spacing_x: float, spacing_y: float,
                  own_element_only: bool = True, reason: str = "") -> dict:
    """Create count copies offset by i*(spacing_x, spacing_y)."""
    return wztc_ops.array_element(element_id, count, spacing_x, spacing_y, own_element_only, reason)


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
    'part6' | 'supplement' | 'stdsht' (empty searches all three). If the
    index is missing, returns one hit with heading INDEX_MISSING (run
    ingest_manuals.py). Multi-word queries that miss under FTS5 AND are
    retried with OR / phrase matching. A genuine empty list means no
    match after those retries — not "manuals unavailable"."""
    return manual_search.search(query, source=source, max_results=max_results)


if __name__ == "__main__":
    mcp.run()
