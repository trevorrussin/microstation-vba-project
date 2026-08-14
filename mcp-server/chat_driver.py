"""
M7 Stage 4 — the agent loop behind the in-MicroStation chat panel
(UserForms/WZTCChatPanel.frm).

"Python owns the brain and the hands; VBA owns only the face" (the M7
plan's architecture decision): this process holds the actual Claude Opus 5
agentic loop, and drives every tool call through the exact same
WZTCBridge.bas / bridge_client.py mechanism M1-M6 already proved live —
via chat_bridge (bridge_client.py), not the module-level `bridge` singleton
server.py's stdio MCP connection uses, so the two processes never race on
the same request.tsv/response.tsv if both happen to be running against the
same MicroStation session at once (see bridge_client.py / WZTCBridge.
RunChatToolRequest).

Protocol (Bridge/, same TSV convention as the rest of this bridge):
  chat-input.tsv  <- WZTCChatPanel.frm appends "<timestamp>\tmessage" on Send
  chat-log.tsv    -> this process appends structured lines WZTCChatPanel.frm
                     polls and renders (see chat_log.ChatLog for the schema)
  chat-history.json -> persisted conversation, reloaded on restart

Run: python chat_driver.py   (persistent process; user-started, not
auto-launched by VBA — see the plan's M7 Stage 6 lifecycle decision)
Requires ANTHROPIC_API_KEY — set as a real environment variable, via
`ant auth login`, or in mcp-server/.env (gitignored; copy .env.example and
fill in your key there — never commit the real file, never paste the key
value into a chat/PR/log). load_dotenv() below only fills in ANTHROPIC_API_KEY
if it isn't already set as a real env var, so a system-level key always wins.

Module split (2026-08-04): this file used to hold everything -- the full
~500-line system prompt, cost tracking, log I/O, input polling, and the
whole history-trimming cluster, alongside the actual agent loop. Those are
now separate, self-contained modules (prompts.py, usage.py, chat_log.py,
input_watcher.py, chat_history.py) that this file imports and wires
together; what's left here is genuinely the driver: tool registration, the
per-turn loop, session-mode/session-state, and process startup.
"""
from __future__ import annotations

import base64
import json
import os
from dataclasses import dataclass, field
from pathlib import Path

import anthropic
from anthropic import beta_tool
from dotenv import load_dotenv

import chat_history
import manual_search
import view_capture
import wztc_ops
from bridge_client import chat_bridge
from chat_log import ChatLog
from input_watcher import InputWatcher
from prompts import MODE_SYSTEM_PROMPT
from usage import UsageTracker

load_dotenv(Path(__file__).parent / ".env")
wztc_ops.set_bridge(chat_bridge)

BRIDGE_DIR = Path(r"c:\repos\microstation-vba-project\Bridge")
CHAT_INPUT_FILE = BRIDGE_DIR / "chat-input.tsv"
CHAT_LOG_FILE = BRIDGE_DIR / "chat-log.tsv"
HISTORY_FILE = BRIDGE_DIR / "chat-history.json"
SESSION_MODE_FILE = BRIDGE_DIR / "chat-session-mode.txt"
USAGE_FILE = BRIDGE_DIR / "chat-usage.tsv"
CHAT_LOG_ARCHIVE_DIR = BRIDGE_DIR / "archive"
CHAT_LOG_MAX_BYTES = 2_000_000  # ~2MB; see ChatLog._rotate_if_oversized

# Overridable without a code edit -- e.g. `set WZTC_CHAT_MODEL=claude-opus-5`
# before running this script to switch back for an A/B comparison. Default
# switched to Sonnet + medium effort 2026-08-01 after cost review; Opus
# remains available via the env var with no code change needed.
MODEL = os.environ.get("WZTC_CHAT_MODEL", "claude-sonnet-5")
EFFORT = os.environ.get("WZTC_CHAT_EFFORT", "medium")
MAX_TOKENS = 16000

# Safety net, not a typical-case tuning knob: the sign-placement debugging
# turn earlier today ran ~10 legitimate tool round-trips investigating a
# real error; this caps a genuinely runaway loop (e.g. stuck retrying a
# failing approach) at roughly 3x that, each round-trip being a separate
# billed API call. Hitting this stops the tool_runner's generator loop
# cleanly (confirmed by reading _beta_runner.py -- no exception, it just
# stops yielding), not a crash -- run_turn below detects the resulting
# empty final_text and substitutes a message that says so explicitly,
# since silently returning nothing would show a blank response in the panel.
#
# 2026-08-03: lowered from 30 after a ~$12 live sheet-build session where
# long tool loops + a 600-message history dominated cost. 2026-08-04: raised
# back to 26 -- the actual cost driver in that session was the prompt-cache
# -busting history-trim bug (see chat_history._trim_history_window), now
# fixed and verified holding cacheRead stable across a turn; a full
# build+visual-QA cycle (build table, place geometry, capture screenshot,
# review, fix) can legitimately need more round-trips than a single
# sign-placement debug turn. Override with WZTC_MAX_TOOL_ITERATIONS if a
# rare deep investigation needs more still.
MAX_TOOL_ITERATIONS = int(os.environ.get("WZTC_MAX_TOOL_ITERATIONS", "26"))

USAGE = UsageTracker(USAGE_FILE)
LOG = ChatLog(CHAT_LOG_FILE, CHAT_LOG_ARCHIVE_DIR, CHAT_LOG_MAX_BYTES)
INPUT = InputWatcher(CHAT_INPUT_FILE)


@dataclass
class SessionState:
    """The two module-level mutable globals this file carried before the
    2026-08-04 split -- touched_element_ids (per-turn, cleared at the start
    of every run_turn) and mode (persists across turns, see
    load_session_mode/save_session_mode). Bundled into one object so
    mutation sites are explicit attribute writes instead of `global`
    reassignment scattered across enter_mode/exit_mode/main."""
    touched_element_ids: set[str] = field(default_factory=set)
    mode: str = "general"


_SESSION = SessionState()


def _collect_element_ids(result) -> None:
    """Pulls any element IDs out of a tool result (createdElementIds,
    elementIds, elementId -- the conventions already used across every
    WZTCBridge op's response) into _SESSION.touched_element_ids, so the
    post-turn auto-focus/screenshot hook (see main()) knows what to pan the
    view to. Added 2026-08-02 after feedback that the view never followed
    the agent's work, so the engineer watching the panel saw whatever was
    on screen before the turn started, not what changed."""
    if not isinstance(result, dict):
        return
    for key in ("createdElementIds", "elementIds"):
        val = result.get(key)
        if val:
            for eid in str(val).split(","):
                eid = eid.strip()
                if eid:
                    _SESSION.touched_element_ids.add(eid)
    eid = result.get("elementId")
    if eid not in (None, ""):
        _SESSION.touched_element_ids.add(str(eid).strip())


def _png_to_bmp(png_path: Path) -> Path:
    """VBA's LoadPicture (used by WZTCChatPanel's image display) does not
    reliably support PNG in this MSForms host -- confirmed live 2026-08-02:
    the panel showed no error (LoadPicture's failure was caught by its own
    On Error Resume Next) but also no image, ever. BMP is LoadPicture's one
    universally-supported format across VBA hosts, so every screenshot/
    reference-image path converts to BMP specifically for the panel rather
    than changing the PNG capture format everyone else (the model, Claude
    Code) relies on."""
    bmp_path = png_path.with_suffix(".bmp")
    view_capture.Image.open(png_path).convert("RGB").save(bmp_path, format="BMP")
    return bmp_path


def _log_screenshot(png_path: Path) -> None:
    """Convert + log the common case -- most screenshot call sites just
    want a BMP written to the panel's SCREENSHOT log line. Only
    _show_reference_image needs a different log call (REFERENCE_IMAGE with
    source/heading/page), so it calls _png_to_bmp directly instead."""
    LOG.screenshot(str(_png_to_bmp(png_path)))


def _show_reference_image(tool_name: str, result) -> None:
    """After a search_reference_manual call, render the top hit's actual
    PDF page and log it so the panel shows the manual/sheet page an answer
    is grounded in -- not just the text excerpt (2026-08-02 feedback,
    same session as _auto_focus_and_capture). Best-effort and non-fatal,
    same pattern as that function: a missing/gitignored PDF, an
    out-of-range page, or any rendering hiccup is purely a lost visual
    aid, never a reason to fail the turn or hide the text excerpt the
    model already has."""
    if tool_name != "search_reference_manual":
        return
    if not isinstance(result, list) or not result:
        return
    hit = result[0]
    if hit.get("heading") == "INDEX_MISSING":
        return
    try:
        png_path = manual_search.render_page_image(
            hit["source"], int(hit["page_start"]),
            BRIDGE_DIR / "captures" / "reference_live.png",
        )
        bmp_path = _png_to_bmp(png_path)
        LOG.reference_image(str(bmp_path), hit.get("source_name", hit["source"]), hit.get("heading", ""), hit["page_start"])
    except Exception as e:
        LOG.error(f"reference-image render failed (non-fatal, text excerpt still stands): {e}")


def _show_view_update(tool_name: str, result) -> None:
    """After a successful adjust_view call, refresh the panel's screenshot
    so the engineer sees the new zoom/pan immediately, same as the
    post-turn auto-focus screenshot does for element changes. adjust_view
    already includes its own ~2s settle delay (view_capture.navigate_view),
    so the capture here sees the real post-adjustment state, not a stale
    repaint. Best-effort and non-fatal, same pattern as
    _show_reference_image -- a failed screenshot here must never make the
    view adjustment itself look like it failed."""
    if tool_name != "adjust_view":
        return
    if not isinstance(result, dict) or result.get("status") != "OK":
        return
    try:
        _log_screenshot(view_capture.capture_microstation())
    except Exception as e:
        LOG.error(f"post-adjust_view screenshot failed (non-fatal, view was still adjusted): {e}")


def _qa_capture_rows(result) -> list[dict]:
    """Pull scripted visual-QA frame dicts (with path) from a tool result."""
    if not isinstance(result, dict):
        return []
    caps = result.get("captures")
    if isinstance(caps, list) and caps:
        return [c for c in caps if isinstance(c, dict) and c.get("path")]
    return []


def _vision_blocks_for_qa_captures(result: dict) -> list[dict] | None:
    """Log each QA frame to the panel SCREENSHOT stream and build Anthropic
    vision content blocks so the model actually sees the scripted frames.

    Previously run_visual_qa_captures only returned file paths — the agent
    marked visualQaPassed without receiving images. Caps at 4 frames (the
    scripted set). Best-effort per frame: a missing file skips that image
    but keeps the text payload."""
    rows = _qa_capture_rows(result)
    if not rows:
        return None

    blocks: list[dict] = [
        {
            "type": "text",
            "text": json.dumps(
                {k: v for k, v in result.items() if k != "captures"},
                ensure_ascii=False, default=str,
            ),
        },
        {
            "type": "text",
            "text": (
                f"Scripted visual QA: {len(rows)} frame(s) attached below. "
                "Review each against the checklist before FINAL. "
                "Do not call capture_view (not a chat tool) — these frames "
                "replace an extra view_drawing for sheet QA."
            ),
        },
    ]
    for row in rows[:4]:
        frame = str(row.get("frame") or "frame")
        path = Path(str(row["path"]))
        try:
            if not path.is_file():
                blocks.append({
                    "type": "text",
                    "text": f"[QA frame {frame}: file missing at {path}]",
                })
                continue
            _log_screenshot(path)
            data = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
            blocks.append({"type": "text", "text": f"QA frame: {frame} ({path.name})"})
            blocks.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": data,
                },
            })
        except Exception as e:
            LOG.error(f"QA frame {frame} vision attach failed (non-fatal): {e}")
            blocks.append({
                "type": "text",
                "text": f"[QA frame {frame}: attach failed: {e}]",
            })
    return blocks


def _summarize(result) -> str:
    """Short one-line summary of a tool result for the TOOL_RESULT log
    line -- the full result already went to the model; this is just for
    the human watching the transcript."""
    if isinstance(result, dict):
        if "note" in result:
            return str(result["note"])
        keys = ", ".join(f"{k}={v}" for k, v in list(result.items())[:4])
        return keys
    if isinstance(result, list):
        return f"{len(result)} result(s)"
    return str(result)[:200]


def _wrap_op(tool_name: str, fn):
    """Wrap a wztc_ops (or manual_search) function as a @beta_tool: same
    signature/docstring — functools.wraps copies __annotations__/__doc__
    onto the wrapper directly, and inspect.signature() follows __wrapped__,
    so @beta_tool's schema generation sees fn's real signature regardless
    of which mechanism it introspects with — with TOOL_CALL/TOOL_RESULT
    logging and error-to-string conversion added so a failed call surfaces
    to the model as a normal (if unhappy) tool result rather than crashing
    the whole turn.

    functools.wraps also copies fn.__name__ onto the wrapper, which is
    wrong whenever tool_name differs from fn's actual Python name (e.g.
    manual_search.search wrapped as the tool "search_reference_manual" —
    confirmed live: without this override the tool silently registered
    itself as "search", not the name callers/the model expect). Set
    __name__ explicitly, after @functools.wraps, so it always matches the
    tool_name this function was actually called with.

    The wrapper's return value is JSON-encoded to a str rather than
    returned as a raw dict/list: @beta_tool's documented result type is
    `str | Iterable[BetaContent]` (each content block needs its own
    "type" field, e.g. {"type": "text", ...}) -- a plain dict like
    {"status": "OK", "elementId": 7} doesn't satisfy that shape and the
    API rejects it with "content.0.tool_result.content.0.type: Field
    required" on the *next* turn (confirmed live). A JSON string is
    valid content and Claude parses embedded JSON fine.

    Exception: run_visual_qa_captures / run_sheet_build results that carry
    scripted QA capture paths return multimodal content (text + images)
    so the model can see the frames; paths are also logged as SCREENSHOT
    for the panel."""
    import functools

    @functools.wraps(fn)
    def wrapper(**kwargs):
        LOG.tool_call(tool_name, kwargs)
        try:
            result = fn(**kwargs)
            LOG.tool_result(tool_name, "OK", _summarize(result))
            _collect_element_ids(result)
            _show_reference_image(tool_name, result)
            _show_view_update(tool_name, result)
            if tool_name in ("run_visual_qa_captures", "run_sheet_build"):
                vision = _vision_blocks_for_qa_captures(result) if isinstance(result, dict) else None
                if vision is not None:
                    return vision
            return json.dumps(result, ensure_ascii=False, default=str)
        except Exception as e:
            LOG.tool_result(tool_name, "ERROR", str(e))
            return json.dumps({"status": "ERROR", "note": str(e)})

    wrapper.__name__ = tool_name
    return beta_tool(wrapper)


def _validate_op_names(names: list[str], source_module, label: str) -> None:
    """Fail loudly at import time if any name in a hand-maintained
    _BASE_OP_NAMES/_WZTC_OP_NAMES list doesn't actually exist on
    source_module -- these lists must stay in sync with wztc_ops.py's real
    exports by hand, and that has already silently drifted once live (see
    the "Added 2026-08-02" comments on _BASE_OP_NAMES below: several ops
    existed in wztc_ops.py/server.py but were never added here, so the
    chat agent silently could not call them despite server.py exposing
    them). Better to crash chat_driver.py's startup with a clear list of
    what's missing than to ship a tool the model discovers is broken
    mid-conversation."""
    missing = [n for n in names if not hasattr(source_module, n)]
    if missing:
        raise ImportError(
            f"{label} names not found on {source_module.__name__}: {missing} -- "
            f"fix the name (typo?) or add the function to {source_module.__name__} "
            f"before it can be registered as a tool."
        )


# Session modes (2026-08-02) -- _BASE_OP_NAMES is loaded in every mode;
# _WZTC_OP_NAMES only loads once the agent calls enter_mode("wztc"). See
# MODE_INFO / _MODE_TOOLS below and the "Session modes" plan for the
# rationale (general MicroStation agent vs. a WZTC-specific pack layered
# on top, so an unrelated session doesn't carry WZTC's tool schemas/rules
# it'll never use).
_BASE_OP_NAMES = [
    "find_elements_near", "get_elements_range", "get_element_vertices",
    "focus_view_on_elements",
    "station_to_point", "get_alignment_stationing",
    "get_alignment_vertices", "get_locked_designer_inputs", "get_plan_status",
    "propose_corridor_source", "lock_corridor_path",
    "propose_work_area_on_path", "snap_work_area_to_path",
    "get_placements", "delete_placements",
    "get_geometry_scorecard", "reflect_sheet_build",
    "check_build_overlap", "get_elements_in_range_box",
    "list_levels", "list_colors", "resolve_color",
    "list_line_styles", "resolve_line_style",
    "cell_library_status", "attach_cell_library", "list_cells",
    "list_cell_libraries", "find_cell",
    "list_fonts", "resolve_font", "list_text_styles", "resolve_text_style",
    "describe_drawing_state", "classify_site_features",
    "handoff",
    "undo_last_op", "get_journal", "list_deferred_handoffs",
    "list_registry_commands", "describe_registry_command", "run_registry_command",
    "move_element", "change_element_level", "edit_text_element", "delete_element",
    # Added 2026-08-02 -- these existed in wztc_ops.py/server.py (the MCP
    # interface) but were never added here, so the actual chat agent could
    # not call any of them despite server.py exposing them. Found live
    # while testing: the agent noticed describe_drawing_state wasn't
    # actually callable even though the system prompt told it to use it.
    "hatch_element", "place_arc", "place_text_label",
    "place_circle", "place_ellipse", "place_block", "place_polyline", "place_polygon",
    "place_lane_highway",
    "place_two_way_highway",
    "place_divided_highway",
    "place_twlt_highway",
    "place_orthogonal_intersection",
    "place_ramp_gore",
    "place_cell",
    "change_element_symbology", "copy_parallel", "crosshatch_element", "remove_hatch",
    "break_line", "extend_line", "fillet_elements", "create_complex_string",
    "place_fence_block", "fence_undefine", "fence_copy_contents",
    "fence_move_contents", "fence_delete_contents",
    "select_element", "clear_selection",
    "copy_element", "rotate_element", "scale_element", "mirror_element", "array_element",
    # Added 2026-08-02 -- see SYSTEM_PROMPT's registry paragraph. Reliable
    # replacement for the now-disabled ZOOM_*/PAN_VIEW_* registry key-ins.
    # 2026-08-04: absolute center_x/center_y/width/height added after agent
    # passed model coords as relative pan and flung the view.
    "adjust_view",
]

# WZTC-specific ops -- only meaningful once compute_spacing/place_sign's
# domain rules (the engineering-judgment boundary, road_type handling)
# are actually in play. Kept out of _BASE_OP_NAMES so a general-mode
# session never carries these schemas or the strict rules that go with
# them.
_WZTC_OP_NAMES = [
    "compute_spacing", "get_sheet_requirements", "get_sheet_build_guide",
    "get_required_designer_inputs",
    "resolve_sign_code",
    "place_perp_line", "place_sign", "place_workspace", "place_element_run",
    "place_cell_on_post", "set_sign_attributes",
    # Added 2026-08-02 -- agent-driven-8-step-wizard plan (Components 1-3):
    # orchestrate the full WZTCDesigner->DrawWorkSpace->AlignDraw->PlacePerp
    # sequence without opening any form. See prompts.WZTC_SYSTEM_PROMPT_ADDENDUM's
    # full-plan-flow section for call order.
    "build_wztc_order_table", "find_reference_linework",
    "define_alignment_segment", "commit_alignment", "adopt_alignment",
    "assemble_corridor", "cross_validate_stations",
    "resolve_sheet_lateral",
    "place_order_table_stations",
    "place_order_table_labels", "place_order_table_dimensions",
    "place_sheet_symbol_cells", "place_order_table_workspace",
    "place_order_table_channelizing", "place_dimension",
    # 2026-08-04: placement-plan compiler (sheet_spec Stages 1-5) — prefer
    # over place_order_table_labels/dimensions/channelizing/workspace/symbol_cells
    # when Data/sheet-specs/<sheet>.json exists.
    "compile_sheet_plan", "place_sheet_geometry",
    "run_sheet_build", "run_visual_qa_captures",
    "begin_sheet_sandbox", "get_sheet_sandbox", "run_sheet_build_sandbox",
    "keep_sheet_sandbox", "revert_sheet_sandbox",
    "clear_plan_elements",
    "delete_construction_guides",
]

_validate_op_names(_BASE_OP_NAMES, wztc_ops, "_BASE_OP_NAMES")
_validate_op_names(_WZTC_OP_NAMES, wztc_ops, "_WZTC_OP_NAMES")


@beta_tool
def ask_user(question: str) -> str:
    """Ask the engineer a clarifying question and wait for their reply
    before continuing. Use this for genuine ambiguity you cannot resolve
    yourself — e.g. which of several close-by candidates to act on, or a
    site condition that needs the engineer's judgment call — not for
    routine decisions you're equipped to make on your own. Blocks until
    the engineer responds in the chat panel; their reply is returned as
    this tool's result."""
    LOG.ask_user(question)
    return INPUT.wait_for_next()


@beta_tool
def ask_user_choice(question: str, options: list[dict] | None = None,
                    allow_point_pick: bool = False,
                    allow_element_pick: bool = False) -> str:
    """Ask the engineer to pick one of a small number of concrete options via
    clickable buttons in the chat panel, instead of a free-form question --
    use this the way you'd use a structured choice UI yourself: for a real
    decision with distinct options (e.g. "which of these two sign clusters
    did you mean"), not for every question -- plain ask_user or just asking
    in your final text is still right for anything else. Each option is
    {"label": short button text, "description": longer context shown in the
    transcript}. Up to 4 options (the panel has 4 button slots -- ask a
    narrower follow-up if you have more).

    Set allow_point_pick=True to show a "Click a point in the drawing"
    button -- reply is coordinates "(x, y, z)". Set allow_element_pick=True
    to add a choice-button option "Identify an element in the drawing"
    (uses an empty btnChoice slot; leave at most 3 of your own options so
    it fits) -- reply is "elementId=… type=… level=… [cell=…]". Prefer
    element pick when the engineer is pointing at an existing sign/cell/
    line. When you ONLY need a location or an identify, pass options=[]
    with the matching allow_* flag. Do NOT invent a fake option like
    "I'll click the point/element" — that only echoes text and dismisses
    the real pick UI (live failure 2026-08-02).

    The engineer can ALWAYS ignore the buttons and type a free-form reply
    instead (the input box never goes away) -- treat whatever comes back as
    the answer, whatever form it takes: an option's exact label text,
    picked coordinates, elementId=… text, or free text. Blocks until they
    respond, same as ask_user."""
    opts = list(options or [])
    if not allow_point_pick and not allow_element_pick and not opts:
        return (
            "ask_user_choice needs options and/or allow_point_pick=True "
            "and/or allow_element_pick=True. For a free-form question use "
            "ask_user instead."
        )
    # Drop option labels that just duplicate the pick buttons — those are
    # what caused the "choice is gone" failure when the engineer clicked
    # them and dismissed btnPickPoint / btnPickElement.
    cleaned = []
    for opt in opts[:4]:
        label = str(opt.get("label", "")).strip()
        low = label.lower()
        if (allow_point_pick or allow_element_pick) and any(
            phrase in low
            for phrase in (
                "i'll click", "i will click", "click the point",
                "click a point", "use the pick", "point pick",
                "point-pick", "click it in the drawing",
                "identify the element", "identify an element",
                "select the element", "pick the element",
                "click the element", "i'll identify", "i will identify",
            )
        ):
            continue
        cleaned.append(opt)
    LOG.ask_user_choice(question, cleaned, allow_point_pick, allow_element_pick)
    return INPUT.wait_for_next()


@beta_tool
def view_drawing() -> list[dict] | str:
    """Take a screenshot of the current MicroStation view and look at it
    yourself, to visually verify your own work -- element placement,
    spacing, obvious overlaps, whether something looks wrong. This costs
    real image tokens (roughly 1500-2000 per call), so use it selectively:
    after a substantial design change (several elements placed or moved
    this turn) or when you suspect something might be off, not after every
    small edit and not as a routine end-of-turn habit. Prefer at most ONE
    or TWO view_drawing calls per turn — each costs ~1500-2000 image tokens
    and prior screenshots are stripped from history anyway. If this turn
    touched any elements, the view is first panned/zoomed to show them
    (same framing the engineer sees in the panel); otherwise it captures
    whatever is currently on screen."""
    try:
        if _SESSION.touched_element_ids:
            ids_csv = ",".join(sorted(_SESSION.touched_element_ids))
            resp = chat_bridge.call("GET_ELEMENTS_RANGE", elementIds=ids_csv)
            if resp.get("status") == "OK":
                low_x, low_y = float(resp["lowX"]), float(resp["lowY"])
                high_x, high_y = float(resp["highX"]), float(resp["highY"])
                width = max(high_x - low_x, 10.0) * 1.3
                height = max(high_y - low_y, 10.0) * 1.3
                view_capture.navigate_view((low_x + high_x) / 2, (low_y + high_y) / 2, width, height)
        path = view_capture.capture_microstation()
    except Exception as e:
        return f"Screenshot capture failed: {e}"

    # Same BMP copy the panel shows after every turn (LoadPicture needs BMP,
    # not PNG -- see _png_to_bmp) so the engineer sees exactly what you're
    # looking at, not a different/stale image.
    _log_screenshot(path)

    # capture_microstation() already resizes to view_capture.MAX_LONG_EDGE
    # (1568px, the point past which Anthropic's own resize makes a larger
    # upload pure waste) -- no separate downscale needed here.
    data = base64.standard_b64encode(path.read_bytes()).decode("utf-8")
    return [
        {"type": "text", "text": "Current MicroStation view:"},
        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": data}},
    ]


# Session modes (2026-08-02): _SESSION.mode is mutated by enter_mode/
# exit_mode below, read by run_turn() when building EACH turn's tools/
# system.
#
# Persisted to SESSION_MODE_FILE (added after a real incident, same day):
# conversation HISTORY already survives a chat_driver.py restart
# (HISTORY_FILE), but the session mode used not to -- every fresh process
# silently reset to "general" while the reloaded history still showed
# turns from when the agent was in "wztc" mode with working tools. The
# model, going by its own (accurate) memory of those turns, had no reason
# to think it needed to call enter_mode again, tried a WZTC tool directly,
# got a real failure, and concluded tooling was broken -- confirmed live
# 2026-08-02 (see Claude Code memory / dev-notes/agent-log.md for the full
# incident). Loading the saved mode at startup keeps mode state consistent
# with the history it's paired with, closing that whole mismatch class
# rather than just prompting the model to guess better under it.

MODE_INFO = {
    "general": "General MicroStation drawing and query -- no domain-specific tools loaded.",
    "wztc": "Workzone traffic control design -- sign placement, spacing/taper calculations, MUTCD/NYSDOT lookups.",
}


def load_session_mode() -> str:
    """Read the last-persisted mode (SESSION_MODE_FILE), defaulting to
    'general' if absent, unreadable, or not a recognized mode -- a
    corrupt/stale value here must never crash startup or silently pick
    an unknown mode with no tools defined for it."""
    try:
        raw = SESSION_MODE_FILE.read_text(encoding="utf-8").strip()
    except (OSError, FileNotFoundError):
        return "general"
    return raw if raw in MODE_INFO else "general"


def save_session_mode(mode: str) -> None:
    SESSION_MODE_FILE.write_text(mode, encoding="utf-8")


@beta_tool
def enter_mode(mode: str) -> str:
    """Switch into a specific domain mode for subsequent turns, loading
    that mode's tools and rules on top of your always-on base tools.
    Available modes: 'wztc' (workzone traffic control -- sign placement,
    spacing/taper calculations, MUTCD/NYSDOT lookups). Call this when the
    engineer clearly wants to start that kind of task (e.g. "I need to
    develop a WZTC plan") -- not for every WZTC-adjacent mention, same
    selectivity spirit as ask_user_choice. The mode change takes effect
    starting next turn (this turn's tools are already fixed)."""
    if mode not in MODE_INFO:
        return f"Unknown mode {mode!r}. Available: {', '.join(MODE_INFO)}."
    _SESSION.mode = mode
    save_session_mode(mode)
    LOG.mode_changed(mode, MODE_INFO[mode])
    return f"Switched to {mode} mode."


@beta_tool
def exit_mode() -> str:
    """Return to general MicroStation mode, dropping the current domain
    mode's tools and any task-specific assumptions that came with it.
    Call this when the engineer's current task is done or they clearly
    move to something unrelated. Takes effect starting next turn."""
    _SESSION.mode = "general"
    save_session_mode("general")
    wztc_ops.reset_plan_session_flags()
    LOG.mode_changed("general", MODE_INFO["general"])
    return "Switched to general mode."


# BASE_TOOLS: always loaded, any mode. WZTC_TOOLS: only loaded once
# enter_mode("wztc") has been called. _MODE_TOOLS is what run_turn() reads
# each turn off _SESSION.mode — see the "Session modes" plan for why
# tools/rules are split this way instead of one flat always-on set.
BASE_TOOLS = [_wrap_op(name, getattr(wztc_ops, name)) for name in _BASE_OP_NAMES]
BASE_TOOLS.append(ask_user)
BASE_TOOLS.append(ask_user_choice)
BASE_TOOLS.append(view_drawing)
BASE_TOOLS.append(enter_mode)
BASE_TOOLS.append(exit_mode)

# Server-side tool (Anthropic-hosted -- no Python function to implement).
# allowed_domains hard-restricts this to MicroStation/VBA developer
# resources only, not general internet access -- see prompts.py for when
# this is (and is not) appropriate to reach for. max_uses caps the worst
# case within a single turn at 3 calls.
#
# Deliberately the BASIC variant (web_search_20250305), not the dynamic-
# filtering web_search_20260209 -- measured live against the same query,
# the dynamic-filtering variant routes results through a code_execution
# pass that more than doubled the real cost ($0.15 vs $0.065) for no
# quality difference on an already-narrow 3-domain allowlist; dynamic
# filtering earns its overhead on broad unrestricted searches, not this.
#
# stackoverflow.com deliberately NOT in allowed_domains -- the API
# rejects it outright with 400 "not accessible to our user agent"
# (confirmed live, not guessed): Stack Overflow/Stack Exchange blocks
# Anthropic's crawler, so it can never actually return results here
# regardless of intent.
BASE_TOOLS.append({
    "type": "web_search_20250305",
    "name": "web_search",
    "max_uses": 3,
    "allowed_domains": [
        "docs.bentley.com",
        "communities.bentley.com",
        "bentleysystems.service-now.com",
    ],
})

# WZTC_TOOLS: search_reference_manual is WZTC-only (MUTCD/NYS Supplement/
# Standard Sheets lookups), unlike the domain-agnostic web_search above.
WZTC_TOOLS = [_wrap_op(name, getattr(wztc_ops, name)) for name in _WZTC_OP_NAMES]
WZTC_TOOLS.append(_wrap_op("search_reference_manual", manual_search.search))

_MODE_TOOLS = {
    "general": BASE_TOOLS,
    "wztc": BASE_TOOLS + WZTC_TOOLS,
}

# Dropped from the prefix while a named 619 sheet plan is active (order
# table locked). Cuts unused highway-catalog / cell-browse / registry
# schemas. Keep place_two_way_highway — real-road finish still needs it.
_PLAN_OMIT_TOOL_NAMES = frozenset({
    "place_lane_highway",
    "place_divided_highway",
    "place_twlt_highway",
    "place_orthogonal_intersection",
    "place_ramp_gore",
    "cell_library_status",
    "attach_cell_library",
    "list_cells",
    "list_cell_libraries",
    "find_cell",
    "list_registry_commands",
    "describe_registry_command",
    "run_registry_command",
})


def _tool_registered_name(tool) -> str:
    if isinstance(tool, dict):
        return str(tool.get("name") or "")
    return str(getattr(tool, "__name__", "") or getattr(tool, "name", "") or "")


def tools_for_turn() -> list:
    """Tools for this API round-trip. Omit unused catalogs when a sheet plan is active."""
    tools = list(_MODE_TOOLS[_SESSION.mode])
    if _SESSION.mode == "wztc" and wztc_ops._PLAN_SESSION.sheet_plan_active():
        tools = [t for t in tools if _tool_registered_name(t) not in _PLAN_OMIT_TOOL_NAMES]
    return tools


def _auto_focus_and_capture() -> None:
    """After a turn touches any elements, pan/zoom the MicroStation view to
    show everything that changed (combined bounding box, 30% margin, 10ft
    floor so a single small sign doesn't zoom in absurdly tight) and take a
    screenshot, so the engineer watching the panel can actually see what
    happened instead of whatever was on screen before the turn started.
    Runs once per completed run_turn() call (which may span several
    ask_user rounds) -- the engineer's own choice over re-focusing after
    every individual tool call, which would jump the view around a lot on
    a multi-step turn. Best-effort: any failure here must never break the
    turn's real answer, so everything is swallowed and logged as a
    (non-fatal) ERROR line rather than raised."""
    if not _SESSION.touched_element_ids:
        return
    try:
        ids_csv = ",".join(sorted(_SESSION.touched_element_ids))
        resp = chat_bridge.call("GET_ELEMENTS_RANGE", elementIds=ids_csv)
        if resp.get("status") != "OK":
            return
        low_x, low_y = float(resp["lowX"]), float(resp["lowY"])
        high_x, high_y = float(resp["highX"]), float(resp["highY"])
        center_x = (low_x + high_x) / 2
        center_y = (low_y + high_y) / 2
        width = max(high_x - low_x, 10.0) * 1.3
        height = max(high_y - low_y, 10.0) * 1.3
        view_capture.navigate_view(center_x, center_y, width, height)
        _log_screenshot(view_capture.capture_microstation())
    except Exception as e:
        LOG.error(f"auto-focus/screenshot failed (non-fatal, turn result unaffected): {e}")


def run_turn(client: anthropic.Anthropic, messages: list[dict], user_text: str) -> str:
    """Run one full agentic turn (think -> act -> think -> ... -> final
    answer) for a single user message, mutating `messages` in place to
    mirror the full exchange (assistant turns + tool-result turns) so the
    NEXT call to this function has correct context. The tool_runner keeps
    its own internal copy of the conversation but does not expose it (per
    the SDK's own documented pause_turn-handling pattern) -- mirroring it
    ourselves via generate_tool_call_response() is the documented way to
    persist history across separate tool_runner() calls, one per turn."""
    _SESSION.touched_element_ids.clear()
    # Normalize SDK content blocks → dicts, then drop unanswered tool_use
    # chains before the API call (live 2026-08-05: repair missed SDK objects
    # and ERROR path left broken in-memory history stuck across "hi" retries).
    for m in messages:
        if isinstance(m, dict) and "content" in m:
            m["content"] = chat_history._to_jsonable(m["content"])
    # HARNESS_P0: repair then clear if still API-400-shaped — do not nudge
    # the model through a broken history loop.
    cleared = chat_history.harness_preflight_or_clear(messages)
    if cleared:
        LOG.error(
            "HARNESS_P0: cleared broken chat history before turn: "
            + "; ".join(cleared[:4])
        )
    chat_history.trim_cache_control(messages)
    messages.append({
        "role": "user",
        "content": [{"type": "text", "text": user_text, "cache_control": {"type": "ephemeral", "ttl": "1h"}}],
    })

    # _SESSION.mode is read fresh here on every call -- enter_mode/exit_mode
    # (called mid-turn, from inside the runner this function is about to
    # build) only take effect on the *next* run_turn() call, since this
    # turn's tool_runner is already under construction by the time a mode-
    # switch tool call could run. That's the intended granularity (a
    # deliberate, coarse switch, not a mid-turn one) -- see the "Session
    # modes" plan.
    runner = client.beta.messages.tool_runner(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        # cache_control on the system block caches tools + system together
        # -- render order is tools -> system -> messages, so a breakpoint
        # on the last (only) system block covers both. Without this, every
        # internal tool-calling round-trip inside a single turn re-sends
        # and re-bills the full system prompt + all TOOLS schemas at full
        # price -- and a turn can easily run 5-10+ of those round-trips
        # (e.g. a multi-query search_reference_manual lookup). Changing
        # _SESSION.mode between turns invalidates this cache on the turn
        # the switch happens (tools/system both changed) -- expected and
        # cheap for a deliberate, infrequent switch.
        system=[{"type": "text", "text": MODE_SYSTEM_PROMPT[_SESSION.mode], "cache_control": {"type": "ephemeral", "ttl": "1h"}}],
        thinking={"type": "adaptive", "display": "summarized"},
        output_config={"effort": EFFORT},
        tools=tools_for_turn(),
        messages=messages,
        # Clears stale tool_result content (search excerpts, journal dumps --
        # the actual bulky payloads, confirmed the dominant cost driver once
        # this conversation ran long) once it's no longer the newest few
        # turns. Leaves the tool_use calls themselves and all conversational
        # text alone, so the record of *what was done* survives -- only the
        # old *raw output* gets dropped. This is pruning, not summarizing;
        # see prompt-caching.md's distinction from compaction (a different
        # feature that summarizes instead of clearing, not used here).
        context_management={"edits": [{"type": "clear_tool_uses_20250919"}]},
        betas=["context-management-2025-06-27"],
        max_iterations=MAX_TOOL_ITERATIONS,
    )

    final_text = ""
    for message in runner:
        messages.append({"role": "assistant", "content": chat_history._to_jsonable(message.content)})
        if message.usage is not None:
            USAGE.record(message.usage, MODEL)

        for block in message.content:
            if getattr(block, "type", None) == "thinking" and getattr(block, "thinking", ""):
                LOG.thinking(block.thinking)
            elif (
                getattr(block, "type", None) == "text"
                and getattr(block, "text", "")
                and message.stop_reason != "end_turn"
            ):
                # A text block on the *final* message is the answer itself --
                # it goes out once, via LOG.final() below (main() calls it
                # with this function's return value). Logging it here too
                # would show the same answer twice in the transcript, once
                # styled as THINKING and once as FINAL (confirmed live).
                LOG.thinking(block.text)

        tool_response = runner.generate_tool_call_response()
        if tool_response is not None:
            # Persist as plain dicts so later repair/trim always sees tool ids.
            messages.append(chat_history._to_jsonable(tool_response))

        if message.stop_reason == "end_turn":
            final_text = "\n".join(b.text for b in message.content if getattr(b, "type", None) == "text")

    if final_text == "":
        # Loop ended without ever seeing end_turn -- either MAX_TOOL_ITERATIONS
        # was hit mid-investigation, or some other stop condition left no
        # final text block. Returning "" here would show a blank FINAL line
        # in the panel with no explanation; say so explicitly instead.
        final_text = (
            f"[Stopped after {MAX_TOOL_ITERATIONS} tool-call round-trips without "
            "reaching a final answer -- this is the MAX_TOOL_ITERATIONS safety "
            "cap, not a normal completion. On your NEXT turn: do NOT resume "
            "fishing (find_elements_near / Default linework / schema theories). "
            "Call get_locked_designer_inputs if needed, finish the next concrete "
            "step (assemble_corridor / place_* / ONE view_drawing), then FINAL "
            "with what is done vs remaining. Check the activity pane above for "
            "what was investigated so far.]"
        )

    return final_text


def main() -> None:
    client = anthropic.Anthropic()
    messages = chat_history.load_history(HISTORY_FILE)
    _SESSION.mode = load_session_mode()
    INPUT.skip_existing()
    hist_chars = sum(chat_history._content_char_len(m.get("content")) for m in messages
                     if isinstance(m, dict))
    print(f"chat_driver.py running -- model={MODEL}, effort={EFFORT}, mode={_SESSION.mode}, "
          f"max_iterations={MAX_TOOL_ITERATIONS}, history={len(messages)} msgs/~{hist_chars} chars "
          f"(caps {chat_history.MAX_HISTORY_MESSAGES}/{chat_history.MAX_HISTORY_CHARS}), "
          f"watching {CHAT_INPUT_FILE}, logging to {CHAT_LOG_FILE}", flush=True)

    restored = wztc_ops.try_restore_sheet_plan()
    if restored.get("loaded"):
        print(f"Restored sheet plan from {restored.get('persistedPath')} "
              f"(sheet={restored.get('sheetNum')}, updated={restored.get('updatedAt')})",
              flush=True)

    while True:
        user_text = INPUT.wait_for_next()
        try:
            final_text = run_turn(client, messages, user_text)
            chat_history.save_history(HISTORY_FILE, messages)
            _auto_focus_and_capture()
            LOG.final(final_text)
            print(f"[usage] session running total: ${USAGE.total_cost_usd:.4f} "
                  f"(see {USAGE_FILE} for the per-call breakdown)")
        except Exception as e:
            # Always repair+persist — otherwise a 400 from unpaired tool_use
            # leaves broken history in memory and every later "hi" 400s too
            # (live 2026-08-05).
            try:
                for m in messages:
                    if isinstance(m, dict) and "content" in m:
                        m["content"] = chat_history._to_jsonable(m["content"])
                chat_history._repair_tool_pairing(messages)
                chat_history.save_history(HISTORY_FILE, messages)
            except Exception as repair_err:
                LOG.error(f"history repair after turn failure also failed: {repair_err}")
            LOG.error(str(e))


if __name__ == "__main__":
    main()
