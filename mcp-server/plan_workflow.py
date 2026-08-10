"""Deterministic checklist for NAMED 619 STANDARD-SHEET builds only.

SCOPE (critical): these gates apply only while a sheet plan is active
(build_wztc_order_table locked a sheet_num / order table this session).
General MicroStation work, one-offs (one_off=True), spacing questions,
edits, and anything outside a sheet build must remain freeform — the
agent still thinks for itself there.

When active: PlanSession + get_plan_status() are source of truth; tools
flip stage bits; out-of-order calls raise PLAN_GATE with nextStep /
missing / accepted (live 619-311 burn 2026-08-04: MAX_TOOL on exploratory QA).
"""
from __future__ import annotations

from typing import Any, Optional


def sheet_plan_active(session: Any) -> bool:
    """True only after build_wztc_order_table locked a sheet this session.
    All PLAN_GATE / anti-fish-after-compiler / scripted-QA refuses must
    check this first — never constrain general CAD tasks."""
    return bool(getattr(session, "order_table_built", False)
                and getattr(session, "designer_inputs", None) is not None)


# Ordered stages for a sheet-spec plan. Each id maps to a PlanSession bit
# (or derived predicate). During an active sheet build the agent should
# call get_plan_status() and follow nextTool.
PLAN_STAGES: list[dict[str, str]] = [
    {"id": "inputs_locked", "label": "Designer inputs locked"},
    {"id": "order_table_built", "label": "Order table built"},
    {"id": "corridor_ready", "label": "Corridor ready (Align1+2)"},
    {"id": "stations_placed", "label": "Stations placed (all required aligns)"},
    {"id": "signs_placed", "label": "Order-table signs placed"},
    {"id": "sign_attrs_applied", "label": "Sign attributes applied"},
    {"id": "compiler_placed", "label": "place_sheet_geometry succeeded"},
    {"id": "geometry_qa_passed", "label": "Post-placement scorecard (registry vs compile)"},
    {"id": "visual_qa_passed", "label": "Scripted visual QA (scorecard-gated)"},
]


def format_plan_gate(
        message: str,
        *,
        tool: str = "",
        current_step: str = "",
        missing: Optional[list[str]] = None,
        accepted: Optional[list[str]] = None,
        next_step: str = "",
        next_tool: str = "") -> str:
    """Structured refuse text the agent can follow without guessing."""
    lines = ["PLAN_GATE"]
    if tool:
        lines.append(f"tool: {tool}")
    lines.append(f"message: {message}")
    if current_step:
        lines.append(f"currentStep: {current_step}")
    if missing:
        lines.append("missing: " + ", ".join(str(m) for m in missing))
    if accepted:
        lines.append("accepted: " + ", ".join(str(a) for a in accepted))
    if next_tool:
        lines.append(f"nextTool: {next_tool}")
    if next_step:
        lines.append(f"nextStep: {next_step}")
    lines.append("hint: call get_plan_status() — do not invent a workaround")
    return "\n".join(lines)


def raise_plan_gate(
        message: str,
        *,
        tool: str = "",
        current_step: str = "",
        missing: Optional[list[str]] = None,
        accepted: Optional[list[str]] = None,
        next_step: str = "",
        next_tool: str = "") -> None:
    raise ValueError(format_plan_gate(
        message, tool=tool, current_step=current_step, missing=missing,
        accepted=accepted, next_step=next_step, next_tool=next_tool))


def stage_done(session: Any) -> dict[str, bool]:
    """Compute checklist booleans from PlanSession fields."""
    inputs_locked = session.designer_inputs is not None
    order = bool(session.order_table_built)
    req = set(session.required_aligns) if session.required_aligns else ({1, 2} if order else set())
    corridor = bool(req) and req <= set(session.aligns_ready)
    stations = bool(req) and req <= set(session.stations_placed_aligns)
    locked_signs = set(session.locked_sign_rows)
    signs = bool(locked_signs) and locked_signs <= set(session.signs_placed_rows)
    if order and not locked_signs:
        # Sheet with no roadside signs (rare) — treat as done once stations exist.
        signs = stations
    return {
        "inputs_locked": inputs_locked,
        "order_table_built": order,
        "corridor_ready": corridor,
        "stations_placed": stations,
        "signs_placed": signs,
        "sign_attrs_applied": bool(session.sign_attrs_applied) or (
            signs and not locked_signs),
        "compiler_placed": bool(session.sheet_geometry_placed),
        "geometry_qa_passed": bool(session.geometry_qa_passed),
        "visual_qa_passed": bool(session.visual_qa_passed),
    }


def first_incomplete(done: dict[str, bool]) -> Optional[str]:
    for s in PLAN_STAGES:
        if not done.get(s["id"]):
            return s["id"]
    return None


def next_action(session: Any, done: dict[str, bool]) -> dict[str, Any]:
    """Concrete nextTool / nextStep for the agent (no LLM inference)."""
    step = first_incomplete(done)
    if step is None:
        return {
            "currentStep": "complete",
            "nextTool": None,
            "nextStep": "FINAL — summarize plan; list any deferred handoffs",
            "remainingSigns": [],
            "stationsNeeded": [],
        }

    remaining_signs = sorted(
        f"align{a}:{c}" for a, c in (session.locked_sign_rows - session.signs_placed_rows)
    )
    req = set(session.required_aligns) if session.required_aligns else {1, 2}
    stations_needed = sorted(req - set(session.stations_placed_aligns))

    table = {
        "inputs_locked": (
            "ask_user_choice for speed/lane/shoulder/area_type/sheet, then "
            "build_wztc_order_table",
            "ask_user_choice",
        ),
        "order_table_built": (
            "build_wztc_order_table(...); show table to engineer before drawing",
            "build_wztc_order_table",
        ),
        "corridor_ready": (
            "ask_user_choice(allow_point_pick=True) for upstream + downstream "
            "WORK AREA edges, then run_sheet_build(upstream_edge, downstream_edge) "
            "(preferred executor) OR assemble_corridor alone",
            "run_sheet_build",
        ),
        "stations_placed": (
            f"run_sheet_build() to finish stations/signs/compiler, or "
            f"place_order_table_stations for align_idx in {stations_needed}",
            "run_sheet_build",
        ),
        "signs_placed": (
            f"run_sheet_build() or place_sign + set_sign_attributes for: "
            f"{remaining_signs or '(none)'}",
            "run_sheet_build",
        ),
        "sign_attrs_applied": (
            "run_sheet_build() or set_sign_attributes on place_sign IDs",
            "run_sheet_build",
        ),
        "compiler_placed": (
            "run_sheet_build() or place_sheet_geometry(dry_run=True then False)",
            "run_sheet_build",
        ),
        "geometry_qa_passed": (
            "get_geometry_scorecard / reflect_sheet_build — fix scorecard.failures, "
            "then re-run place_sheet_geometry (force=True only if engineer accepts)",
            "get_geometry_scorecard",
        ),
        "visual_qa_passed": (
            "run_visual_qa_captures() after scorecard passes — scripted frames; "
            "do NOT free-pan with adjust_view",
            "run_visual_qa_captures",
        ),
    }
    next_step, next_tool = table.get(step, ("call get_plan_status", "get_plan_status"))
    return {
        "currentStep": step,
        "nextTool": next_tool,
        "nextStep": next_step,
        "remainingSigns": remaining_signs,
        "stationsNeeded": stations_needed,
    }


def build_status_dict(session: Any) -> dict[str, Any]:
    done = stage_done(session)
    action = next_action(session, done)
    checklist = [
        {"id": s["id"], "label": s["label"], "done": bool(done.get(s["id"]))}
        for s in PLAN_STAGES
    ]
    locked = session.get_locked_inputs_dict()
    return {
        "status": "OK",
        "checklist": checklist,
        "allComplete": action["currentStep"] == "complete",
        **action,
        "lockedInputs": locked,
        "alignsReady": sorted(session.aligns_ready),
        "stationsPlacedAligns": sorted(session.stations_placed_aligns),
        "requiredAligns": sorted(session.required_aligns) if session.required_aligns else [],
        "signsPlaced": sorted(f"align{a}:{c}" for a, c in session.signs_placed_rows),
        "sheetGeometryPlaced": bool(session.sheet_geometry_placed),
        "visualQaPassed": bool(session.visual_qa_passed),
        "geometryQaPassed": bool(session.geometry_qa_passed),
        "lastFailedPhase": getattr(session, "last_failed_phase", "") or None,
        "scorecardPassed": (
            None if getattr(session, "last_scorecard", None) is None
            else bool((session.last_scorecard or {}).get("passed"))
        ),
    }
