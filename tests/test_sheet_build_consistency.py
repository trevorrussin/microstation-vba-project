"""The agent's tool path must produce the reference 619-311 build every time.

Regression net for 2026-08-20: the agent's L-bend build diverged from the
known-good reference (fresh L band, work bay ~(92570, 300000)) because of
three defects in the tool layer — none of them model judgment:

1. lock_corridor_path(source=last_placed) locked the road's OUTER edge as the
   alignment. Every known-good build offsets to the closed-lane cone line
   (CHAN_OFF = 2*lane + gap + lane = 38 ft for a 4-lane right closure), so
   the whole plan built 38 ft off the correct lateral band.
2. resolve_sheet_lateral took shoulder width from stale designer inputs
   (">= 8 ft" -> 8.0) instead of the actual road (0.0 here), and treated a
   measured 0.0 shoulder as "missing" via or-style fallbacks. half_len came
   out 20 instead of 12.
3. Designer inputs restored from a 5-hour-old sheet-plan.json counted as
   "locked", so a brand-new build silently reused them. The engineer's
   directive: a new build ALWAYS re-asks the designer questions.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import corridor_path as cp  # noqa: E402
import wztc_ops as ops  # noqa: E402


def _fresh_session():
    ops._PLAN_SESSION.reset()
    ops._LAST_PLACED_ROAD = None


def _place_road(shoulder: float = 0.0, side: str = "right", lanes: int = 4):
    ops._remember_placed_road(
        road_type="two_way_undivided", lanes=lanes, lane_width_ft=12,
        shoulder_width_ft=shoulder, yellow_gap_ft=2, side=side,
        verts=[[0, 0], [6000, 0]], x1=0, y1=0, x2=6000, y2=0, length=6000,
    )


def _lock_311_inputs(shoulder_band: str = ">= 8 ft"):
    ops._PLAN_SESSION.lock_designer_inputs(
        sheet_num="619-311", speed=55, road_type="Non-Freeway",
        lane_width=12, shoulder_width=shoulder_band, area_type="URBAN",
        closure_type="", exposure_condition="", protective_vehicle_gvw=0)


# ---------------------------------------------------------------- Fix 1: align offset


def test_locked_corridor_is_the_cone_line_not_the_road_edge():
    _fresh_session()
    _place_road()
    _lock_311_inputs()
    out = ops.lock_corridor_path("last_placed")
    assert out["status"] == "OK"
    assert abs(out["alignOffsetFt"] - 38.0) < 1e-6, "4-lane right closure: 2L+gap+L"
    assert out["edgeRole"] == "closed_lane_edge"
    # +X travel, side='right': cone line is 38 ft right of travel = y -38.
    ys = [v[1] for v in out["path_vertices"]]
    assert all(abs(y - (-38.0)) < 1e-6 for y in ys)
    _fresh_session()


def test_left_side_road_offsets_the_other_way():
    _fresh_session()
    _place_road(side="left")
    _lock_311_inputs()
    out = ops.lock_corridor_path("last_placed")
    assert abs(out["alignOffsetFt"] - (-38.0)) < 1e-6
    ys = [v[1] for v in out["path_vertices"]]
    assert all(abs(y - 38.0) < 1e-6 for y in ys)
    _fresh_session()


def test_six_lane_offset_scales():
    """3 lanes/dir right closure: 3L + gap + 2L = 36+2+24 = 62."""
    _fresh_session()
    _place_road(lanes=6)
    _lock_311_inputs()
    out = ops.lock_corridor_path("last_placed")
    assert abs(out["alignOffsetFt"] - 62.0) < 1e-6
    _fresh_session()


def test_no_derivable_closed_side_leaves_path_and_warns():
    """Without a sheet, do not guess an offset — warn instead."""
    _fresh_session()
    _place_road()
    out = ops.lock_corridor_path("last_placed")
    assert out["alignOffsetFt"] == 0.0
    assert out["edgeRole"] == "first_travel_outer"
    assert "closed side" in (out["alignOffsetNote"] or "")
    _fresh_session()


# ---------------------------------------------------------------- Fix 2: road-fact lateral


def test_half_len_uses_the_actual_road_shoulder_not_the_band():
    """Road has NO shoulder; designer band says >= 8 ft. Road wins: 12, not 20."""
    _fresh_session()
    _place_road(shoulder=0.0)
    _lock_311_inputs(">= 8 ft")
    ops.lock_corridor_path("last_placed")
    lat = ops.resolve_sheet_lateral([1000, -38, 0], [1300, -38, 0], "right")
    assert lat["lateralSource"] == "locked_road"
    assert abs(lat["half_len"] - 12.0) < 1e-6
    assert abs(lat["shoulder_width_ft"]) < 1e-6
    _fresh_session()


def test_half_len_matches_reference_when_road_has_shoulder():
    _fresh_session()
    _place_road(shoulder=8.0)
    _lock_311_inputs()
    ops.lock_corridor_path("last_placed")
    lat = ops.resolve_sheet_lateral([1000, -38, 0], [1300, -38, 0], "right")
    assert abs(lat["half_len"] - 20.0) < 1e-6, "lane 12 + shoulder 8 (reference)"
    _fresh_session()


def test_conflicting_caller_kwargs_are_overridden_and_reported():
    _fresh_session()
    _place_road(shoulder=0.0)
    _lock_311_inputs()
    ops.lock_corridor_path("last_placed")
    lat = ops.resolve_sheet_lateral(
        [1000, -38, 0], [1300, -38, 0], "right", shoulder_width_ft=8.0)
    assert abs(lat["half_len"] - 12.0) < 1e-6
    assert lat["overrodeCaller"] and "shoulder_width_ft" in lat["overrodeCaller"]
    _fresh_session()


def test_designer_inputs_still_used_when_no_road_facts():
    """Corridor from clicked points (no road) keeps the old fallback."""
    _fresh_session()
    _lock_311_inputs(">= 8 ft")
    lat = ops.resolve_sheet_lateral([1000, 0, 0], [1300, 0, 0], "right")
    assert lat["lateralSource"] == "designer_inputs"
    assert abs(lat["half_len"] - 20.0) < 1e-6
    _fresh_session()


# ---------------------------------------------------------------- Fix 3: per-build inputs


def _stale_plan_payload(updated_at: str) -> dict:
    return {
        "schemaVersion": "1",
        "updatedAt": updated_at,
        "sheetNum": "619-311",
        "designerInputs": {
            "sheet_num": "619-311", "speed": 55, "road_type": "Non-Freeway",
            "lane_width": 12, "shoulder_width": ">= 8 ft", "area_type": "URBAN",
            "closure_type": "", "exposure_condition": "",
            "protective_vehicle_gvw": 0,
        },
        "checklist": {"inputs_locked": True, "order_table_built": True},
        "order_table_built": True,
    }


def test_stale_restored_plan_requires_reconfirm(tmp_path):
    _fresh_session()
    old = (datetime.now(timezone.utc) - timedelta(hours=5)).isoformat()
    p = tmp_path / "sheet-plan.json"
    p.write_text(json.dumps(_stale_plan_payload(old)), encoding="utf-8")
    r = ops._load_sheet_plan(p)
    assert r["loaded"] is True
    locked = ops.get_locked_designer_inputs()
    assert locked["locked"] is False
    assert locked["needsConfirm"] is True
    assert locked["previous"]["speed"] == 55, "previous values offered as defaults"
    _fresh_session()


def test_fresh_restored_plan_is_a_resume_not_a_new_build(tmp_path):
    """Driver restart minutes into a build must NOT force a re-ask
    (2026-08-04 lesson: never re-ask mid-build)."""
    _fresh_session()
    recent = (datetime.now(timezone.utc) - timedelta(minutes=5)).isoformat()
    p = tmp_path / "sheet-plan.json"
    p.write_text(json.dumps(_stale_plan_payload(recent)), encoding="utf-8")
    ops._load_sheet_plan(p)
    locked = ops.get_locked_designer_inputs()
    assert locked["locked"] is True
    assert "needsConfirm" not in locked
    _fresh_session()


def test_relocking_inputs_clears_the_confirm_flag():
    """build_wztc_order_table's lock IS the confirmation."""
    _fresh_session()
    ops._PLAN_SESSION.inputs_need_confirm = True
    _lock_311_inputs()
    assert ops._PLAN_SESSION.inputs_need_confirm is False
    assert ops.get_locked_designer_inputs()["locked"] is True
    _fresh_session()


def test_prompt_states_locks_are_per_build():
    import prompts

    text = prompts.WZTC_SYSTEM_PROMPT_ADDENDUM
    assert "LOCKS ARE PER-BUILD" in text
    assert "needsConfirm" in text


# ---------------------------------------------------------------- Fix 4: guide
# cleanup blinding visual QA


def test_visual_qa_survives_guide_deletion(monkeypatch, tmp_path):
    """The exact 2026-08-20 miss: scorecard passed (22 elements), but
    visual_qa failed with "could not read alignment vertices for framing"
    because delete_construction_guides had just removed the live alignment
    elements run_visual_qa_captures reads from. Caching the bbox before
    cleanup must make QA succeed on the cached points.

    Isolates SHEET_PLAN_PATH -- run_visual_qa_captures's success path saves
    the plan, and an earlier version of this test clobbered the real
    project's Bridge/sheet-plan.json with fixture values (live 2026-08-20).
    """
    monkeypatch.setattr(ops, "SHEET_PLAN_PATH", tmp_path / "sheet-plan.json")
    _fresh_session()
    _lock_311_inputs()
    ops._PLAN_SESSION.order_table_built = True
    ops._PLAN_SESSION.sheet_geometry_placed = True
    ops._PLAN_SESSION.last_scorecard = {"passed": True, "citations": []}
    ops._PLAN_SESSION.required_aligns = {1, 2}
    ops._PLAN_SESSION.last_alignment_bbox_pts = None

    # get_alignment_vertices raises like it does once the guide lines are gone.
    monkeypatch.setattr(ops, "get_alignment_vertices",
                        lambda a: (_ for _ in ()).throw(RuntimeError("gone")))
    monkeypatch.setattr(
        "sheet_scorecard.visual_qa_prechecks", lambda *a, **k: [])
    monkeypatch.setattr(
        "placement_registry.resolve_latest_placements", lambda **k: [])

    # Without the cache: the original bug reproduces exactly.
    with pytest.raises(ValueError, match="could not read alignment vertices"):
        ops.run_visual_qa_captures(force=True)

    # With the cache populated (as run_sheet_build now does before cleanup):
    ops._PLAN_SESSION.last_alignment_bbox_pts = [(0.0, 0.0), (6000.0, 0.0)]
    calls = []
    monkeypatch.setattr(ops, "adjust_view", lambda **k: calls.append(k))
    monkeypatch.setattr(ops, "capture_view", lambda: {"path": ""})
    r = ops.run_visual_qa_captures(force=True)
    assert r["status"] == "OK"
    assert len(calls) == 4, "four QA frames must still get framed"
    _fresh_session()


def test_bbox_cache_persists_across_save_load(tmp_path):
    _fresh_session()
    _lock_311_inputs()
    ops._PLAN_SESSION.order_table_built = True
    ops._PLAN_SESSION.last_alignment_bbox_pts = [(100.0, 200.0), (300.0, 400.0)]
    p = tmp_path / "sheet-plan.json"
    import wztc_ops as _o
    old_path = _o.SHEET_PLAN_PATH
    _o.SHEET_PLAN_PATH = p
    try:
        _o._save_sheet_plan()
        _fresh_session()
        r = _o._load_sheet_plan(p)
        assert r["loaded"] is True
        assert _o._PLAN_SESSION.last_alignment_bbox_pts == [[100.0, 200.0], [300.0, 400.0]]
    finally:
        _o.SHEET_PLAN_PATH = old_path
        _fresh_session()
