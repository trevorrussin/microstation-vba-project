"""Constrained sheet-build sandbox: try → score → keep / revert.

Builds on an offset Y band so the prior kept corridor is not destroyed.
Revert deletes only sandbox-era journal/registry elements.
"""
from __future__ import annotations

import json
import shutil
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

_ROOT = Path(__file__).resolve().parent.parent
_BRIDGE = _ROOT / "Bridge"
STATE_PATH = _BRIDGE / "sandbox-state.json"
CHECKPOINT_DIR = _BRIDGE / "sandbox-checkpoints"


def _now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _read_state() -> Optional[dict]:
    if not STATE_PATH.exists():
        return None
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _write_state(state: dict) -> None:
    _BRIDGE.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def clear_state() -> None:
    if STATE_PATH.exists():
        STATE_PATH.unlink()


def offset_edge(edge: list[float], dy_ft: float) -> list[float]:
    if not edge or len(edge) < 2:
        raise ValueError("edge must be [x, y, ...] with at least x,y")
    out = list(edge)
    out[1] = float(out[1]) + float(dy_ft)
    return out


def begin_sandbox(
    *,
    upstream_edge: list[float],
    downstream_edge: list[float],
    offset_y_ft: float = 2000.0,
    sheet_num: str = "",
    registry_path: Path | None = None,
    plan_path: Path | None = None,
) -> dict:
    """Start a sandbox band north of the reference edges.

    Does NOT clear the prior kept build. Copies registry + sheet-plan for
    diagnostics. Returns offset edges for assemble/run_sheet_build.
    """
    if offset_y_ft == 0:
        raise ValueError(
            "offset_y_ft must be non-zero — in-place rebuild is not a "
            "sandbox (would destroy the kept corridor). Use ~2000 ft."
        )
    up = offset_edge(upstream_edge, offset_y_ft)
    dn = offset_edge(downstream_edge, offset_y_ft)
    band_id = uuid.uuid4().hex[:10]
    CHECKPOINT_DIR.mkdir(parents=True, exist_ok=True)
    reg_src = registry_path or (_BRIDGE / "placement-registry.jsonl")
    plan_src = plan_path or (_BRIDGE / "sheet-plan.json")
    reg_cp = CHECKPOINT_DIR / f"{band_id}-registry.jsonl"
    plan_cp = CHECKPOINT_DIR / f"{band_id}-sheet-plan.json"
    if reg_src.exists():
        shutil.copy2(reg_src, reg_cp)
    if plan_src.exists():
        shutil.copy2(plan_src, plan_cp)

    state = {
        "status": "active",
        "bandId": band_id,
        "sheetNum": sheet_num or "",
        "offsetYFt": float(offset_y_ft),
        "referenceUpstream": list(upstream_edge),
        "referenceDownstream": list(downstream_edge),
        "sandboxUpstream": up,
        "sandboxDownstream": dn,
        "startedAt": _now(),
        "checkpointRegistry": str(reg_cp) if reg_cp.exists() else None,
        "checkpointPlan": str(plan_cp) if plan_cp.exists() else None,
        "registryLenAtStart": (
            len(reg_src.read_text(encoding="utf-8").splitlines())
            if reg_src.exists() else 0
        ),
    }
    _write_state(state)
    return {
        "status": "OK",
        "sandbox": state,
        "upstream_edge": up,
        "downstream_edge": dn,
        "note": (
            f"Sandbox band {band_id} active (Y +{offset_y_ft:g} ft). "
            "Call run_sheet_build with the returned edges (or "
            "run_sheet_build_sandbox). KEEP with keep_sheet_sandbox; "
            "REVERT with revert_sheet_sandbox (clears sandbox placements only)."
        ),
    }


def get_sandbox() -> dict:
    st = _read_state()
    if not st:
        return {"status": "OK", "active": False, "note": "No sandbox band."}
    return {
        "status": "OK",
        "active": st.get("status") == "active",
        "sandbox": st,
    }


def keep_sandbox() -> dict:
    st = _read_state()
    if not st or st.get("status") != "active":
        return {"status": "ERROR", "note": "No active sandbox to keep."}
    st["status"] = "kept"
    st["keptAt"] = _now()
    _write_state(st)
    return {
        "status": "OK",
        "sandbox": st,
        "note": (
            f"Sandbox band {st.get('bandId')} KEPT. Prior reference band "
            f"(offset {-float(st.get('offsetYFt') or 0):g} ft relative) "
            "is unchanged — clear it manually if obsolete."
        ),
    }


def mark_reverted(extra: Optional[dict] = None) -> dict:
    st = _read_state() or {}
    st["status"] = "reverted"
    st["revertedAt"] = _now()
    if extra:
        st["revertDetail"] = extra
    _write_state(st)
    return st
