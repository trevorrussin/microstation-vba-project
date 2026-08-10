"""Geometry-faithful placement checks (compile expectations vs registry).

Extends presence-only scorecard with:
  - expected tip/mid/xy from compile vs registry.extra coords
  - duplicate roadside signs (same code near-identical XY)
  - kind-count under/over flood
  - re-surface compile-time Shapely geometry-qa strings already in gates
"""
from __future__ import annotations

import math
from collections import Counter, defaultdict
from typing import Any, Optional


_COORD_TOL_FT = 2.0  # tip/mid/label must land within this of compile intent
_DUP_SIGN_TOL_FT = 1.0
_KIND_OVER_RATIO = 2.5  # placed > expected * ratio → flood/duplicate fail


def _dist(ax: float, ay: float, bx: float, by: float) -> float:
    return math.hypot(ax - bx, ay - by)


def _mid_of_tips(tip1, tip2) -> Optional[tuple[float, float]]:
    try:
        return (
            0.5 * (float(tip1[0]) + float(tip2[0])),
            0.5 * (float(tip1[1]) + float(tip2[1])),
        )
    except (TypeError, ValueError, IndexError):
        return None


def _iter_compile_primitives(compiled: dict) -> list[dict]:
    """Flatten placeable compile primitives with expected geometry."""
    plan = (compiled or {}).get("plan") or {}
    out: list[dict] = []
    for a_str, prims in (plan.get("planByAlign") or {}).items():
        for p in prims or []:
            kind = p.get("kind")
            if kind == "dimension":
                mid = _mid_of_tips(p.get("tip1"), p.get("tip2"))
                out.append({
                    "primitiveId": str(p.get("primitiveId") or ""),
                    "kind": "dimension",
                    "alignIdx": int(a_str) if str(a_str).isdigit() else 0,
                    "text": p.get("text"),
                    "tip1": p.get("tip1"),
                    "tip2": p.get("tip2"),
                    "midX": mid[0] if mid else None,
                    "midY": mid[1] if mid else None,
                })
            elif kind == "label":
                out.append({
                    "primitiveId": str(p.get("primitiveId") or ""),
                    "kind": "label",
                    "alignIdx": int(a_str) if str(a_str).isdigit() else 0,
                    "text": p.get("text"),
                    "x": p.get("x"),
                    "y": p.get("y"),
                })
    for a_str, prims in (plan.get("symbolsByAlign") or {}).items():
        for p in prims or []:
            kind = p.get("kind")
            if kind not in ("protectiveVehicle", "arrowPanel", "label"):
                continue
            out.append({
                "primitiveId": str(p.get("primitiveId") or ""),
                "kind": kind,
                "alignIdx": int(a_str) if str(a_str).isdigit() else 0,
                "id": p.get("id"),
                "x": p.get("x"),
                "y": p.get("y"),
                "stationFt": p.get("stationFt"),
                "altGroup": p.get("altGroup"),
            })
    for a_str, prims in (plan.get("channelizingByAlign") or {}).items():
        # One expected run head per run id (batched markers).
        seen_runs: set[str] = set()
        for p in prims or []:
            if p.get("kind") != "cone":
                continue
            run = str(p.get("run") or "run")
            pid = str(p.get("primitiveId") or f"{a_str}:{run}:cone")
            if pid in seen_runs:
                continue
            seen_runs.add(pid)
            out.append({
                "primitiveId": pid,
                "kind": "cone",
                "alignIdx": int(a_str) if str(a_str).isdigit() else 0,
                "run": run,
                "x": p.get("x"),
                "y": p.get("y"),
                "stationFt": p.get("stationFt"),
            })
    return out


def check_geometry_faithfulness(
    compiled: dict | None,
    registry_rows: list[dict] | None = None,
    *,
    coord_tol_ft: float = _COORD_TOL_FT,
) -> list[str]:
    """Return hard failure strings when placed geometry drifts from compile."""
    fails: list[str] = []
    compiled = compiled or {}
    rows = registry_rows or []
    by_pid = {
        str(r.get("primitiveId") or ""): r
        for r in rows
        if r.get("primitiveId")
    }

    for exp in _iter_compile_primitives(compiled):
        pid = exp.get("primitiveId") or ""
        if not pid:
            continue
        got = by_pid.get(pid)
        if got is None:
            continue  # missing handled by presence scorecard
        kind = exp["kind"]
        # Dimension: registry.extra mid vs compile mid
        if kind == "dimension":
            emx, emy = exp.get("midX"), exp.get("midY")
            gmx, gmy = got.get("midX"), got.get("midY")
            if emx is None or gmx is None:
                # Fall back: if registry stored tip1/tip2
                gmid = _mid_of_tips(got.get("tip1"), got.get("tip2"))
                if gmid and emx is not None:
                    gmx, gmy = gmid
            if emx is not None and gmx is not None:
                d = _dist(float(emx), float(emy), float(gmx), float(gmy))
                if d > coord_tol_ft:
                    fails.append(
                        f"geometry-faithful: dimension '{pid}' mid drifted "
                        f"{d:.1f} ft from compile (tol {coord_tol_ft:g} ft)"
                    )
        elif kind in ("label", "protectiveVehicle", "arrowPanel"):
            ex, ey = exp.get("x"), exp.get("y")
            gx, gy = got.get("x"), got.get("y")
            if ex is None or gx is None:
                continue
            d = _dist(float(ex), float(ey), float(gx), float(gy))
            if d > coord_tol_ft:
                fails.append(
                    f"geometry-faithful: {kind} '{pid}' at "
                    f"({gx:.1f},{gy:.1f}) is {d:.1f} ft from compile "
                    f"({ex:.1f},{ey:.1f}) (tol {coord_tol_ft:g} ft)"
                )
        elif kind == "cone":
            # Batched run: first cone XY in extra vs compile first cone
            ex, ey = exp.get("x"), exp.get("y")
            gx, gy = got.get("x"), got.get("y")
            if ex is None or gx is None:
                continue
            d = _dist(float(ex), float(ey), float(gx), float(gy))
            if d > coord_tol_ft * 2:  # run start slightly looser
                fails.append(
                    f"geometry-faithful: channelizing run '{pid}' start "
                    f"drifted {d:.1f} ft from compile"
                )

    return fails


def check_duplicate_signs(
    registry_rows: list[dict] | None = None,
    *,
    tol_ft: float = _DUP_SIGN_TOL_FT,
) -> list[str]:
    """Fail when two sign assemblies share code + near-identical tip XY."""
    fails: list[str] = []
    signs = [
        r for r in (registry_rows or [])
        if str(r.get("kind") or "") == "sign"
    ]
    # Group by (alignIdx, normalized code)
    buckets: dict[tuple[int, str], list[dict]] = defaultdict(list)
    for r in signs:
        code = str(
            (r.get("specRef") or {}).get("signNum")
            or r.get("signNum")
            or r.get("primitiveId")
            or ""
        ).strip().upper()
        # primitiveId often "1:W20-05RA:sign"
        if ":sign" in code or code.count(":") >= 2:
            parts = str(r.get("primitiveId") or "").split(":")
            if len(parts) >= 2:
                code = parts[1].upper()
        align = int(r.get("alignIdx") or 0)
        buckets[(align, code)].append(r)

    for (align, code), group in buckets.items():
        if len(group) < 2 or not code:
            continue
        for i, a in enumerate(group):
            ax, ay = a.get("x"), a.get("y")
            if ax is None or ay is None:
                # Try tip from extra
                tip = a.get("tip") or a.get("pt1")
                if isinstance(tip, (list, tuple)) and len(tip) >= 2:
                    ax, ay = tip[0], tip[1]
            if ax is None:
                continue
            for b in group[i + 1:]:
                bx, by = b.get("x"), b.get("y")
                if bx is None:
                    tip = b.get("tip") or b.get("pt1")
                    if isinstance(tip, (list, tuple)) and len(tip) >= 2:
                        bx, by = tip[0], tip[1]
                if bx is None:
                    continue
                d = _dist(float(ax), float(ay), float(bx), float(by))
                if d <= tol_ft:
                    fails.append(
                        f"geometry-faithful: duplicate sign {code!r} on "
                        f"align {align} within {d:.2f} ft "
                        f"(primitiveIds "
                        f"{a.get('primitiveId')!r} / {b.get('primitiveId')!r})"
                    )
    return fails


def check_kind_count_flood(
    expected_by_kind: dict[str, int] | None,
    placed_by_kind: dict[str, int] | None,
    *,
    over_ratio: float = _KIND_OVER_RATIO,
) -> list[str]:
    fails: list[str] = []
    for kind, n in (expected_by_kind or {}).items():
        if n <= 0:
            continue
        got = int((placed_by_kind or {}).get(kind) or 0)
        if got > max(n * over_ratio, n + 2):
            fails.append(
                f"geometry-faithful: kind={kind!r} flood — expected ~{n}, "
                f"registry has {got} (likely duplicate place_sheet_geometry)"
            )
    return fails


def check_automated_visual_rules(
    compiled: dict | None,
    registry_rows: list[dict] | None = None,
) -> list[str]:
    """Deterministic 'visual' rules that must pass before LLM eyeballing.

    These replace soft model judgment for: duplicate signs, missing sheet
    labels that were compiled, PV/AP presence when compiled.
    """
    fails: list[str] = []
    fails.extend(check_duplicate_signs(registry_rows))
    compiled = compiled or {}
    plan = compiled.get("plan") or {}
    rows = registry_rows or []
    by_kind = Counter(str(r.get("kind") or "") for r in rows)

    # If compile emitted PV/AP, registry must have them (presence already
    # in scorecard; reinforce here for visual gate messaging).
    sym_kinds = set()
    for prims in (plan.get("symbolsByAlign") or {}).values():
        for p in prims or []:
            if p.get("kind") in ("protectiveVehicle", "arrowPanel"):
                sym_kinds.add(p["kind"])
    for k in sym_kinds:
        if by_kind.get(k, 0) < 1:
            fails.append(
                f"auto-visual: compiled {k} but registry has none — "
                f"not sheet-faithful"
            )

    # Label texts that compile emitted should appear (substring) in registry
    expected_labels = []
    for prims in (plan.get("planByAlign") or {}).values():
        for p in prims or []:
            if p.get("kind") == "label" and p.get("text"):
                expected_labels.append(str(p["text"]).strip().upper())
    for prims in (plan.get("symbolsByAlign") or {}).values():
        for p in prims or []:
            if p.get("kind") == "label" and p.get("text"):
                expected_labels.append(str(p["text"]).strip().upper())
    placed_label_text = " | ".join(
        str(r.get("text") or (r.get("specRef") or {}).get("text") or "").upper()
        for r in rows if r.get("kind") == "label"
    )
    for lab in expected_labels:
        if not lab:
            continue
        # Allow short match on first significant token
        token = lab.split()[0] if lab.split() else lab
        if token and token not in placed_label_text and lab not in placed_label_text:
            fails.append(
                f"auto-visual: expected label text containing {lab!r} "
                f"not found in registry labels"
            )

    return fails
