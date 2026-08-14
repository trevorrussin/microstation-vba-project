"""Post-placement scorecard: compile expectations vs placement registry.

Deterministic checks only — no image/OCR. Used to set geometry_qa_passed
after place_sheet_geometry and to gate visual_qa_passed.

Includes geometry-faithful checks (tip/mid drift, duplicate signs, kind flood)
via sheet_geometry_faithful.
"""
from __future__ import annotations

from collections import Counter
from typing import Any, Optional

import sheet_geometry_faithful as geom_faithful


# Layers the scorecard cares about (stations/signs tracked separately).
_COUNTED_KINDS = (
    "dimension",
    "label",
    "cone",
    "protectiveVehicle",
    "arrowPanel",
    "hatch",
    "transverseRun",
)


def _expected_from_compiled(compiled: dict) -> dict[str, Any]:
    """Count expected placeable primitives from compile_sheet_plan output."""
    plan = (compiled or {}).get("plan") or {}
    expected_ids: list[str] = []
    by_kind: Counter[str] = Counter()

    for a_str, prims in (plan.get("planByAlign") or {}).items():
        for p in prims or []:
            kind = p.get("kind")
            if kind in ("dimension", "label"):
                by_kind[kind] += 1
                if p.get("primitiveId"):
                    expected_ids.append(str(p["primitiveId"]))

    # Channelizing: one registry row per run (batched markers), not per cone.
    chan_runs: set[str] = set()
    for a_str, prims in (plan.get("channelizingByAlign") or {}).items():
        for p in prims or []:
            if p.get("kind") != "cone":
                continue
            run = str(p.get("run") or "run")
            pid = str(p.get("primitiveId") or f"{a_str}:{run}:cone")
            if pid not in chan_runs:
                chan_runs.add(pid)
                expected_ids.append(pid)
                by_kind["cone"] += 1

    for a_str, prims in (plan.get("symbolsByAlign") or {}).items():
        for p in prims or []:
            kind = p.get("kind")
            if kind in ("protectiveVehicle", "arrowPanel", "label"):
                by_kind[kind] += 1
                if p.get("primitiveId"):
                    expected_ids.append(str(p["primitiveId"]))

    for p in plan.get("hatch") or []:
        kind = p.get("kind")
        if kind in ("hatch", "transverseRun"):
            by_kind[kind] += 1
            if p.get("primitiveId"):
                expected_ids.append(str(p["primitiveId"]))

    # de-dupe expected ids preserving order
    seen: set[str] = set()
    uniq_ids: list[str] = []
    for i in expected_ids:
        if i not in seen:
            seen.add(i)
            uniq_ids.append(i)

    return {
        "byKind": dict(by_kind),
        "primitiveIds": uniq_ids,
        "counts": dict((compiled or {}).get("counts") or {}),
    }


def _placed_from_registry(registry_rows: list[dict]) -> dict[str, Any]:
    by_kind: Counter[str] = Counter()
    ids: list[str] = []
    empty_ids: list[str] = []
    for r in registry_rows:
        kind = str(r.get("kind") or "unknown")
        if kind == "sign":
            continue  # signs are order-table, not compiler scorecard
        by_kind[kind] += 1
        pid = str(r.get("primitiveId") or "")
        if pid:
            ids.append(pid)
        eids = r.get("elementIds") or []
        if not eids:
            empty_ids.append(pid or kind)
    return {
        "byKind": dict(by_kind),
        "primitiveIds": ids,
        "emptyElementIdRecords": empty_ids,
    }


def build_placement_scorecard(
    compiled: dict | None,
    registry_rows: list[dict] | None = None,
    executed: dict | None = None,
    gate_failures: list[str] | None = None,
    model_rows: list[dict] | None = None,
) -> dict:
    """Compare compiled plan expectations to registry heads.

    Returns:
      status: OK | FAIL
      passed: bool  (True only if no hard failures)
      failures: list[str]
      expected / placed summaries
      citations: sample primitiveId → elementIds / reqId for reflection
    """
    gates = list(gate_failures if gate_failures is not None
                 else (compiled or {}).get("gateFailures") or [])
    expected = _expected_from_compiled(compiled or {})
    placed = _placed_from_registry(registry_rows or [])
    exec_errors = list((executed or {}).get("errors") or [])

    failures: list[str] = []
    for g in gates:
        failures.append(f"compile-gate: {g}")
    for e in exec_errors:
        failures.append(f"execute: {e}")

    exp_ids = set(expected["primitiveIds"])
    got_ids = set(placed["primitiveIds"])
    # Only require coverage for kinds we expect from the compiler path.
    missing = sorted(exp_ids - got_ids)
    # Soft: extras are OK (signs, re-places) — don't fail on them.
    if missing:
        # Cap noise
        sample = missing[:12]
        more = len(missing) - len(sample)
        msg = f"scorecard: missing {len(missing)} registry primitiveId(s): {sample}"
        if more > 0:
            msg += f" (+{more} more)"
        failures.append(msg)

    for pid in placed.get("emptyElementIdRecords") or []:
        failures.append(
            f"scorecard: placed '{pid}' has empty elementIds "
            f"(bridge did not return createdElementIds)"
        )

    # Kind coverage: if we expected dims/labels/cones, require at least 1
    # of each expected kind that had count > 0.
    for kind, n in (expected.get("byKind") or {}).items():
        if n <= 0:
            continue
        got = int((placed.get("byKind") or {}).get(kind) or 0)
        if got == 0:
            failures.append(
                f"scorecard: expected kind={kind!r} count>={n} but registry has 0"
            )

    # Geometry-faithful layer (coords, duplicates, flood).
    failures.extend(
        geom_faithful.check_geometry_faithfulness(compiled, registry_rows)
    )
    failures.extend(geom_faithful.check_duplicate_signs(registry_rows))
    failures.extend(
        geom_faithful.check_kind_count_flood(
            expected.get("byKind"), placed.get("byKind"))
    )
    if model_rows:
        import build_overlap as ov
        for d in ov.tier1_duplicates(model_rows):
            t, cx, cy = d["key"][0], d["key"][1], d["key"][2]
            failures.append(
                f"scorecard: stacked {d['count']}x {t} at ({cx},{cy}) "
                f"ids={d['elementIds'][:6]}"
            )

    citations: list[dict] = []
    for r in (registry_rows or [])[:24]:
        citations.append({
            "primitiveId": r.get("primitiveId"),
            "kind": r.get("kind"),
            "elementIds": r.get("elementIds") or [],
            "reqId": r.get("reqId") or "",
            "specRef": r.get("specRef") or {},
            "bridgeOp": r.get("bridgeOp") or "",
            "midX": r.get("midX"),
            "midY": r.get("midY"),
            "x": r.get("x"),
            "y": r.get("y"),
        })

    passed = len(failures) == 0
    return {
        "status": "OK" if passed else "FAIL",
        "passed": passed,
        "failures": failures,
        "expected": expected,
        "placed": {
            "byKind": placed.get("byKind") or {},
            "primitiveIdCount": len(placed.get("primitiveIds") or []),
            "emptyElementIdRecords": placed.get("emptyElementIdRecords") or [],
        },
        "missingPrimitiveIds": missing[:40],
        "executeErrors": exec_errors,
        "compileGateFailures": gates,
        "citations": citations,
        "geometryFaithful": True,
    }


def visual_qa_prechecks(
    scorecard: Optional[dict],
    registry_rows: list[dict] | None = None,
    sheet_geometry_placed: bool = False,
    compiled: dict | None = None,
) -> list[str]:
    """Hard prerequisites before marking visual_qa_passed.

    Captures alone are not enough — scorecard must pass, registry must
    have compiler artifacts, and automated visual rules must pass.
    """
    fails: list[str] = []
    if not sheet_geometry_placed:
        fails.append("visual-qa: place_sheet_geometry has not succeeded")
    if scorecard is None:
        fails.append(
            "visual-qa: no scorecard — re-run place_sheet_geometry "
            "(or get_geometry_scorecard)"
        )
    elif not scorecard.get("passed"):
        n = len(scorecard.get("failures") or [])
        fails.append(
            f"visual-qa: geometry scorecard FAILED ({n} issue(s)) — "
            f"fix scorecard.failures before scripted captures count as pass"
        )
    rows = registry_rows or []
    compiler_kinds = {r.get("kind") for r in rows} & {
        "dimension", "label", "cone", "hatch", "protectiveVehicle", "arrowPanel"
    }
    if sheet_geometry_placed and not compiler_kinds:
        fails.append(
            "visual-qa: placement registry has no compiler kinds "
            "(dimension/label/cone/hatch/symbols) — rebuild geometry"
        )
    fails.extend(
        geom_faithful.check_automated_visual_rules(compiled, registry_rows)
    )
    return fails
