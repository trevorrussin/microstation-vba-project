"""Placement registry: link compiled / placed primitives to DGN element IDs.

Append-only JSONL at Bridge/placement-registry.jsonl for the active sheet
build. Truncated on clear_plan_elements / new build_wztc_order_table.
Authority is agent-placed geometry only — no auto-rebind after hand edits.

Latest-wins: re-placing the same primitiveId appends a new line with
supersedes=<prior record id>; resolve_latest_placements() returns only
current heads (not superseded / deleted).
"""
from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

_ROOT = Path(__file__).resolve().parent.parent
REGISTRY_PATH = _ROOT / "Bridge" / "placement-registry.jsonl"


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clear_registry() -> None:
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
    REGISTRY_PATH.write_text("", encoding="utf-8")


def _read_all() -> list[dict]:
    if not REGISTRY_PATH.exists():
        return []
    rows: list[dict] = []
    for line in REGISTRY_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return rows


def _write_all(rows: list[dict]) -> None:
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
    REGISTRY_PATH.write_text(
        "".join(json.dumps(r, separators=(",", ":")) + "\n" for r in rows),
        encoding="utf-8",
    )


def append_placement(
    *,
    sheet_num: str,
    align_idx: int,
    kind: str,
    primitive_id: str,
    bridge_op: str,
    element_ids: list[str],
    spec_ref: Optional[dict] = None,
    req_id: str = "",
    extra: Optional[dict] = None,
) -> dict:
    """Append one placement record. Same primitiveId re-place supersedes
    the prior head (latest-wins). Returns the record written."""
    prior_heads = {
        r["primitiveId"]: r
        for r in resolve_latest_placements(sheet_num=sheet_num)
        if r.get("primitiveId")
    }
    supersedes = ""
    if primitive_id and primitive_id in prior_heads:
        supersedes = str(prior_heads[primitive_id].get("recordId") or "")

    rec: dict[str, Any] = {
        "recordId": uuid.uuid4().hex[:12],
        "ts": _now_iso(),
        "sheetNum": sheet_num or "",
        "alignIdx": int(align_idx or 0),
        "kind": kind,
        "primitiveId": primitive_id,
        "specRef": spec_ref or {},
        "bridgeOp": bridge_op,
        "reqId": str(req_id or ""),
        "elementIds": [str(e) for e in element_ids if str(e).strip()],
        "supersedes": supersedes,
        "deleted": False,
    }
    if extra:
        # Don't let extras clobber provenance keys accidentally.
        for k, v in extra.items():
            if k not in ("recordId", "supersedes", "deleted"):
                rec[k] = v
    REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with REGISTRY_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, separators=(",", ":")) + "\n")
    return rec


def parse_created_ids(resp: dict | None) -> list[str]:
    """Extract element id strings from a bridge place_* response."""
    if not isinstance(resp, dict):
        return []
    out: list[str] = []
    for key in ("createdElementIds", "elementIds"):
        raw = resp.get(key)
        if raw is None:
            continue
        if isinstance(raw, list):
            out.extend(str(x).strip() for x in raw if str(x).strip())
        else:
            out.extend(p.strip() for p in str(raw).split(",") if p.strip())
    eid = resp.get("elementId")
    if eid is not None and str(eid).strip():
        out.append(str(eid).strip())
    seen: set[str] = set()
    uniq: list[str] = []
    for i in out:
        if i not in seen:
            seen.add(i)
            uniq.append(i)
    return uniq


def _matches(
    rec: dict,
    *,
    sheet_num: str = "",
    kind: str = "",
    zone: str = "",
    run: str = "",
    align_idx: int = 0,
) -> bool:
    if sheet_num and str(rec.get("sheetNum", "")) != str(sheet_num):
        return False
    if kind and str(rec.get("kind", "")).lower() != kind.lower():
        return False
    if align_idx and int(rec.get("alignIdx") or 0) != int(align_idx):
        return False
    spec = rec.get("specRef") or {}
    if zone and str(spec.get("zone") or "") != zone:
        return False
    if run and str(spec.get("run") or rec.get("run") or "") != run:
        return False
    return True


def load_placements(
    sheet_num: str = "",
    kind: str = "",
    zone: str = "",
    run: str = "",
    align_idx: int = 0,
    include_superseded: bool = False,
) -> list[dict]:
    """Filter registry records. By default returns latest-wins heads only."""
    if include_superseded:
        return [
            r for r in _read_all()
            if not r.get("deleted")
            and _matches(r, sheet_num=sheet_num, kind=kind, zone=zone,
                         run=run, align_idx=align_idx)
        ]
    return resolve_latest_placements(
        sheet_num=sheet_num, kind=kind, zone=zone, run=run, align_idx=align_idx)


def resolve_latest_placements(
    sheet_num: str = "",
    kind: str = "",
    zone: str = "",
    run: str = "",
    align_idx: int = 0,
) -> list[dict]:
    """Latest non-deleted record per primitiveId (supersedes chain)."""
    rows = _read_all()
    superseded: set[str] = set()
    for r in rows:
        sid = str(r.get("supersedes") or "")
        if sid:
            superseded.add(sid)
    heads: dict[str, dict] = {}
    order: list[str] = []
    for r in rows:
        if r.get("deleted"):
            continue
        rid = str(r.get("recordId") or "")
        if rid and rid in superseded:
            continue
        if not _matches(r, sheet_num=sheet_num, kind=kind, zone=zone,
                        run=run, align_idx=align_idx):
            continue
        pid = str(r.get("primitiveId") or "") or rid or f"_anon_{len(order)}"
        if pid not in heads:
            order.append(pid)
        heads[pid] = r  # later line wins for same primitiveId
    return [heads[p] for p in order if p in heads]


def mark_deleted(primitive_ids: set[str]) -> int:
    """Mark latest heads for primitiveIds as deleted (soft). Returns count."""
    if not primitive_ids:
        return 0
    target_rids = {
        r.get("recordId")
        for r in resolve_latest_placements()
        if r.get("primitiveId") in primitive_ids and r.get("recordId")
    }
    if not target_rids:
        return 0
    rows = _read_all()
    removed = 0
    for r in rows:
        if r.get("recordId") in target_rids and not r.get("deleted"):
            r["deleted"] = True
            removed += 1
    if removed:
        _write_all(rows)
    return removed


def group_by_kind(rows: Optional[list[dict]] = None) -> dict[str, list[dict]]:
    rows = rows if rows is not None else resolve_latest_placements()
    out: dict[str, list[dict]] = {}
    for r in rows:
        out.setdefault(str(r.get("kind") or "unknown"), []).append(r)
    return out
