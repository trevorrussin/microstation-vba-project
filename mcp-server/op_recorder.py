"""Operation recorder for the QA gate (operation contracts, qa/operations.py).

Hooked into Bridge.call_batch (bridge_client.py). For every edit op the agent sends through the bridge it
appends one JSON line to outputs/qa/oplog.jsonl: the op, its exact parameters, the response, and the
geometry/properties of the operand BEFORE and of the operand + any new elements AFTER, read back from MicroStation
by this module, not reported by the agent. The gate (qa/checks/operations_check.py) validates each record
independently.

Never raises into the caller: recording is best-effort, and a failed recording shows up in the gate as an
unverifiable edit, never as a silent pass.
"""
from __future__ import annotations

import json
import os
import threading
import time
from pathlib import Path

LOG = Path(__file__).resolve().parents[1] / "outputs" / "qa" / "oplog.jsonl"
_lock = threading.Lock()

# ops with a geometric contract in qa/operations.py (operand-based edits and primitive placements)
CONTRACT_OPS = {
    "MOVE_ELEMENT", "COPY_ELEMENT", "COPY_PARALLEL", "ROTATE_ELEMENT", "SCALE_ELEMENT", "MIRROR_ELEMENT",
    "ARRAY_ELEMENT", "EXTEND_LINE", "BREAK_LINE", "CHANGE_ELEMENT_LEVEL", "CHANGE_ELEMENT_SYMBOLOGY",
    "EDIT_TEXT_ELEMENT", "DELETE_ELEMENT", "PLACE_POLYLINE", "PLACE_CIRCLE", "PLACE_ARC", "PLACE_BLOCK",
    "PLACE_TEXT_LABEL", "PLACE_CELL",
}
_ID_KEYS = ("newElementId", "newElementIds", "createdElementIds", "elementId")


def log_size() -> int:
    try:
        return LOG.stat().st_size
    except OSError:
        return 0


def _ids(resp: dict) -> list[str]:
    out: list[str] = []
    for k in ("newElementId", "newElementIds", "createdElementIds"):
        for part in str(resp.get(k) or "").split(","):
            part = part.strip()
            if part and part not in out:
                out.append(part)
    return out


def _read(bridge, eid: str) -> dict:
    """Geometry (raw vertices) + CAD props of one element, via the un-hooked batch call."""
    info: dict = {}
    try:
        r = bridge._call_batch_impl([("GET_ELEMENT_VERTICES", {"elementId": eid})])[0]
        if str(r.get("status")) == "OK":
            info["verts"] = [[float(x["x"]), float(x["y"]), float(x.get("z") or 0)] for x in r.get("rows") or []]
    except Exception:
        pass
    try:
        r = bridge._call_batch_impl([("GET_ELEMENT_CAD_PROPS", {"minId": 0, "ids": eid, "maxRows": 5})])[0]
        rows = [x for x in r.get("rows") or [] if x.get("elementId") != "TRUNCATED"]
        info["props"] = rows[0] if rows else None  # None = element no longer exists
    except Exception:
        info["props"] = "unreadable"
    return info


def begin(bridge, ops):
    edits = [(op, p) for op, p in ops if op in CONTRACT_OPS]
    if not edits:
        return {"before": {}, "n_edits": 0}
    before = {}
    for op, p in edits:
        eid = str(p.get("elementId") or "").strip()
        if eid and eid not in before:
            before[eid] = _read(bridge, eid)
    return {"before": before, "n_edits": len(edits)}


def finish(bridge, tok, ops, results):
    if tok is None:
        return
    lines = []
    for (op, params), resp in zip(ops, results):
        is_contract = op in CONTRACT_OPS
        new_ids = _ids(resp)
        if not new_ids and op.startswith("PLACE_") and resp.get("elementId"):
            new_ids = [str(resp["elementId"])]  # placement ops report the placed element as elementId
        if str(resp.get("status")) != "OK" or not (is_contract or new_ids):
            continue
        operand = str(params.get("elementId") or "").strip()
        after = {}
        for eid in ([operand] if operand else []) + [i for i in new_ids if i != operand]:
            after[eid] = _read(bridge, eid)
        keep = ("note", "priorDeltaX", "priorDeltaY", "priorAngleDeg", "priorLevel", "priorText")
        lines.append(json.dumps({
            "ts": time.time(), "op": op, "params": {k: v for k, v in params.items() if v is not None},
            "contract": is_contract, "new_ids": new_ids, "batch_edits": tok["n_edits"],
            "resp": {k: resp[k] for k in keep if k in resp},
            "before": {operand: tok["before"][operand]} if operand in tok["before"] else {},
            "after": after,
        }))
    if lines:
        with _lock:
            LOG.parent.mkdir(parents=True, exist_ok=True)
            with open(LOG, "a", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")


def read_since(offset: int) -> list[dict]:
    """Records appended after byte `offset` (the gate stores the offset at task start / turn start)."""
    if not LOG.exists():
        return []
    with open(LOG, "rb") as f:
        f.seek(max(int(offset or 0), 0))
        data = f.read().decode("utf-8", errors="replace")
    out = []
    for ln in data.splitlines():
        try:
            out.append(json.loads(ln))
        except json.JSONDecodeError:
            pass
    return out
