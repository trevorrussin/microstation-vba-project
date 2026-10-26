"""Pre/post-edit change set and drawing-state fingerprint for the QA gate.

The element-id baseline only sees elements created after `start`. This module snapshots every element in the
active model (the same read-back rows the checks use, plus a hash of each linear element's vertices) and diffs a
later read, so modified and deleted existing elements are caught too, including vertex edits that leave the
range box unchanged.

`fingerprint()` is one hash of that whole state. The gate stores it with a PASS and `close` recomputes it: the
artifact released must be the artifact that was verified.
"""
from __future__ import annotations

import hashlib
import json

FIELDS = ("type", "level", "color", "weight", "lineStyle", "cellName", "textHeight", "textWidth", "textFont",
          "textRotDeg", "text", "rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY", "zLow", "zHigh", "childLevels")
LINEAR = ("line", "curve", "arc", "shape", "string", "complex", "other")
VH_MAX = 1500  # vertex hashes are read for at most this many elements per state read (~35 ms each)


def is_linear(row: dict) -> bool:
    t = (row.get("type") or "").lower()
    return not any(k in t for k in ("text", "cell", "dimension")) and any(k in t for k in LINEAR)


def vhash(points) -> str:
    """Hash of a vertex list, rounded to 1e-4 ft so read-back noise does not register as an edit."""
    flat = [[round(float(c), 4) for c in p[:3]] for p in points or []]
    return hashlib.sha1(json.dumps(flat).encode()).hexdigest()[:16]


def snapshot(rows: list[dict], vh: dict | None = None) -> dict:
    snap = {str(r["elementId"]): {k: str(r.get(k) or "") for k in FIELDS} for r in rows if r.get("elementId")}
    for eid, h in (vh or {}).items():
        if eid in snap:
            snap[eid]["vh"] = h
    return snap


def diff(before: dict, rows_after: list[dict], vh_after: dict | None = None) -> dict:
    """created/modified rows are the current-state read-back rows (so the normal checks apply to them);
    modified_fields maps id -> changed field names; deleted maps id -> the old snapshot entry."""
    after = {str(r["elementId"]): r for r in rows_after if r.get("elementId")}
    created, modified, fields = [], [], {}
    for eid, row in after.items():
        old = before.get(eid)
        if old is None:
            created.append(row)
            continue
        changed = [k for k in FIELDS if str(row.get(k) or "") != old.get(k, "")]
        if vh_after is not None and old.get("vh") and eid in vh_after and vh_after[eid] != old["vh"]:
            changed.append("vertices")
        if changed:
            modified.append(row)
            fields[eid] = changed
    deleted = {eid: old for eid, old in before.items() if eid not in after}
    return {"created": created, "modified": modified, "modified_fields": fields, "deleted": deleted}


def fingerprint(rows: list[dict], vh: dict | None = None, model: str = "") -> str:
    """Canonical hash of the whole model state: every element's checked fields + vertex hash, plus the model name."""
    snap = snapshot(rows, vh)
    body = json.dumps([model, sorted(snap.items())], sort_keys=True)
    return hashlib.sha256(body.encode()).hexdigest()
