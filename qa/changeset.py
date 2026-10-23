"""Pre/post-edit change set for the QA gate.

The element-id baseline only sees elements created after `start`. This module snapshots every element in the
active model (the same read-back rows the checks use) and diffs a later read, so modified and deleted
existing elements are caught too.
"""
from __future__ import annotations

FIELDS = ("type", "level", "color", "weight", "lineStyle", "cellName", "textHeight", "textWidth", "textFont",
          "textRotDeg", "text", "rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY", "zLow", "zHigh")


def snapshot(rows: list[dict]) -> dict:
    return {str(r["elementId"]): {k: str(r.get(k) or "") for k in FIELDS} for r in rows if r.get("elementId")}


def diff(before: dict, rows_after: list[dict]) -> dict:
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
        if changed:
            modified.append(row)
            fields[eid] = changed
    deleted = {eid: old for eid, old in before.items() if eid not in after}
    return {"created": created, "modified": modified, "modified_fields": fields, "deleted": deleted}
