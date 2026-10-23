"""Harness sheet state machine — on-disk authority for curb classify progress.

States: NOT_STARTED | IN_PROGRESS | VERIFYING | PASSED | BLOCKED | ESCALATED

Agent must not advance to NEXT sheet while any sheet is BLOCKED.
"""
from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Any

OUT_DIR = r"C:\repos\microstation-vba-project\outputs"
HARNESS_MANIFEST = os.path.join(OUT_DIR, "engineering-harness-manifest.json")

VALID = {
    "NOT_STARTED",
    "IN_PROGRESS",
    "VERIFYING",
    "PASSED",
    "BLOCKED",
    "ESCALATED",
}


def _empty() -> dict[str, Any]:
    return {
        "version": 1,
        "updated": datetime.now().isoformat(timespec="seconds"),
        "domain": "loc-curb-classify",
        "sandbox": r"C:\repos\microstation-vba-project\Screenshots",
        "sheets": {},
        "last_blocked": None,
        "notes": [],
    }


def load(path: str = HARNESS_MANIFEST) -> dict[str, Any]:
    if not os.path.exists(path):
        return _empty()
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("sheets", {})
    data.setdefault("last_blocked", None)
    data.setdefault("notes", [])
    return data


def save(data: dict[str, Any], path: str = HARNESS_MANIFEST) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data["updated"] = datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def set_sheet(
    data: dict[str, Any],
    file_name: str,
    status: str,
    *,
    detail: dict[str, Any] | None = None,
) -> None:
    if status not in VALID:
        raise ValueError(f"invalid status {status}")
    sheets = data.setdefault("sheets", {})
    row = sheets.get(file_name, {})
    row["status"] = status
    row["updated"] = datetime.now().isoformat(timespec="seconds")
    if detail:
        row["last"] = detail
    sheets[file_name] = row
    if status == "BLOCKED":
        data["last_blocked"] = {"file": file_name, "detail": detail}
    elif data.get("last_blocked", {}) and data["last_blocked"].get("file") == file_name:
        if status in ("PASSED", "NOT_STARTED"):
            # clear if this sheet unblocked; else point at any remaining BLOCKED
            other = None
            for name, srow in sheets.items():
                if name != file_name and srow.get("status") == "BLOCKED":
                    other = name
                    break
            data["last_blocked"] = (
                {"file": other, "detail": sheets.get(other, {}).get("last")}
                if other
                else None
            )


def any_blocked(data: dict[str, Any]) -> str | None:
    for name, row in data.get("sheets", {}).items():
        if row.get("status") == "BLOCKED":
            return name
    return None


def counts(data: dict[str, Any]) -> dict[str, int]:
    c = {s: 0 for s in VALID}
    for row in data.get("sheets", {}).values():
        st = row.get("status", "NOT_STARTED")
        c[st] = c.get(st, 0) + 1
    return c
