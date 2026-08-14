"""Append-only sheet-build ledger (survives journal rotation / registry wipe).

One JSONL row per run_sheet_build. Never truncated by clear_plan_elements.
Retention: keep the newest MAX_ROWS. No COM.
"""
from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

_ROOT = Path(__file__).resolve().parent.parent
LEDGER_PATH = _ROOT / "Bridge" / "build-ledger.jsonl"
MAX_ROWS = 40
MAX_PATH_PTS = 40


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _read_all() -> list[dict]:
    if not LEDGER_PATH.exists():
        return []
    rows: list[dict] = []
    for line in LEDGER_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return rows


def _write_all(rows: list[dict]) -> None:
    LEDGER_PATH.parent.mkdir(parents=True, exist_ok=True)
    LEDGER_PATH.write_text(
        "".join(json.dumps(r, separators=(",", ":")) + "\n" for r in rows),
        encoding="utf-8",
    )


def decimate_path(verts: list | None, max_pts: int = MAX_PATH_PTS) -> list[list[float]]:
    pts = []
    for v in verts or []:
        if not v or len(v) < 2:
            continue
        pts.append([float(v[0]), float(v[1])])
    if len(pts) <= max_pts:
        return pts
    step = (len(pts) - 1) / (max_pts - 1)
    out = []
    for i in range(max_pts):
        out.append(pts[int(round(i * step))])
    return out


def load_builds(sheet_num: str = "") -> list[dict]:
    rows = _read_all()
    sn = (sheet_num or "").strip()
    if sn:
        rows = [r for r in rows if str(r.get("sheetNum") or "") == sn]
    return rows


def append_build(
    *,
    sheet_num: str,
    origin: list[float],
    path_vertices: list | None = None,
    sta0: float = 0.0,
    sta1: float = 0.0,
    lateral_half_width: float = 40.0,
    element_id_min: str = "",
    element_id_max: str = "",
    bbox: Optional[dict] = None,
    extra: Optional[dict] = None,
) -> dict:
    rec: dict[str, Any] = {
        "buildId": uuid.uuid4().hex[:12],
        "ts": _now_iso(),
        "sheetNum": str(sheet_num or "").strip(),
        "origin": [float(origin[0]), float(origin[1])],
        "sta0": float(sta0),
        "sta1": float(sta1),
        "lateralHalfWidth": float(lateral_half_width),
        "elementIdMin": str(element_id_min or ""),
        "elementIdMax": str(element_id_max or ""),
        "bbox": bbox or {},
        "path_vertices": decimate_path(path_vertices),
    }
    if extra:
        rec["extra"] = extra
    rows = _read_all()
    rows.append(rec)
    if len(rows) > MAX_ROWS:
        rows = rows[-MAX_ROWS:]
    _write_all(rows)
    return rec
