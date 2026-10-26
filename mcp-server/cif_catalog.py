"""CIF op catalog (Data/cif-ops.yaml): schemas, ORD version, verify hints."""
from __future__ import annotations

from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_CATALOG = _ROOT / "Data" / "cif-ops.yaml"
_cache: dict | None = None


def load_catalog(force: bool = False) -> dict:
    global _cache
    if _cache is not None and not force:
        return _cache
    import yaml
    _cache = yaml.safe_load(_CATALOG.read_text(encoding="utf-8")) or {}
    return _cache


def list_ops() -> list[dict]:
    cat = load_catalog()
    rows = []
    for name, spec in (cat.get("ops") or {}).items():
        rows.append({
            "name": name,
            "kind": spec.get("kind"),
            "cmd": spec.get("cmd"),
            "inputs": spec.get("inputs") or {},
            "verify": spec.get("verify") or "",
            "production_ok": bool(spec.get("production_ok")),
            "note": spec.get("note") or "",
            "ord_version": cat.get("version"),
            "sdk": cat.get("sdk"),
        })
    return rows


def get_op(name: str) -> dict | None:
    for row in list_ops():
        if row["name"] == name:
            return row
    return None


def verify_hint(name: str) -> str:
    op = get_op(name)
    return (op or {}).get("verify") or ""
