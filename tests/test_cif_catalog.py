"""Offline checks for Civil/CIF catalog + named wrappers (no ORD required)."""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))


def test_catalog_loads_expected_ops():
    import cif_catalog
    cif_catalog.load_catalog(force=True)
    names = {r["name"] for r in cif_catalog.list_ops()}
    for required in (
        "list_alignments",
        "get_alignment_geometry",
        "list_profiles",
        "get_profile_pvis",
        "create_profile",
        "set_active_profile",
        "list_corridors",
        "get_corridor_properties",
        "add_template_drop",
        "list_terrains",
        "get_active_terrain",
        "set_active_terrain",
        "list_templates",
    ):
        assert required in names, required
    blocked = cif_catalog.get_op("set_active_terrain")
    assert blocked["kind"] == "engineer"
    assert blocked["cmd"] is None
    assert blocked["production_ok"] is False


def test_set_active_terrain_wrapper_handoff():
    import wztc_ops
    out = wztc_ops.set_active_terrain("EG")
    assert out["status"] == "ERROR"
    assert out.get("needsEngineer") is True
    assert "ActiveSurface" in out["note"] or "get-only" in out["note"]


def test_chat_driver_registers_civil_names():
    # Importing chat_driver validates _BASE_OP_NAMES against wztc_ops.
    # Avoid starting watchers: only load the name lists via a light check.
    import wztc_ops
    names = [
        "list_cif_ops", "list_alignments", "get_alignment_geometry", "list_profiles",
        "get_profile_pvis", "create_profile", "set_active_profile",
        "list_corridors", "get_corridor_properties", "add_template_drop",
        "list_terrains", "get_active_terrain", "set_active_terrain", "list_templates",
    ]
    missing = [n for n in names if not hasattr(wztc_ops, n)]
    assert missing == [], missing


def test_switch_cmds_documented_in_cs():
    civil = (ROOT / "ord-addin" / "src" / "CifOpsCivil.cs").read_text(encoding="utf-8")
    ops = (ROOT / "ord-addin" / "src" / "CifOps.cs").read_text(encoding="utf-8")
    for cmd in (
        "get_corridor_properties",
        "set_active_profile",
        "get_active_terrain",
        "add_template_drop",
    ):
        assert f'case "{cmd}"' in ops, cmd
    assert "AddTemplateDrop" in civil
    assert "SetActiveProfile" in civil
    assert "ActiveSurface" in civil
    assert "get_corridor_properties" in (ROOT / "Data" / "cif-ops.yaml").read_text(encoding="utf-8")
