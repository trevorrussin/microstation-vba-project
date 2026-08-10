"""Facade over the sheet-spec module split (2026-08-04).

This module used to hold all of it: spec loading/resolve, the placement-
plan compiler, and the pre-draw rules gate, in one ~1150-line file. It's now
split by concern:

  sheet_resolve.py -- "what does this sheet need" (load/resolve/legend,
                       order-table rows, station walk)
  sheet_compile.py -- "turn that into coordinates" (compile_plan/
                       compile_channelizing/compile_symbols/compile_hatch,
                       check_taper_continuity)
  sheet_rules.py    -- "validate before drawing" (check_corridor_topology,
                       run_rules_gate)

Every name that was previously `sheet_spec.X` is re-exported here unchanged,
so no existing call site (wztc_ops.py, scripts/test_compile_plan_parity.py,
etc.) needs to change. New code may import sheet_resolve/sheet_compile/
sheet_rules directly instead, but importing sheet_spec is still correct and
is not deprecated.
"""
from __future__ import annotations

from sheet_resolve import (
    SPEC_DIR,
    SpecError,
    allowed_speeds,
    shoulder_band,
    spec_path,
    has_spec,
    build_guide_path,
    load_build_guide,
    load,
    load_raw_path,
    legend_suffix,
    resolve,
    sign_library_key,
    zone_length,
    canonical_order_label,
    order_table_rows,
    station_walk,
)

from sheet_compile import (
    PERP_HALF_LEN_FT,
    annotation_style,
    channelizing_representation,
    compile_plan,
    compile_channelizing,
    check_taper_continuity,
    compile_symbols,
    compile_hatch,
)

from sheet_rules import (
    check_corridor_topology,
    compare_station_tables,
    run_rules_gate,
)

__all__ = [
    "SPEC_DIR", "SpecError", "allowed_speeds", "shoulder_band", "spec_path",
    "has_spec", "load", "load_raw_path", "legend_suffix", "resolve", "sign_library_key",
    "zone_length", "canonical_order_label", "order_table_rows", "station_walk",
    "PERP_HALF_LEN_FT", "annotation_style", "channelizing_representation",
    "compile_plan", "compile_channelizing",
    "check_taper_continuity", "compile_symbols", "compile_hatch",
    "check_corridor_topology", "compare_station_tables", "run_rules_gate",
]
