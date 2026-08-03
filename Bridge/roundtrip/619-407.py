"""Round-trip for 619-407.json — Family 6."""
from __future__ import annotations

import importlib.util
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("f6", ROOT / "619-family6.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
fails = mod.check_sheet(407)
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
