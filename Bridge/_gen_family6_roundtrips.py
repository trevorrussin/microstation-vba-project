"""Generate Bridge/roundtrip/619-NNN.py wrappers for Family 6."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "roundtrip"
SHEETS = [307, 308, 309, 314, 321, 322, 323, 324, 407, 421, 422, 519, 524, 90, 91]

for n in SHEETS:
    name = f"619-{n:03d}" if n < 100 else f"619-{n}"
    text = f'''"""Round-trip for {name}.json — Family 6."""
from __future__ import annotations

import importlib.util
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("f6", ROOT / "619-family6.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
fails = mod.check_sheet({n})
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
'''
    (OUT / f"{name}.py").write_text(text, encoding="utf-8")
    print("wrote", name)
