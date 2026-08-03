"""Round-trip wrapper — delegates to 619-family5.py batch checker."""
from __future__ import annotations

import pathlib
import runpy
import sys

ROOT = pathlib.Path(__file__).resolve().parent
sys.exit(runpy.run_path(str(ROOT / "619-family5.py"), run_name="__main__") or 0)
