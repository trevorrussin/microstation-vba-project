"""Round-trip for 619-418 — Family 5 batch checker."""
from __future__ import annotations
import pathlib, runpy, sys
sys.exit(runpy.run_path(str(pathlib.Path(__file__).resolve().parent / '619-family5.py'), run_name='__main__') or 0)
