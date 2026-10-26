"""Ask the capability catalog: python scripts/capability.py "create a profile from PVIs" [--mode development] [-k 5]"""
from __future__ import annotations

import argparse
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import capabilities  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("request")
    ap.add_argument("--mode", choices=["production", "development"], default="production")
    ap.add_argument("-k", type=int, default=5)
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    r = capabilities.find(a.request, a.k, a.mode)
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(r, indent=1) if a.json else capabilities.format_result(r))
    return 0 if r["status"] == "FOUND" else 3


if __name__ == "__main__":
    sys.exit(main())
