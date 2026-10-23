"""Condense a tutorial transcript for reading: drop ribbon/toolbar boilerplate and OCR fragments already seen.

  python scripts/tutorial_digest.py                 # every transcript -> <folder>/digest.md
  python scripts/tutorial_digest.py --only Profiles

Keeps every SAY line (speech) and, for SCREEN lines, only OCR segments not shown earlier in the same video, so
dialog titles, field names/values and status-bar prompts stand out. Timestamps are preserved for citation.
"""
from __future__ import annotations

import argparse
import glob
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DERIVED = os.path.join(ROOT, "Standards", "derived", "tutorials")
BOILER = re.compile(r"^(File|Home|Terrain|Geometry|Site|Corridors|Model Detailing|Drawing Production|Drawing|Utilities|Collaborate|View|Help|"
                    r"Primary|Selection|Placement|Manipulate|Modify|Groups|Tools|Attributes|Search Ribbon.*|Default.*none.*|Element Selection|"
                    r"[\d|\] ]{6,}|X?-?\d+\.\d+|Microsoft Teams|Recorded by|Organized by|Lyons, Addyson)$", re.I)


def digest(path: str) -> str:
    seen: set[str] = set()
    out = []
    lines = open(path, encoding="utf-8").read().splitlines()
    out.append(lines[0])
    for ln in lines[1:]:
        m = re.match(r"\[([\d:]+)\] (SAY|SCREEN): (.*)", ln)
        if not m:
            continue
        t, kind, txt = m.groups()
        if kind == "SAY":
            out.append(f"[{t}] SAY: {txt}")
            continue
        parts = [p.strip() for p in txt.split(" / ")]
        keep = [p for p in parts if p not in seen and len(p) > 3 and not BOILER.match(p)]
        seen.update(parts)
        if keep:
            out.append(f"[{t}] SCREEN: " + " / ".join(keep)[:700])
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    for tr in sorted(glob.glob(os.path.join(DERIVED, "*", "transcript.md"))):
        if a.only and a.only.lower() not in tr.lower():
            continue
        with open(os.path.join(os.path.dirname(tr), "digest.md"), "w", encoding="utf-8") as f:
            f.write(digest(tr))


if __name__ == "__main__":
    main()
