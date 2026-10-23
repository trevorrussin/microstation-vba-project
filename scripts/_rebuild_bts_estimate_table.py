"""BTS ESTIMATE OF QUANTITIES — TEXT ONLY.

Authority: `.cursor/skills/plan-markup-edits/SKILL.md` §1 + prior working
script pattern in this file's history / `scripts/_fix_bts_table_format.py`.

DO:
  - Edit item / unit / qty via TextNode SetTextLine (or DeleteAllTextLines
    + AddTextLine on short cells only).
  - Place / replace DESCRIPTION as CreateTextElement1 (never TextNode —
    TCSH_P_T TextNodes auto-wrap ~22 chars).
  - For new rows: Clone an existing full-width horizontal row band and
    Move it by (0, dy, 0). Same for short verticals if present.
  - Extend tall column fills only by replacing with a new axis-aligned
    CreateShapeElement1 rectangle (never ModifyVertex on live shapes).

DO NOT:
  - ModifyVertex on table border shapes (creates diagonal lines).
  - Freestyle rebuild of the whole border grid.
  - Invent description wraps — use SAMPLE LIST OF PAY ITEMS / markup.
"""
from __future__ import annotations

# Re-export the safe text-format path as the default entry.
from _fix_bts_table_format import ROWS, main  # noqa: F401

if __name__ == "__main__":
    raise SystemExit(
        "Do not run a freestyle border rebuild. "
        "Use scripts/_fix_bts_table_format.py for text, or Clone+Move "
        "full-width bands for new rows. See module docstring."
    )
