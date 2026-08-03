"""Dump plan callouts / dimension labels for Family 1 sheets (page 1)."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import group_rows, row_text, words_in_window

KEYS = (
    "TAPER", "BUFFER", "ROLL", "DOWNSTREAM", "SHOULDER", "MERGING", "SHIFTING",
    "SEE TABLE", "ARROW", "VEH", "W20", "W4-", "W9-", "G20", "NYW", "NYR", "R4-",
    "L/2", "L/3", "WORK AREA", "SPOTTER", "CONE", "CHANNEL"
)

for sheet in ["203", "202", "317", "312", "325", "414", "412", "423", "523"]:
    doc = fitz.open(ROOT / f"Bridge/captures/619-{sheet}.pdf")
    p = doc[0]
    # For 412 rotation=270, words are still in upright user space via get_text
    W = p.get_text("words")
    print(f"\n######## 619-{sheet} p1 rot={p.rotation} ########")
    # Prefer left/plan region; for 412 include more
    x1 = 900 if sheet != "412" else 800
    for r in group_rows(words_in_window(W, 0, 0, x1, 792), y_tol=5):
        t = row_text(r)
        up = t.upper()
        if any(k in up for k in KEYS):
            print(f"  y={r[0][1]:5.0f} x={r[0][0]:5.0f} {t[:130]}")
    # notes count
    notes_hits = [w for w in W if w[4] in ("NOTES:", "NOTES")]
    print(f"  NOTES markers: {[(round(w[0]), round(w[1])) for w in notes_hits]}")
