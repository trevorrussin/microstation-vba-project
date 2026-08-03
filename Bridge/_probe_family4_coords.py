"""Probe Family 4 PDF table coordinates for extraction windows."""
from __future__ import annotations

from collections import defaultdict
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]


def dump_sheet(sheet: str) -> None:
    pdf = ROOT / f"Bridge/captures/{sheet}.pdf"
    doc = fitz.open(pdf)
    pg = doc[0]
    words = pg.get_text("words")
    print("=" * 70)
    print(sheet, "rot", pg.rotation, "rect", tuple(pg.rect), "nwords", len(words))

    for w in words:
        if w[4].upper() == "TABLE":
            nearby = sorted(
                [x for x in words if abs(x[1] - w[1]) < 5 and x[0] >= w[0] - 5],
                key=lambda x: x[0],
            )
            line = " ".join(x[4] for x in nearby[:14])
            print(f"  TABLE @ ({w[0]:.0f},{w[1]:.0f}): {line}")

    interesting = (
        "MERGING", "SHOULDER", "DOWNSTREAM", "ROLL", "BUFFER",
        "1000", "1500", "2640", "500", "1320",
        "W20-1", "W20-5R", "W4-2R", "G20-2", "NYW8-33", "W8-23",
        "PARKWAY", "MOWING", "WARNING", "45", "50", "55", "65", "40",
        "P,", "PVH", "TMIA", "120/", "160/", "200/", "240/", "280/",
        "360/", "480/", "560/", "640/", "80/", "L/3", "L (",
    )
    print("  -- notable tokens --")
    for w in sorted(words, key=lambda x: (x[1], x[0])):
        t = w[4]
        if any(k in t for k in interesting) or (
            t.count("/") >= 1 and t[0].isdigit() and len(t) <= 12
        ):
            print(f"    ({w[0]:6.1f},{w[1]:6.1f}) {t}")

    # notes numbers
    print("  -- notes-ish lines (left half) --")
    left = [w for w in words if w[0] < 420]
    rows: dict[int, list] = defaultdict(list)
    for w in left:
        rows[round(w[1] / 4)].append(w)
    for k in sorted(rows)[:40]:
        cells = sorted(rows[k], key=lambda x: x[0])
        txt = " ".join(c[4] for c in cells)
        if any(ch.isdigit() for ch in txt[:3]) or "NOTE" in txt.upper():
            print(f"    y~{cells[0][1]:.0f}: {txt[:120]}")


if __name__ == "__main__":
    for s in ("619-306", "619-212", "619-114", "619-041"):
        dump_sheet(s)
