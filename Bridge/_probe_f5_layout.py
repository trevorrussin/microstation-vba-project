"""Probe Family 5 PDF layouts — table titles, plan tokens, notes."""
from __future__ import annotations

import re
import pathlib
import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]
SHEETS = [318, 316, 319, 113, 211, 416, 417, 418, 517, 518]


def analyze(sheet: int) -> None:
    pdf = ROOT / f"Bridge/captures/619-{sheet}.pdf"
    doc = fitz.open(str(pdf))
    print(f"\n######## {sheet} pages={doc.page_count} ########")
    for pi, p in enumerate(doc):
        words = list(p.get_text("words"))
        print(f"--- page {pi + 1} rot={p.rotation} words={len(words)} ---")
        for i, w in enumerate(words):
            if re.fullmatch(r"\d{3}-\d{2}", w[4]):
                title = [
                    x[4]
                    for x in words
                    if abs(x[1] - w[1]) < 6 and x[0] >= w[0] - 40 and x[0] < w[0] + 450
                ]
                joined = " ".join(title)[:120]
                print(f"  {w[4]} @x={w[0]:.0f} y={w[1]:.0f} :: {joined}")

        for tok in [
            "MERGING",
            "DOWNSTREAM",
            "SHOULDER",
            "BUFFER",
            "ROLL",
            "ARROW",
            "RAMP",
            "1000",
            "1500",
            "1320",
            "2640",
            "500",
            "W20-1",
            "W20-5",
            "W4-",
            "W21-",
            "NYW8",
            "G20-2",
            "W4-1",
            "W4-2",
        ]:
            hits = [
                (round(w[0]), round(w[1]), w[4])
                for w in words
                if tok.upper() in w[4].upper()
            ]
            if not hits:
                continue
            if len(hits) <= 6:
                print(f"  tok {tok}: {hits}")
            else:
                print(f"  tok {tok}: n={len(hits)} first={hits[0]} last={hits[-1]}")

        # NOTES region: look for numbered notes
        note_starts = [
            (round(w[0]), round(w[1]), w[4])
            for w in words
            if w[4] in ("NOTES:", "NOTES") or re.fullmatch(r"\d\.", w[4])
        ]
        print(f"  note-ish: {note_starts[:20]}")


if __name__ == "__main__":
    for s in SHEETS:
        analyze(s)
