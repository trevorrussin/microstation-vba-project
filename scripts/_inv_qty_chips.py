"""Inventory cyan Square qty chips + green pay-item FreeTexts on BTS page 3."""
from __future__ import annotations

import fitz
from pathlib import Path

pdf = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)
page = fitz.open(pdf)[2]

print("=== ALL Square annots ===")
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "Square":
        continue
    r = a.rect
    content = ((a.info or {}).get("content") or "").strip()
    colors = a.colors or {}
    print(
        f"[{i}] rect=({r.x0:.1f},{r.y0:.1f})-({r.x1:.1f},{r.y1:.1f}) "
        f"stroke={colors.get('stroke')} fill={colors.get('fill')} content={content!r}"
    )

print("\n=== ALL annots with short numeric content ===")
for i, a in enumerate(page.annots() or []):
    content = ((a.info or {}).get("content") or "").strip()
    if not content:
        continue
    if len(content) > 20:
        continue
    r = a.rect
    colors = a.colors or {}
    print(
        f"[{i}] {a.type[1]} ({r.x0:.0f},{r.y0:.0f})-({r.x1:.0f},{r.y1:.0f}) "
        f"stroke={colors.get('stroke')} {content!r}"
    )

# Also dump FreeText that look like P-codes
print("\n=== P-code / PB FreeTexts ===")
for i, a in enumerate(page.annots() or []):
    content = ((a.info or {}).get("content") or "").strip()
    if content.upper().startswith("P") and len(content) <= 6:
        r = a.rect
        print(f"[{i}] {a.type[1]} ({r.x0:.0f},{r.y0:.0f}) {content!r}")
