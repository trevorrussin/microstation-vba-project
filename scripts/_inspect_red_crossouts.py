"""Inspect red ink/crossouts on BTS page 3 — what text do they cover?"""
from __future__ import annotations

from pathlib import Path
import fitz

PDF = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)

page = fitz.open(PDF)[2]
words = page.get_text("words")

print("=== RED annotations (stroke ~ red) ===")
for i, a in enumerate(page.annots() or []):
    stroke = (a.colors or {}).get("stroke") or []
    fill = (a.colors or {}).get("fill") or []
    # red-ish
    is_red = False
    for c in (stroke, fill):
        if c and len(c) >= 3 and c[0] > 0.7 and c[1] < 0.35 and c[2] < 0.35:
            is_red = True
    if not is_red and a.type[1] not in ("StrikeOut", "Squiggly"):
        continue
    at = a.type[1]
    content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    r = a.rect
    # expand a bit for ink that sits on glyphs
    pad = 3
    covered = [
        w
        for w in words
        if w[0] < r.x1 + pad
        and w[2] > r.x0 - pad
        and w[1] < r.y1 + pad
        and w[3] > r.y0 - pad
    ]
    # sort reading order
    covered.sort(key=lambda w: (round(w[1], 0), w[0]))
    cov_txt = " ".join(w[4] for w in covered)
    print(f"\n[{i}] {at} stroke={stroke} fill={fill}")
    print(f"  rect=({r.x0:.1f},{r.y0:.1f})-({r.x1:.1f},{r.y1:.1f}) "
          f"size={r.width:.1f}x{r.height:.1f}")
    print(f"  annot content: {content!r}")
    print(f"  covered words ({len(covered)}): {cov_txt[:400]!r}")

print("\n=== StrikeOut type count ===")
n = sum(1 for a in (page.annots() or []) if a.type[1] == "StrikeOut")
print("StrikeOut annots:", n)

# Also list green FreeTexts near red inks for pairing
print("\n=== Red ink near green conduit FreeTexts? ===")
greens = []
reds = []
for i, a in enumerate(page.annots() or []):
    stroke = (a.colors or {}).get("stroke") or []
    r = a.rect
    content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    if a.type[1] == "FreeText" and stroke == [0.0, 1.0, 0.0] and "TRAFFIC" in content.upper():
        greens.append((i, r, content.splitlines()[0] if content else ""))
    if a.type[1] == "Ink" and stroke and stroke[0] > 0.7 and stroke[1] < 0.35:
        reds.append((i, r, content))

for ri, rr, rc in reds:
    near = []
    for gi, gr, gf in greens:
        # distance between rect centers
        dx = ((rr.x0 + rr.x1) / 2) - ((gr.x0 + gr.x1) / 2)
        dy = ((rr.y0 + rr.y1) / 2) - ((gr.y0 + gr.y1) / 2)
        dist = (dx * dx + dy * dy) ** 0.5
        if dist < 120:
            near.append((dist, gi, gf))
    near.sort()
    print(f"red[{ri}] '{rc}' @({rr.x0:.0f},{rr.y0:.0f}) near_greens={near[:3]}")
