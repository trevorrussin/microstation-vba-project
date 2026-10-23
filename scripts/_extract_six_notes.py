"""Dump full NOTES area from r1 page 3 (words + FreeTexts near notes)."""
from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
page = fitz.open(PDF)[2]
print("page size", page.rect)

# All FreeTexts that look like notes 1-6
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "FreeText":
        continue
    c = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    r = a.rect
    cu = c.upper()
    if any(
        k in cu
        for k in (
            "PULL IN",
            "ENSURE",
            "AND DETAILS",
            "FOOTING",
            "SIDEWALK BOX",
            "REFER TO DWG",
            "REFER TO TSP",
            "REFER TO NYCDOT",
            "LEFT JUSTIFY",
            "ADD PERIOD",
        )
    ) or (r.x0 < 250 and len(c) > 15):
        print(f"[{i}] @({r.x0:.0f},{r.y0:.0f}-{r.y1:.0f}) {c!r}")

# Reconstruct left-column note text from words sorted by y
words = page.get_text("words")
left = [w for w in words if w[0] < 280]
left.sort(key=lambda w: (round(w[1], 0), w[0]))
# group into lines
lines = []
cur_y = None
cur = []
for w in left:
    y = round(w[1], 0)
    if cur_y is None or abs(y - cur_y) <= 3:
        cur.append(w[4])
        cur_y = y if cur_y is None else cur_y
    else:
        lines.append((cur_y, " ".join(cur)))
        cur = [w[4]]
        cur_y = y
if cur:
    lines.append((cur_y, " ".join(cur)))

print("\n=== LEFT COLUMN LINES (y sorted) ===")
for y, t in lines:
    tu = t.upper()
    if any(
        k in tu
        for k in (
            "NOTE",
            "REFER",
            "ENSURE",
            "PULL",
            "FOOTING",
            "SIDEWALK",
            "FURNISH",
            "CONDUIT",
            "TSN",
            "TSP",
            "NYCDOT",
            "DETAILS",
            "1.",
            "2.",
            "3.",
            "4.",
            "5.",
            "6.",
        )
    ) or (len(t) > 40 and t[0].isdigit()):
        print(f"{y:6.0f} {t}")
