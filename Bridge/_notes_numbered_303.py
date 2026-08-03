"""Numbered plan notes for 619-303 (left NOTES column near plan)."""
from collections import defaultdict
import fitz
import re

pg = fitz.open("Bridge/captures/619-303.pdf")[0]
W = pg.get_text("words")

# Notes body: between plan and tables
ww = [w for w in W if 380 <= w[0] <= 780 and 280 <= w[1] <= 720]
rows = defaultdict(list)
for w in ww:
    rows[round(w[1], 1)].append(w)

lines = []
for y in sorted(rows):
    text = " ".join(w[4] for w in sorted(rows[y], key=lambda t: t[0]))
    lines.append((y, text))

# Print with y for splitting
for y, t in lines:
    print(f"{y:7.1f} {t}")

# Heuristic: a note starts when a line begins with N. where N is 1-9
print("\n=== assembled ===")
notes = {}
cur = None
buf = []
for y, t in lines:
    m = re.match(r"^([1-9])\.\s+(.*)$", t)
    if m:
        if cur is not None:
            notes[cur] = " ".join(buf)
        cur = int(m.group(1))
        buf = [m.group(2)]
    elif cur is not None:
        # stop if we hit legend/sign soup
        if t.startswith("W20") or t.startswith("G20") or t.startswith("TABLE") or t.startswith("END "):
            notes[cur] = " ".join(buf)
            cur = None
            buf = []
        else:
            buf.append(t)
if cur is not None:
    notes[cur] = " ".join(buf)
for n in sorted(notes):
    print(f"\n{n}. {notes[n]}")
