"""Clean corridor evidence for 619-303 from dimension columns + labels."""
from __future__ import annotations

import collections

import fitz

ORANGE = (1.0, 0.5, 0.0)
pg = fitz.open("Bridge/captures/619-303.pdf")[0]
dr = pg.get_drawings()
words = pg.get_text("words")


def near(c, t, tol=0.06):
    return c is not None and len(c) == 3 and all(abs(a - b) <= tol for a, b in zip(c, t))


# Main dimension column candidates with clean labels
cols = collections.defaultdict(set)
for p in dr:
    for it in p["items"]:
        if it[0] != "l":
            continue
        (ax, ay), (bx, by) = it[1], it[2]
        if abs(ax - bx) < 0.4 and abs(by - ay) > 12 and 140 < ax < 300:
            cols[round(ax, 1)].add((round(min(ay, by), 1), round(max(ay, by), 1)))

print("=== dimension columns worth reading ===")
for x in sorted(cols):
    segs = sorted(cols[x])
    if len(segs) < 2:
        continue
    span = segs[-1][1] - segs[0][0]
    print(f"x={x:6.1f} n={len(segs):2d} span={span:6.1f}")

# Labels in plan with multi-line join, restricted
print("\n=== plan label blocks (x 100-320) ===")
rows = collections.defaultdict(list)
for w in words:
    if 100 < w[0] < 320 and 50 < w[1] < 760:
        rows[round(w[1], 1)].append(w)
blocks, cur = [], []
for y in sorted(rows):
    if cur and y - cur[-1] > 8:
        blocks.append(cur)
        cur = []
    cur.append(y)
if cur:
    blocks.append(cur)
for b in blocks:
    text = " ".join(" ".join(w[4] for w in sorted(rows[y], key=lambda w: w[0])) for y in b)
    cy = (b[0] + b[-1]) / 2
    # skip pure notes soup
    if any(k in text.upper() for k in (
        "TAPER", "BUFFER", "ROLL", "SHOULDER", "DOWNSTREAM", "VEH", "A (", "B (", "C (",
        "ROAD WORK", "LANE CLOSED", "END ROAD", "2L", " MERGING")):
        print(f"  y~{cy:6.1f}  {text[:90]}")

# Match main column x=281.8 segments to labels
print("\n=== x=281.8 segments ===")
for a, b in sorted(cols.get(281.8, [])):
    mid = (a + b) / 2
    cands = []
    for blk in blocks:
        cy = (blk[0] + blk[-1]) / 2
        if a - 5 <= cy <= b + 5:
            text = " ".join(" ".join(w[4] for w in sorted(rows[y], key=lambda w: w[0])) for y in blk)
            cands.append((abs(cy - mid), text[:70]))
    lab = min(cands)[1] if cands else "?"
    print(f"  {a:7.1f}->{b:7.1f} len={b-a:6.1f}  {lab}")

print("\n=== x=154.5 segments (left, tapers) ===")
for a, b in sorted(cols.get(154.5, [])):
    mid = (a + b) / 2
    cands = []
    for blk in blocks:
        cy = (blk[0] + blk[-1]) / 2
        if a - 5 <= cy <= b + 5:
            text = " ".join(" ".join(w[4] for w in sorted(rows[y], key=lambda w: w[0])) for y in blk)
            cands.append((abs(cy - mid), text[:70]))
    lab = min(cands)[1] if cands else "?"
    print(f"  {a:7.1f}->{b:7.1f} len={b-a:6.1f}  {lab}")

# Notes column
print("\n=== NOTES (x 480-720, y 250-750) ===")
nw = [w for w in words if 480 <= w[0] <= 720 and 250 <= w[1] <= 750]
nr = collections.defaultdict(list)
for w in nw:
    nr[round(w[1] / 3)].append(w)
for k in sorted(nr):
    print(" ", " ".join(w[4] for w in sorted(nr[k], key=lambda t: t[0])))

# Render plan crop for visual anchor check
pix = pg.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), clip=fitz.Rect(40, 30, 450, 760))
pix.save("Bridge/captures/sheet_619303_plan.png")
print("\nwrote Bridge/captures/sheet_619303_plan.png")
