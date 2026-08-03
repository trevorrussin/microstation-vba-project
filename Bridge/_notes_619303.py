"""Notes extraction for 619-303 with tighter windows."""
import fitz
from collections import defaultdict

pg = fitz.open("Bridge/captures/619-303.pdf")[0]
W = pg.get_text("words")

# Find NOTES: header
for w in W:
    if w[4] == "NOTES:" or (w[4] == "NOTES" and w[0] > 450):
        print("NOTES header at", w[0], w[1])

# Try notes column - from recon NOTES was around x=505
# Also there's NOTES: at 251.6, 791.8 - that's in tables area
# Plan notes might be left of tables. Try x 400-780, y 250-620

def dump(x0, x1, y0, y1, title):
    print(f"\n=== {title} ({x0}-{x1}, {y0}-{y1}) ===")
    ww = [w for w in W if x0 <= w[0] <= x1 and y0 <= w[1] <= y1]
    rows = defaultdict(list)
    for w in ww:
        rows[round(w[1] / 2.5)].append(w)
    for k in sorted(rows):
        print(" ".join(x[4] for x in sorted(rows[k], key=lambda t: t[0])))

dump(400, 780, 240, 620, "mid notes candidate")
dump(780, 1220, 240, 420, "right notes near tables")

# Numbered note starts
print("\n=== numbered note anchors ===")
for w in sorted(W, key=lambda t: (t[1], t[0])):
    if w[4].rstrip(".") in [str(i) for i in range(1, 12)] and 400 < w[0] < 900 and 240 < w[1] < 700:
        # neighbors
        band = [x for x in W if abs(x[1]-w[1]) < 3 and w[0]-2 <= x[0] <= w[0]+200]
        print(f"  y={w[1]:6.1f} x={w[0]:6.1f} {' '.join(x[4] for x in sorted(band, key=lambda t:t[0]))[:100]}")
