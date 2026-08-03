"""Dump the plan geometry of a 619 standard sheet from its PDF vector layer.

Run this BEFORE writing a sheet spec's `corridor` section. It replaces
"look at the drawing and describe the layout" -- which is where sheet specs
get silently wrong -- with a deterministic evidence dump.

Why it works: these sheets are vector PDFs, not scans. Dimension lines are
long vertical strokes in narrow x bands; their endpoints ARE the segment
boundaries. Each dimension's text label sits at the midpoint of the segment
it dimensions, so matching label centres to segments labels every segment.
Symbols are orange paths whose bounding boxes give exact stations.

Read the "DATUM SHARING" section carefully. Two dimensions that start at the
same y are measured from the same point, which means one may lie INSIDE the
other rather than following it. That distinction is invisible when eyeballing
a not-to-scale drawing and it changes the station math for everything
upstream (it was wrong in the first draft of 619-311).

Usage:
    python scripts/extract_plan_geometry.py Bridge/captures/619-311.pdf
    python scripts/extract_plan_geometry.py <pdf> --page 0 --plan-x 60 430
"""
from __future__ import annotations

import argparse
import collections
import pathlib

import fitz

ORANGE = (1.0, 0.5, 0.0)
YELLOW = (0.94, 0.94, 0.0)
GREY = (0.94, 0.94, 0.94)


def near(c, t, tol=0.06):
    return c is not None and len(c) == 3 and all(abs(a - b) <= tol for a, b in zip(c, t))


def is_colour(p, t):
    return near(p.get("color"), t) or near(p.get("fill"), t)


def dim_columns(dr, x0, x1, min_len=12.0):
    """Vertical strokes grouped by x -- candidate dimension/witness lines."""
    cols = collections.defaultdict(set)
    for p in dr:
        for it in p["items"]:
            if it[0] != "l":
                continue
            (ax, ay), (bx, by) = it[1], it[2]
            if abs(ax - bx) < 0.4 and abs(by - ay) > min_len and x0 < ax < x1:
                cols[round(ax, 1)].add((round(min(ay, by), 1), round(max(ay, by), 1)))
    return {x: sorted(s) for x, s in cols.items() if len(s) >= 2}


def label_blocks(words, x0, x1, y_max=780.0, gap=9.0):
    """Text grouped into multi-line label blocks, with each block's centre y."""
    rows = collections.defaultdict(list)
    for w in words:
        if x0 < w[0] < x1 and w[1] < y_max:
            rows[round(w[1], 1)].append(w)
    blocks, cur = [], []
    for y in sorted(rows):
        if cur and y - cur[-1] > gap:
            blocks.append(cur)
            cur = []
        cur.append(y)
    if cur:
        blocks.append(cur)
    out = []
    for b in blocks:
        text = " ".join(" ".join(w[4] for w in sorted(rows[y], key=lambda w: w[0]))
                        for y in b)
        out.append(((b[0] + b[-1] + 6) / 2, text.strip()))
    return out


def merge_blobs(rects, pad=3.0):
    """Union overlapping rects until stable (single pass leaves fragments)."""
    blobs = [+r for r in rects]
    changed = True
    while changed:
        changed = False
        out = []
        for r in blobs:
            for b in out:
                if r.intersects(b + (-pad, -pad, pad, pad)):
                    b |= r
                    changed = True
                    break
            else:
                out.append(+r)
        blobs = out
    return blobs


def classify(b):
    ar = b.width / b.height if b.height else 99
    if b.width > 45 and b.height > 45 and 0.85 < ar < 1.2:
        return "sign diamond"
    if 1.7 < ar < 2.8 and b.width > 30:
        return "panel / rect sign"
    if ar < 0.55 and b.height > 35:
        return "vehicle"
    if 0.8 < ar < 1.25 and 12 < b.width < 30:
        return "small square symbol"
    return "?"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", type=pathlib.Path)
    ap.add_argument("--page", type=int, default=0)
    ap.add_argument("--plan-x", type=float, nargs=2, default=(60, 430),
                    help="x window of the main plan column (exclude legend/tables)")
    args = ap.parse_args()

    pg = fitz.open(args.pdf)[args.page]
    dr = pg.get_drawings()
    words = pg.get_text("words")
    px0, px1 = args.plan_x

    print(f"{args.pdf.name} page {args.page}  "
          f"({len(dr)} vector paths, {len(words)} words)")
    print(f"plan x window: {px0}-{px1}\n")

    cols = dim_columns(dr, px0, px1)
    print("=== DIMENSION COLUMNS ===")
    for x in sorted(cols):
        segs = cols[x]
        span = segs[-1][1] - segs[0][0]
        # A real dimension column has several segments covering a long run.
        tag = "  <-- likely main dimension column" if len(segs) >= 4 and span > 200 else ""
        print(f"  x={x:7.1f}  {len(segs):2d} segments{tag}")

    print("\n=== SEGMENT -> LABEL ===")
    # One label per segment: the block whose centre is closest to the segment
    # midpoint, searched only near that column. A wide search window drags in
    # the notes column and produces word salad.
    best = {}
    for x in sorted(cols):
        near_x = label_blocks(words, max(px0, x - 60), min(px1, x + 70))
        for a, b in cols[x]:
            if b - a > 400:            # page-height border line, not a dimension
                continue
            mid = (a + b) / 2
            cands = [(abs(c - mid), t) for c, t in near_x if a - 3 <= c <= b + 3]
            if not cands:
                continue
            dist, text = min(cands)
            key = (a, b, x)
            if key not in best or dist < best[key][0]:
                best[key] = (dist, text)
    labelled = [(a, b, x, t) for (a, b, x), (_, t) in best.items()]
    for a, b, x, text in sorted(labelled):
        print(f"  y {a:7.1f} -> {b:7.1f}  len={b - a:6.1f}  x={x:6.1f}   {text[:60]}")

    print("\n=== DATUM SHARING (segments starting at the same y) ===")
    starts = collections.defaultdict(list)
    for a, b, x, text in labelled:
        starts[a].append((b, x, text[:40]))
    found = False
    for a in sorted(starts):
        if len({t for _, _, t in starts[a]}) > 1:
            found = True
            print(f"  datum y={a}:")
            for b, x, t in sorted(starts[a]):
                print(f"      -> y {b:7.1f} (len {b - a:6.1f}, x={x:6.1f})  {t}")
            print("      NOTE: shorter one lies INSIDE the longer one; it is an "
                  "overlay, not a sequential station.")
    if not found:
        print("  (none -- every dimension follows the previous one)")

    print("\n=== ORANGE SYMBOL BLOBS (station order, low y = downstream) ===")
    rects = [+p["rect"] for p in dr
             if px0 <= p["rect"].x0 and p["rect"].x1 <= px1 and is_colour(p, ORANGE)]
    for b in sorted((b for b in merge_blobs(rects)
                     if b.width >= 12 and b.height >= 10), key=lambda r: r.y0):
        print(f"  y {b.y0:7.1f} -> {b.y1:7.1f}  x {b.x0:6.1f}-{b.x1:6.1f}  "
              f"{b.width:5.1f} x {b.height:5.1f}   {classify(b)}")

    print("\n=== LATERAL REFERENCE (long vertical runs by colour) ===")
    for name, col in (("yellow/centreline", YELLOW), ("grey/pavement", GREY)):
        xs = sorted({round((p["rect"].x0 + p["rect"].x1) / 2, 1) for p in dr
                     if is_colour(p, col) and p["rect"].height > 60
                     and p["rect"].width < 4 and px0 <= p["rect"].x0 <= px1})
        print(f"  {name:18s} x = {xs[:12]}")

    print("\nNext: write the corridor from the SEGMENT -> LABEL list in y order, "
          "then reconcile DATUM SHARING before assigning order-table rows.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
