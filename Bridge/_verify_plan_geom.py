"""Verify every `confidence: "drawing"` claim in Data/sheet-specs/619-311.json
against the PDF vector layer -- no pixel measurement, no visual inference.

Method: the plan's dimension lines are long vertical strokes in narrow x
bands. Their endpoints are the true segment boundaries. Each dimension's text
label sits at the midpoint of the segment it dimensions, so matching label
centres to segments labels every segment deterministically. Symbols (arrow
panel, vehicles, signs) are orange vector blobs whose y-extent can then be
tested against those boundaries.
"""
import collections

import fitz

ORANGE = (1.0, 0.5, 0.0)
pg = fitz.open("Bridge/captures/619-311.pdf")[0]
dr = pg.get_drawings()
words = pg.get_text("words")


def near(c, t, tol=0.06):
    return c is not None and all(abs(a - b) <= tol for a, b in zip(c, t))


# ------------------------------------------------ 1. all vertical dim columns
cols = collections.defaultdict(list)
for p in dr:
    for it in p["items"]:
        if it[0] != "l":
            continue
        (x0, y0), (x1, y1) = it[1], it[2]
        if abs(x0 - x1) < 0.4 and abs(y1 - y0) > 12 and 60 < x0 < 430:
            cols[round(x0, 1)].append((round(min(y0, y1), 1), round(max(y0, y1), 1)))

print("=== vertical dimension columns (>=2 segments) ===")
for x in sorted(cols):
    if len(cols[x]) >= 2:
        print(f"  x={x:6.1f}  {len(cols[x])} segs  "
              f"{sorted(cols[x])}")

# ------------------------------------------------ 2. labels on both sides
def label_blocks(x0, x1):
    rows = collections.defaultdict(list)
    for w in words:
        if x0 < w[0] < x1 and w[1] < 770:
            rows[round(w[1], 1)].append(w)
    blocks, cur = [], []
    for y in sorted(rows):
        if cur and y - cur[-1] > 9:
            blocks.append(cur)
            cur = []
        cur.append(y)
    if cur:
        blocks.append(cur)
    out = []
    for b in blocks:
        text = " ".join(" ".join(w[4] for w in sorted(rows[y], key=lambda w: w[0]))
                        for y in b)
        out.append(((b[0] + b[-1] + 6) / 2, text))
    return out


print("\n=== label -> segment match, per column ===")
matched = {}
for x in sorted(cols):
    if len(cols[x]) < 2:
        continue
    segs = sorted(set(cols[x]))
    side = (x - 60, x + 80) if x < 200 else (x - 8, x + 80)
    for centre, text in label_blocks(*side):
        hits = [(a, b) for a, b in segs if a - 3 <= centre <= b + 3]
        if not hits:
            continue
        a, b = min(hits, key=lambda s: s[1] - s[0])
        print(f"  x={x:6.1f}  y {a:7.1f}->{b:7.1f}  len={b - a:6.1f}   {text[:52]}")
        matched.setdefault(text, (a, b, x))

# ------------------------------------------------ 3. orange blobs (union-find)
rects = [+p["rect"] for p in dr
         if 60 <= p["rect"].x0 and p["rect"].x1 <= 430
         and (near(p.get("color"), ORANGE) or near(p.get("fill"), ORANGE))]
blobs = [+r for r in rects]
changed = True
while changed:
    changed = False
    out = []
    for r in blobs:
        for b in out:
            if r.intersects(b + (-3, -3, 3, 3)):
                b |= r
                changed = True
                break
        else:
            out.append(+r)
    blobs = out
big = sorted((b for b in blobs if b.width >= 12 and b.height >= 10), key=lambda r: r.y0)


def classify(b):
    ar = b.width / b.height
    if b.width > 45 and b.height > 45 and 0.85 < ar < 1.2:
        return "sign diamond"
    if 1.7 < ar < 2.6 and b.width > 35:
        return "panel/rect sign"
    if ar < 0.55 and b.height > 35:
        return "VEHICLE"
    return "?"


print("\n=== orange symbol blobs in main plan ===")
for b in big:
    print(f"  y {b.y0:7.1f}->{b.y1:7.1f}  x {b.x0:6.1f}-{b.x1:6.1f}  "
          f"{b.width:5.1f}x{b.height:5.1f}  {classify(b)}")


# ------------------------------------------------ 4. claim checks
def seg(kw):
    for t, v in matched.items():
        if kw in t:
            return v
    return None


def check(name, cond, detail):
    print(f"  [{'PASS' if cond else 'FAIL'}] {name}\n         {detail}")


print("\n=== SPEC CLAIM CHECKS ===")
lane, shld = seg("LANE TAPER"), seg("SHOULDER TAPER")
buf, roll = seg("BUFFER SPACE"), seg("ROLL AHEAD")
gA, gB, gC = seg("A (SEE"), seg("B (SEE"), seg("C (SEE")
down = seg("DOWNSTREAM TAPER")

if lane and shld:
    check("taper continuity: lane taper's upstream end == shoulder taper's downstream end",
          abs(lane[1] - shld[0]) < 3,
          f"lane {lane[0]}->{lane[1]}, shoulder {shld[0]}->{shld[1]}")
if gA and shld and lane:
    check("gap A is datumed at the SHOULDER taper upstream end (spec claim)",
          abs(gA[0] - shld[1]) < 3,
          f"A starts y={gA[0]}; shoulder taper upstream end y={shld[1]}; "
          f"lane taper upstream end y={lane[1]}")
if gA and gB and gC:
    check("A/B/C are contiguous walking upstream",
          abs(gA[1] - gB[0]) < 3 and abs(gB[1] - gC[0]) < 3,
          f"A {gA[0]}->{gA[1]}, B {gB[0]}->{gB[1]}, C {gC[0]}->{gC[1]}")

panels = [b for b in big if classify(b) == "panel/rect sign" and b.y0 > 200]
if panels and lane and shld:
    ap = min(panels, key=lambda b: b.y0)
    check("arrow panel sits at the lane-taper / shoulder-taper junction",
          min(abs(ap.y0 - lane[1]), abs(ap.y1 - lane[1]),
              abs((ap.y0 + ap.y1) / 2 - lane[1])) < 15,
          f"AP y {ap.y0:.1f}-{ap.y1:.1f}; junction y={lane[1]}; "
          f"shoulder taper upstream end y={shld[1]}")

vehs = [b for b in big if classify(b) == "VEHICLE"]
print(f"\n  vehicles: {[f'{v.y0:.0f}-{v.y1:.0f}' for v in vehs]}")
if vehs and roll:
    v2 = min(vehs, key=lambda b: b.y0)
    check("VEH #2 at the upstream end of the roll ahead distance",
          abs(v2.y0 - roll[1]) < 6,
          f"vehicle y {v2.y0:.1f}-{v2.y1:.1f}; roll ahead {roll[0]}->{roll[1]}")
if len(vehs) > 1 and panels:
    v1 = max(vehs, key=lambda b: b.y0)
    ap = min(panels, key=lambda b: b.y0)
    check("VEH #1 shares the arrow panel station (the sheet's 'OR')",
          abs((v1.y0 + v1.y1) / 2 - (ap.y0 + ap.y1) / 2) < 20,
          f"VEH#1 y {v1.y0:.1f}-{v1.y1:.1f}; AP y {ap.y0:.1f}-{ap.y1:.1f}")

diam = sorted((b for b in big if classify(b) == "sign diamond"), key=lambda b: b.y0)
print(f"\n  sign diamonds at y: {[f'{(d.y0 + d.y1) / 2:.0f}' for d in diam]}")
if len(diam) >= 2 and gA and gB:
    check("first advance sign (W4-2R) sits at the upstream end of gap A",
          abs((diam[0].y0 + diam[0].y1) / 2 - gA[1]) < 15,
          f"diamond centre y={(diam[0].y0 + diam[0].y1) / 2:.1f}; A ends y={gA[1]}")
    check("second advance sign (W20-5R) sits at the upstream end of gap B",
          abs((diam[1].y0 + diam[1].y1) / 2 - gB[1]) < 15,
          f"diamond centre y={(diam[1].y0 + diam[1].y1) / 2:.1f}; B ends y={gB[1]}")
if diam and shld:
    check("all advance signs upstream of the shoulder taper",
          all((d.y0 + d.y1) / 2 > shld[1] for d in diam),
          f"shoulder taper upstream end y={shld[1]}")
