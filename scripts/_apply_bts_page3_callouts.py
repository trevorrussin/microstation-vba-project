"""Apply BTS r1 page-3 plan callouts to fea_tsp (NEW conduit/pull + EDIT items)."""
from __future__ import annotations

import math
import sys
import time

import numpy as np

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

import fitz  # noqa: E402

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
adf = app.ActiveDesignFile
q = app.CadInputQueue
as_ = app.ActiveSettings

assert "fea_tsp" in adf.FullName.lower(), adf.FullName
print("file", adf.FullName, "model", mr.Name)

# Digit-anchor affine (BTS_2 / same sheet geometry)
ANCHORS = [
    (662.6, 453.2, 1029505.95, 239730.47),
    (649.1, 468.6, 1029501.96, 239722.75),
    (633.7, 468.6, 1029494.05, 239721.88),
    (609.8, 363.9, 1029480.91, 239776.98),
    (624.4, 427.7, 1029488.63, 239743.33),
    (625.7, 443.1, 1029490.00, 239734.42),
]
_A = np.array([[1, px, py] for px, py, _, _ in ANCHORS], float)
_AX = np.linalg.lstsq(_A, np.array([dx for _, _, dx, _ in ANCHORS]), rcond=None)[0]
_AY = np.linalg.lstsq(_A, np.array([dy for _, _, _, dy in ANCHORS]), rcond=None)[0]


def pdf_to_dgn(px: float, py: float) -> tuple[float, float]:
    return float(_AX[0] + _AX[1] * px + _AX[2] * py), float(
        _AY[0] + _AY[1] * px + _AY[2] * py
    )


def pt3(x, y, z=0.0):
    return app.Point3dFromXYZ(float(x), float(y), float(z))


def next_gg() -> int:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    mx = 7000
    while ee.MoveNext():
        try:
            mx = max(mx, int(ee.Current.GraphicGroup))
        except Exception:
            pass
    return mx + 1


def set_tn_lines(tn_el, lines: list[str]) -> None:
    tn = tn_el.AsTextNodeElement
    # Set existing lines; AddTextLine for extras. Never DeleteAllTextLines.
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    # blank leftover lines if any
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    tn_el.Rewrite()


def nearest_tci_point(xy: tuple[float, float]) -> tuple[float, float]:
    """Snap tip to nearest TCI / TCI_P element range center."""
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    best = None
    best_d = 1e18
    while ee.MoveNext():
        e = ee.Current
        try:
            lv = e.Level.Name if e.Level else ""
            if lv not in ("TCI", "TCI_P"):
                continue
            r = e.Range
            cx = 0.5 * (float(r.Low.X) + float(r.High.X))
            cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
            d = (cx - xy[0]) ** 2 + (cy - xy[1]) ** 2
            if d < best_d:
                best_d = d
                best = (cx, cy)
        except Exception:
            continue
    return best if best is not None else xy


# ---------------------------------------------------------------------------
# EDIT existing callouts
# ---------------------------------------------------------------------------
EDITS = [
    # tn_id, new lines
    (
        6132,
        [
            "VEHICULAR SIGNAL HEADS AND APS PUSH BUTTON",
            "(ITEM NOS. 680.818122011 (1), 680.81822011 (1), 680.8223 (1), 680.8151 (1))",
        ],
    ),
    (
        6120,
        [
            "STREETLIGHT W/PEDESTRIAN TRAFFIC SIGNALS AND APS PUSH BUTTON",
            "(ITEM NOS. 680.8141 (2), 680.815002 (2), 680.8151 (1), 680.83080011 (1))",
        ],
    ),
    (
        6129,
        [
            "APS W/POST",
            "(ITEM NOS. 680.5001, 680.8152 (1))",
        ],
    ),
    (
        6126,
        [
            "APS W/POST",
            "(ITEM NOS. 680.5001, 680.8152 (1))",
        ],
    ),
    (
        6280,
        [
            "APS W/POST",
            "(ITEM NOS. 680.5001, 680.8152 (1))",
        ],
    ),
]

print("=== EDITs ===")
for tn_id, lines in EDITS:
    el = mr.GetElementByID64(tn_id)
    before = " | ".join(
        el.AsTextNodeElement.TextLine(i)
        for i in range(1, int(el.AsTextNodeElement.TextLinesCount) + 1)
    )
    set_tn_lines(el, lines)
    after = " | ".join(
        mr.GetElementByID64(tn_id).AsTextNodeElement.TextLine(i)
        for i in range(1, int(mr.GetElementByID64(tn_id).AsTextNodeElement.TextLinesCount) + 1)
    )
    print(f"  {tn_id}: {before!r}")
    print(f"       -> {after!r}")

# ---------------------------------------------------------------------------
# NEW callouts — clone stock GG 6119 (TN 6132 / stub 6143 / dim 6150)
# ---------------------------------------------------------------------------
TMPL_TN = 6132
TMPL_STUB = 6143
TMPL_DIM = 6150

tmpl_tn = mr.GetElementByID64(TMPL_TN)
tmpl_origin = tmpl_tn.AsTextNodeElement.Origin
tmpl_ox, tmpl_oy = float(tmpl_origin.X), float(tmpl_origin.Y)

pdf = fitz.open(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)
page = pdf[2]

# Use PDF text-block anchors for conduit blocks (FreeText rects are oversized).
# Collect first-line positions for matching.
text_lines = []
for b in page.get_text("dict")["blocks"]:
    if b.get("type") != 0:
        continue
    for line in b.get("lines", []):
        spans = line.get("spans", [])
        if not spans:
            continue
        t = "".join(s["text"] for s in spans).strip()
        if not t:
            continue
        bb = line["bbox"]
        text_lines.append((t, (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2, bb))


def find_block_origin(first_line: str, freerect) -> tuple[float, float, float, float]:
    """Return (pdf_x_right_or_left, pdf_y, dgn_ox, dgn_oy) for RJ callout origin."""
    candidates = []
    for t, cx, cy, bb in text_lines:
        if t.strip() == first_line.strip() or (
            first_line.startswith("TRAFFIC SIGNAL") and t.startswith(first_line[:28])
        ):
            # prefer inside/near FreeText rect
            if freerect.x0 - 40 <= cx <= freerect.x1 + 40 and freerect.y0 - 40 <= cy <= freerect.y1 + 40:
                candidates.append((t, cx, cy, bb))
    if not candidates:
        # fallback: FreeText edge
        cx = (freerect.x0 + freerect.x1) / 2
        cy = (freerect.y0 + freerect.y1) / 2
        if cx >= 550:
            px = freerect.x0
        else:
            px = freerect.x1
        dx, dy = pdf_to_dgn(px, cy)
        return px, cy, dx, dy
    # pick candidate whose y is closest to FreeText center
    fcy = (freerect.y0 + freerect.y1) / 2
    t, cx, cy, bb = min(candidates, key=lambda c: abs(c[2] - fcy))
    # right-justified: origin at right of text if text is left of feature (cx < 600),
    # else origin at left of text (toward plan) for right-side stacks
    if cx >= 550:
        px = bb[0]  # left edge — RJ text grows left from a west-side origin? 
        # Stock RJ origin is at RIGHT of text. For right-side callouts pointing left,
        # origin should be at LEFT of text (near leader stub toward plan).
        # We'll place origin at left edge and keep just=12 (RJ) so text extends further left —
        # that pushes text away from plan. Better: origin at right edge of text block
        # (east), stub west toward tip. That mirrors stock flipped.
        # Skill stock: text LEFT of feature, origin RIGHT of text, stub RIGHT to tip EAST.
        # Right-side markup: text RIGHT of feature, tip WEST. So origin = LEFT of text,
        # stub LEFT toward tip. Justification should be LEFT (0) for that — but keep stock
        # RJ and put origin at left edge ≈ stub side: RJ will extend text further left
        # overlapping tip. So use origin at RIGHT of text (bb[2]) for RJ stock clone —
        # text sits east of origin; we put tip west of origin. Good.
        px = bb[2]
    else:
        px = bb[2]  # right edge toward plan (east of left-side text)
    dx, dy = pdf_to_dgn(px, cy)
    return px, cy, dx, dy


NEWS = []
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "FreeText":
        continue
    if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
        continue
    content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    u = content.upper()
    is_new = (
        "TRAFFIC SIGNAL CONDUIT" in u
        or "TRAFFIC SIGNAL CABLE (" in u
        or "PULL NEW CABLES" in u
    )
    if not is_new:
        continue
    lines = [ln.strip() for ln in content.splitlines() if ln.strip()]
    first = lines[0]
    _, _, ox, oy = find_block_origin(first, a.rect)
    # tip: from text origin toward plan conduit, snap to TCI
    tip_guess = nearest_tci_point((ox - 25.0 if ox > 1029480 else ox + 25.0, oy))
    NEWS.append({"idx": i, "lines": lines, "origin": (ox, oy), "tip": tip_guess})

print(f"=== NEW count {len(NEWS)} ===")
for n in NEWS:
    print(f"  [{n['idx']}] origin={n['origin']} tip={n['tip']} lines={len(n['lines'])} {n['lines'][0][:50]}")

as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = False

placed = []
for n in NEWS:
    gg = next_gg()
    ox, oy = n["origin"]
    tip = n["tip"]
    dx, dy = ox - tmpl_ox, oy - tmpl_oy
    move = pt3(dx, dy, 0)

    tn_c = tmpl_tn.Clone()
    stub_c = mr.GetElementByID64(TMPL_STUB).Clone()
    dim_c = mr.GetElementByID64(TMPL_DIM).Clone()

    tn_c.Move(move)
    stub_c.Move(move)
    dim_c.Move(move)

    tn_c.GraphicGroup = gg
    stub_c.GraphicGroup = gg
    dim_c.GraphicGroup = gg

    mr.AddElement(tn_c)
    tn_c.Rewrite()
    mr.AddElement(stub_c)
    stub_c.Rewrite()
    mr.AddElement(dim_c)
    dim_c.Rewrite()

    tn_id = int(tn_c.ID64)
    stub_id = int(stub_c.ID64)
    dim_id = int(dim_c.ID64)

    # rewrite dim tip → stub end
    stub_el = mr.GetElementByID64(stub_id).AsLineElement
    s0, s1 = stub_el.StartPoint, stub_el.EndPoint
    # stub end = east end (higher X) for stock; after move keep endpoint farther from tip as text-end
    p0 = (float(s0.X), float(s0.Y))
    p1 = (float(s1.X), float(s1.Y))
    if (p0[0] - tip[0]) ** 2 + (p0[1] - tip[1]) ** 2 > (p1[0] - tip[0]) ** 2 + (
        p1[1] - tip[1]
    ) ** 2:
        stub_end = p0
    else:
        stub_end = p1

    dim_el = mr.GetElementByID64(dim_id)
    de = dim_el.AsDimensionElement
    while de.ReferencePointsCount > 0:
        de.DeletePoint(1)
    de.AddReferencePoint(mr, pt3(tip[0], tip[1]))
    de.AddReferencePoint(mr, pt3(stub_end[0], stub_end[1]))
    try:
        de.PrimaryText = "*"
    except Exception:
        pass
    try:
        dim_el.IsHidden = False
    except Exception:
        pass
    dim_el.Rewrite()

    # set text (re-get element)
    set_tn_lines(mr.GetElementByID64(tn_id), n["lines"])

    placed.append(
        {
            "idx": n["idx"],
            "gg": gg,
            "tn": tn_id,
            "stub": stub_id,
            "dim": dim_id,
            "origin": (ox, oy),
            "tip": tip,
            "nlines": len(n["lines"]),
        }
    )
    print(f"  placed [{n['idx']}] gg={gg} tn={tn_id} dim={dim_id}")

as_.GraphicGroupLockEnabled = True
q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)

# ---------------------------------------------------------------------------
# Overlap gate: Range boxes among all TCPP_P_T text nodes
# ---------------------------------------------------------------------------
print("=== OVERLAP GATE ===")
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
texts = []
while ee.MoveNext():
    e = ee.Current
    try:
        if not (e.IsTextElement or e.IsTextNodeElement):
            continue
        if e.Level and e.Level.Name not in ("TCPP_P_T", "TCSH_P_T"):
            continue
        r = e.Range
        lo = (float(r.Low.X), float(r.Low.Y))
        hi = (float(r.High.X), float(r.High.Y))
        if e.IsTextNodeElement:
            tn = e.AsTextNodeElement
            txt = " | ".join(tn.TextLine(i) for i in range(1, int(tn.TextLinesCount) + 1))
        else:
            txt = e.AsTextElement.Text
        texts.append((int(e.ID64), lo, hi, txt[:60]))
    except Exception:
        continue

PAD = 0.25
clashes = []
for i in range(len(texts)):
    for j in range(i + 1, len(texts)):
        a, b = texts[i], texts[j]
        # AABB intersect with pad
        if (
            a[1][0] - PAD <= b[2][0]
            and b[1][0] - PAD <= a[2][0]
            and a[1][1] - PAD <= b[2][1]
            and b[1][1] - PAD <= a[2][1]
        ):
            clashes.append((a[0], b[0], a[3], b[3]))

print(f"scanned {len(texts)} text elems; clashes {len(clashes)}")
for c in clashes[:30]:
    print("  CLASH", c[0], c[1], repr(c[2]), repr(c[3]))

print("=== DONE ===")
print("placed", len(placed), "edited", len(EDITS))
for p in placed:
    print(p)
