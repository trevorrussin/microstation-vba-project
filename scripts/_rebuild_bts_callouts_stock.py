"""Rebuild NEW BTS conduit callouts to match stock anatomy.

Stock recipe (GG 6114–6120):
  TextNode (just=12 RJ) + ~5' horizontal stub toward tip + note-dim
  DimensionType 53, PrimaryText='*', style ny_Plan, level TCPP_P_T
  dim pt1 = tip, pt2 = stub tip-side end
  stub always extends FROM text TOWARD tip
  PLACE NOTE for visible arrowhead
"""
from __future__ import annotations

import math
import sys
import time

import numpy as np

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
adf = app.ActiveDesignFile
q = app.CadInputQueue
as_ = app.ActiveSettings

assert "fea_tsp" in adf.FullName.lower(), adf.FullName

# --- PDF→DGN affine ---
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


# Existing broken NEW callouts to remove
OLD = {
    21: (7001, 7022, 7025, 7026),
    33: (7002, 7027, 7030, 7031),
    59: (7003, 7036, 7039, 7040),
    60: (7004, 7044, 7047, 7048),
    61: (7005, 7053, 7056, 7057),
    64: (7006, 7063, 7066, 7067),
    72: (7007, 7073, 7076, 7077),
    74: (7008, 7084, 7087, 7088),
    75: (7009, 7092, 7095, 7096),
    76: (7010, 7103, 7106, 7107),
    77: (7011, 7113, 7116, 7117),
    80: (7012, 7121, 7124, 7125),
}

# Unique PDF header anchors (right-edge x, y) + full text lines
# Matched FreeText content from r1 BTS page 3
CALLOUTS = [
    {
        "idx": 21,
        "pdf": (540.7, 23.0),
        "lines": ["PULL NEW CABLES AT EXIST. CONDUIT"],
    },
    {
        "idx": 33,
        "pdf": (799.3, 410.2),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (7)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (4)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 59,
        "pdf": (805.1, 351.4),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (2)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 60,
        "pdf": (798.8, 294.6),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (8)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (5)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 61,
        "pdf": (827.7, 230.5),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (2)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (5)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (2) (SEE NOTE 5)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1) (SEE NOTE 5)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 64,
        "pdf": (215.1, 389.7),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (6)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (2)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 72,
        "pdf": (843.1, 111.4),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (2)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (15)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (7)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (3)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (2)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 74,
        "pdf": (166.7, 454.7),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (2)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 75,
        "pdf": (606.0, 705.3),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (2)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (16)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (7)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (3)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (3)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 76,
        "pdf": (944.6, 29.2),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (7)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (3)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (2)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 77,
        "pdf": (724.5, 18.7),
        "lines": [
            "TRAFFIC SIGNAL CONDUIT (1)",
            '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
            "TRAFFIC SIGNAL CABLE (2)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (1)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
        ],
    },
    {
        "idx": 80,
        "pdf": (1104.5, 212.7),
        "lines": [
            "TRAFFIC SIGNAL CABLE (23)",
            "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (10)",
            "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (5)",
            "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (5)",
            "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
            "(GROUND WIRE 10 AWG) (ITEM NO. 670.70011) (1)",
            "(SEE NOTE 6)",
        ],
    },
]


def nearest_tci(xy: tuple[float, float]) -> tuple[float, float]:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    best = None
    best_d = 1e18
    for lv_ok in (("TCI", "TCI_P"), ("UUO",)):
        ee = mr.Scan(o)
        while ee.MoveNext():
            e = ee.Current
            try:
                lv = e.Level.Name if e.Level else ""
                if lv not in lv_ok:
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
        if best is not None and best_d < 80.0**2:
            return best
    return best if best is not None else xy


def next_gg(start: int = 7200) -> int:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    mx = start
    while ee.MoveNext():
        try:
            mx = max(mx, int(ee.Current.GraphicGroup))
        except Exception:
            pass
    return mx + 1


def all_dim_ids() -> set[int]:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = set()
    while ee.MoveNext():
        try:
            if int(ee.Current.Type) == 33:
                out.add(int(ee.Current.ID64))
        except Exception:
            pass
    return out


def newest_dim(before: set[int]) -> int | None:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    best = None
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.Type) != 33:
                continue
            eid = int(e.ID64)
            if eid in before:
                continue
            if best is None or eid > best:
                best = eid
        except Exception:
            pass
    return best


def set_tn_lines(tn_el, lines: list[str]) -> None:
    tn = tn_el.AsTextNodeElement
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    # remove leftover blank by setting space then... keep only needed:
    # if template had more lines, blank extras (stock had 2; we may need 1)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, "")
    tn_el.Rewrite()


# Stock relative offsets (from GG 6119): TN origin → stub start / stub length
# TN (1029437.857, 239729.750)
# stub (1029438.911, 239728.252) → (1029443.911, 239728.252)
STUB_DX0 = 1.054  # stub-start X - TN origin X (toward tip side for east-pointing)
STUB_DY0 = -1.498
STUB_LEN = 5.0

# Template TN to clone (stock APS short — will rewrite lines)
TMPL_TN = 6126

level = adf.Levels("TCPP_P_T")
style = adf.DimensionStyles("ny_Plan")

as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = False

# 1) Delete old broken NEW callouts
print("=== delete old NEW ===")
for idx, (gg, tn, stub, dim) in OLD.items():
    for eid in (tn, stub, dim):
        try:
            mr.RemoveElement(mr.GetElementByID64(eid))
            print(f"  del {eid}")
        except Exception as ex:
            print(f"  del fail {eid}: {ex}")

# 2) Rebuild each
print("=== rebuild ===")
tmpl = mr.GetElementByID64(TMPL_TN)
placed = []

for c in CALLOUTS:
    gg = next_gg()
    ox, oy = pdf_to_dgn(*c["pdf"])
    tip = nearest_tci((ox, oy))

    # Tip side: +1 if tip is east of text origin, else -1 (stub points toward tip)
    tip_dir = 1.0 if tip[0] >= ox else -1.0

    # Stock RJ: origin is at RIGHT of text. For tip east of text (stock), stub
    # extends east from near origin. For tip west of text (right-side callouts),
    # stub extends west; origin stays at the tip-side edge of the text block
    # (which for RJ is still the right edge only when tip is east).
    # When tip is west: place origin so RJ text sits east of stub — i.e. origin
    # near stub at the WEST edge of the would-be text. Easiest match to stock
    # look: keep just=12 and put stub on the tip-ward side of origin.
    stub_start = (ox + tip_dir * abs(STUB_DX0), oy + STUB_DY0)
    stub_end = (stub_start[0] + tip_dir * STUB_LEN, stub_start[1])

    # Clone TN, move to origin
    tn_c = tmpl.Clone()
    cur = tn_c.AsTextNodeElement.Origin
    tn_c.Move(pt3(ox - float(cur.X), oy - float(cur.Y), 0))
    tn_c.GraphicGroup = gg
    mr.AddElement(tn_c)
    tn_c.Rewrite()
    tn_id = int(tn_c.ID64)
    set_tn_lines(mr.GetElementByID64(tn_id), c["lines"])

    # Create stub line toward tip (CreateLineElement1 returns (element, pts))
    stub_el = app.CreateLineElement1(None, (pt3(*stub_start), pt3(*stub_end)))
    if isinstance(stub_el, tuple):
        stub_el = stub_el[0]
    stub_el.Level = level
    stub_el.Color = 3
    stub_el.LineWeight = 0
    stub_el.GraphicGroup = gg
    mr.AddElement(stub_el)
    stub_el.Rewrite()
    stub_id = int(stub_el.ID64)

    # PLACE NOTE tip → stub_end (stock recipe for arrowhead)
    before = all_dim_ids()
    q.SendKeyin("NULL")
    time.sleep(0.05)
    q.SendKeyin("PLACE NOTE")
    time.sleep(0.1)
    q.SendDataPoint(pt3(*tip), 1)
    time.sleep(0.05)
    q.SendDataPoint(pt3(*stub_end), 1)
    time.sleep(0.25)
    q.SendReset()
    time.sleep(0.15)

    dim_id = newest_dim(before)
    if dim_id is None:
        print(f"  FAIL no dim [{c['idx']}]")
        placed.append({"idx": c["idx"], "fail": True})
        continue

    dim_el = mr.GetElementByID64(dim_id)
    de = dim_el.AsDimensionElement
    try:
        de.PrimaryText = "*"
    except Exception as ex:
        print("  PrimaryText", ex)
    try:
        de.DimensionStyle = style
    except Exception as ex:
        print("  style", ex)
    try:
        dim_el.Level = level
    except Exception:
        pass
    try:
        dim_el.Color = 3
    except Exception:
        pass
    try:
        dim_el.IsHidden = False
    except Exception:
        pass
    dim_el.GraphicGroup = gg
    dim_el.Rewrite()

    # verify points; rewrite if PLACE NOTE swapped order
    de = mr.GetElementByID64(dim_id).AsDimensionElement
    p1 = (float(de.Point(1).X), float(de.Point(1).Y))
    p2 = (float(de.Point(2).X), float(de.Point(2).Y))
    # ensure pt1≈tip, pt2≈stub_end
    d_tip_p1 = (p1[0] - tip[0]) ** 2 + (p1[1] - tip[1]) ** 2
    d_tip_p2 = (p2[0] - tip[0]) ** 2 + (p2[1] - tip[1]) ** 2
    if d_tip_p2 < d_tip_p1:
        while de.ReferencePointsCount > 0:
            de.DeletePoint(1)
        de.AddReferencePoint(mr, pt3(*tip))
        de.AddReferencePoint(mr, pt3(*stub_end))
        try:
            de.PrimaryText = "*"
        except Exception:
            pass
        dim_el = mr.GetElementByID64(dim_id)
        dim_el.Rewrite()

    placed.append(
        {
            "idx": c["idx"],
            "gg": gg,
            "tn": tn_id,
            "stub": stub_id,
            "dim": dim_id,
            "origin": (ox, oy),
            "tip": tip,
            "stub_end": stub_end,
            "tip_dir": tip_dir,
            "nlines": len(c["lines"]),
        }
    )
    print(
        f"  [{c['idx']}] gg={gg} tn={tn_id} stub={stub_id} dim={dim_id} "
        f"dir={tip_dir:+.0f} tip=({tip[0]:.1f},{tip[1]:.1f})"
    )

as_.GraphicGroupLockEnabled = True
as_.AssociationLockEnabled = False
q.SendKeyin("NULL")
q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)

# Overlap gate
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
            txt = " | ".join(
                tn.TextLine(i) for i in range(1, int(tn.TextLinesCount) + 1)
            )[:50]
        else:
            txt = e.AsTextElement.Text[:50]
        texts.append((int(e.ID64), lo, hi, txt))
    except Exception:
        continue

PAD = 0.25
clashes = []
for i in range(len(texts)):
    for j in range(i + 1, len(texts)):
        a, b = texts[i], texts[j]
        if (
            a[1][0] - PAD <= b[2][0]
            and b[1][0] - PAD <= a[2][0]
            and a[1][1] - PAD <= b[2][1]
            and b[1][1] - PAD <= a[2][1]
        ):
            clashes.append((a[0], b[0], a[3], b[3]))

print(f"scanned {len(texts)}; clashes {len(clashes)}")
for c in clashes:
    print(" CLASH", c)

print("=== DONE ===")
for p in placed:
    print(p)
