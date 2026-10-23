"""Dump live non-conduit callouts + nearby 09.18 red qty strikes."""
from __future__ import annotations

import sys
from pathlib import Path

import fitz
import numpy as np

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
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


def pdf_to_dgn(px, py):
    return float(_AX[0] + _AX[1] * px + _AX[2] * py), float(
        _AY[0] + _AY[1] * px + _AY[2] * py
    )


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines)
            u = t.upper()
            if "TRAFFIC SIGNAL CONDUIT" in u:
                return
            if u.startswith("TRAFFIC SIGNAL CABLE") and "680." in u:
                return
            keys = (
                "ITEM NO",
                "APS",
                "STREETLIGHT",
                "MAST ARM",
                "VEHICULAR",
                "PUSH BUTTON",
                "CONTROL CABINET",
                "DUAL 25",
                "PULLBOX",
                "PULL NEW",
            )
            if any(k in u for k in keys):
                rng = el.Range
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": lines,
                        "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
live = []
while ee.MoveNext():
    walk(ee.Current, live)
by = {x["id"]: x for x in live}
live = sorted(by.values(), key=lambda x: -x["cy"])

print("=== LIVE non-conduit callouts ===")
for L in live:
    print(f"\nid={L['id']} cy={L['cy']:.1f} cx={L['cx']:.0f}")
    for ln in L["lines"]:
        print(f"  {ln}")

page = fitz.open(PDF)[2]
words = page.get_text("words")
print("\n=== ALL red Line/Ink covering short tokens ===")
for i, a in enumerate(page.annots() or []):
    stroke = (a.colors or {}).get("stroke") or []
    fill = (a.colors or {}).get("fill") or []
    is_red = False
    for c in (stroke, fill):
        if c and len(c) >= 3 and c[0] > 0.7 and c[1] < 0.35 and c[2] < 0.35:
            is_red = True
    if not is_red:
        continue
    if a.type[1] not in ("Line", "Ink", "PolyLine", "Square", "Circle", "Highlight"):
        continue
    r = a.rect
    cov = sorted(
        [
            w
            for w in words
            if w[0] < r.x1 and w[2] > r.x0 and w[1] < r.y1 and w[3] > r.y0
        ],
        key=lambda w: (round(w[1], 0), w[0]),
    )
    if not cov:
        continue
    txt = " ".join(w[4] for w in cov)
    # keep qty / item-ish short strikes
    if len(txt) > 80:
        continue
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    dx, dy = pdf_to_dgn(cx, cy)
    near = sorted(live, key=lambda L: abs(L["cy"] - dy) + 0.3 * abs(L["cx"] - dx))
    nid = near[0]["id"] if near else None
    ndy = abs(near[0]["cy"] - dy) if near else None
    print(
        f"[{i}] {a.type[1]} strike={txt!r} dgn=({dx:.0f},{dy:.0f}) "
        f"near={nid} dy={ndy:.1f}"
    )
