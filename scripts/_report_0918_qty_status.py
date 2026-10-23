"""Report live conduit qtys vs remaining 09.18 green qty chips."""
from __future__ import annotations

import re
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
            t = "\n".join(lines).upper()
            if "TRAFFIC SIGNAL CONDUIT" in t or (
                "TRAFFIC SIGNAL CABLE" in t and "680." in t
            ):
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
by_id = {x["id"]: x for x in live}
live = sorted(by_id.values(), key=lambda x: -x["cy"])

print("=== LIVE SUMMARY ===")
for L in live:
    cable = next((ln for ln in L["lines"] if ln.upper().startswith("TRAFFIC SIGNAL CABLE")), "")
    print(f"\nid={L['id']} cy={L['cy']:.1f} {cable}")
    for ln in L["lines"]:
        if any(k in ln.upper() for k in ("CONDUCTORS", "GROUND", "CABLE (")):
            print(" ", ln)

page = fitz.open(PDF)[2]
print("\n=== QTY CHIPS mapped ===")
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "FreeText":
        continue
    if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
        continue
    content = ((a.info or {}).get("content") or "").strip()
    if not (
        re.fullmatch(r"\(\d+\)", content)
        or re.fullmatch(r"\d+", content)
        or re.fullmatch(r"730514 \(\d+\)", content)
        or re.fullmatch(r"730714 \(\d+\)", content)
        or "13 CONDUCTORS" in content.upper()
    ):
        continue
    r = a.rect
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    dx, dy = pdf_to_dgn(cx, cy)
    near = sorted(live, key=lambda L: abs(L["cy"] - dy) + 0.2 * abs(L["cx"] - dx))[0]
    print(
        f"[{i}] {content!r} pdf=({cx:.0f},{cy:.0f}) dgn=({dx:.0f},{dy:.0f}) "
        f"-> nearest id={near['id']} dy={abs(near['cy']-dy):.1f}"
    )
