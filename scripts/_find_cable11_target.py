"""Find callout for CABLE(6)->(11) markup match; show live + markup text."""
from __future__ import annotations

import sys
from pathlib import Path

import fitz

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF_R1 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
PDF18 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"


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
                        "el": el,
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
print("DGN", app.ActiveDesignFile.FullName)
mr = app.ActiveModelReference
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
live = []
while ee.MoveNext():
    walk(ee.Current, live)
by = {x["id"]: x for x in live}
live = sorted(by.values(), key=lambda x: -x["cy"])

print("=== LIVE top->bottom ===")
for i, L in enumerate(live, 1):
    cable = next((ln for ln in L["lines"] if "CABLE" in ln.upper()), "")
    print(f"#{i} id={L['id']} cy={L['cy']:.1f} {cable.strip()!r}")
    for ln in L["lines"]:
        print(f"    {ln}")

print("\n=== r1 FreeText [64] ===")
page = fitz.open(PDF_R1)[2]
for i, a in enumerate(page.annots() or []):
    if i != 64:
        continue
    print(((a.info or {}).get("content") or "").replace("\r", "\n"))

print("\n=== 09.18 FreeTexts with 11 / CABLE / (6) ===")
page18 = fitz.open(PDF18)[2]
for i, a in enumerate(page18.annots() or []):
    if a.type[1] != "FreeText":
        continue
    c = ((a.info or {}).get("content") or "").strip()
    cu = c.upper()
    if (
        "11" in c
        or "(6)" in c
        or "CABLE" in cu
        or "CONDUIT" in cu
    ):
        if len(c) > 200 and "CONDUIT" not in cu and "CABLE" not in cu:
            continue
        r = a.rect
        print(f"[{i}] @({r.x0:.0f},{r.y0:.0f}) {c[:180]!r}")
