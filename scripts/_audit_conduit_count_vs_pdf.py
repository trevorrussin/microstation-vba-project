"""Count printed TRAFFIC SIGNAL CONDUIT blocks on 09.18 p3 vs live fea_tsp."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import fitz

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF18 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
PDF_R1 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines).upper()
            if "TRAFFIC SIGNAL CONDUIT" in t or (
                "TRAFFIC SIGNAL CABLE" in t and "680." in t
            ):
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": [ln.rstrip() for ln in lines if ln.strip()],
                        "has_conduit": "TRAFFIC SIGNAL CONDUIT" in t,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def blocks_from_page(page):
    """Group words into text blocks that look like conduit/cable callouts."""
    blocks = page.get_text("blocks")
    found = []
    for b in blocks:
        # b: x0,y0,x1,y1,text,block_no,block_type
        if len(b) < 5:
            continue
        text = (b[4] or "").replace("\r", "\n")
        u = text.upper()
        if "TRAFFIC SIGNAL CONDUIT" in u or (
            "TRAFFIC SIGNAL CABLE" in u and ("680." in u or "CONDUCTOR" in u)
        ):
            # skip estimate table / pay item lists
            if "LF" in u and "ITEM" in u and "QUANTITY" in u:
                continue
            if "ESTIMATE" in u:
                continue
            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
            found.append({"y": (b[1] + b[3]) / 2, "x": (b[0] + b[2]) / 2, "lines": lines})
    return found


def freetexts(page):
    out = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        cu = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" not in cu and "TRAFFIC SIGNAL CABLE" not in cu:
            continue
        if "SINGLE CONDUCTOR CABLE 10 GAGE" in cu and "LF" in cu:
            continue
        lines = [ln.strip() for ln in content.split("\n") if ln.strip()]
        out.append({"i": i, "lines": lines})
    return out


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
live = list(by.values())

print(f"LIVE total cable/conduit nodes: {len(live)}")
print(f"  with CONDUIT header: {sum(1 for x in live if x['has_conduit'])}")
print(f"  cable-only: {sum(1 for x in live if not x['has_conduit'])}")
for x in sorted(live, key=lambda z: z["id"]):
    head = x["lines"][0] if x["lines"] else ""
    print(f"  id={x['id']} conduit={x['has_conduit']} | {head}")

page18 = fitz.open(PDF18)[2]
printed = blocks_from_page(page18)
print(f"\n09.18 printed CONDUIT/CABLE text blocks: {len(printed)}")
for i, b in enumerate(sorted(printed, key=lambda z: z["y"])):
    head = b["lines"][0] if b["lines"] else ""
    print(f"  [{i}] y={b['y']:.0f} | {head[:70]}")

ft18 = freetexts(page18)
print(f"\n09.18 FreeText CONDUIT/CABLE replacements: {len(ft18)}")

if PDF_R1.exists():
    page_r1 = fitz.open(PDF_R1)[2]
    ft_r1 = freetexts(page_r1)
    print(f"\nr1 FreeText CONDUIT/CABLE replacements: {len(ft_r1)}")
    for M in ft_r1:
        head = M["lines"][0] if M["lines"] else ""
        # fingerprint cable qty
        cable = None
        for ln in M["lines"]:
            m = re.search(r"CABLE\s*\((\d+)\)", ln.upper())
            if m:
                cable = int(m.group(1))
                break
        print(f"  [{M['i']}] cable={cable} | {head[:60]}")

    # Which r1 FreeTexts have no live fingerprint match on cable qty + first conductor qtys
    def sig(lines):
        cable = None
        parts = []
        for ln in lines:
            u = ln.upper()
            m = re.search(r"CABLE\s*\((\d+)\)", u)
            if m:
                cable = int(m.group(1))
            if "CONDUCTORS" in u or "GROUND" in u or "SINGLE CONDUCTOR" in u:
                qm = re.search(r"\((\d+)\)\s*(\(SEE NOTE \d+\))?$", u)
                if qm:
                    parts.append(int(qm.group(1)))
        return (cable, tuple(parts))

    live_sigs = {sig(x["lines"]) for x in live}
    print("\n=== r1 FreeTexts with no live signature match ===")
    for M in ft_r1:
        s = sig(M["lines"])
        if s not in live_sigs:
            print(f"  MISSING live twin [{M['i']}] sig={s}")
            for ln in M["lines"][:7]:
                print(f"    {ln}")
else:
    print("\nr1 PDF not found:", PDF_R1)
