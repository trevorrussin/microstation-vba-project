"""Inventory live paragraph callouts + PDF FreeText/strike for conduit edits."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
print("DGN", app.ActiveDesignFile.FullName)
print("MODEL", app.ActiveModelReference.Name)
mr = app.ActiveModelReference

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
print("\n=== LIVE callouts (CONDUIT / long TRAFFIC / paragraph) ===")
rows = []
while ee.MoveNext():
    el = ee.Current
    try:
        if not el.IsTextNodeElement:
            continue
        tn = el.AsTextNodeElement
        lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
        t = "\n".join(lines)
        u = t.upper()
        n = len(lines)
        if "CONDUIT" in u or (
            "TRAFFIC" in u and (n >= 3 or len(t) > 80)
        ):
            rng = el.Range
            rows.append(
                (
                    el.ID.Low,
                    float(rng.Low.Y),
                    float(rng.Low.X),
                    n,
                    lines,
                )
            )
    except Exception:
        continue

for rid, ly, lx, n, lines in sorted(rows, key=lambda r: (-r[1], r[2])):
    print(f"\nid={rid} lines={n} xy=({lx:.1f},{ly:.1f})")
    for ln in lines:
        print(" ", repr(ln))

pdf_path = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)
print("\n=== PDF exists", pdf_path.exists(), "===")
if pdf_path.exists():
    import fitz

    doc = fitz.open(pdf_path)
    for pi in range(min(4, len(doc))):
        page = doc[pi]
        print(f"\n--- page {pi+1} FreeText / StrikeOut ---")
        for i, a in enumerate(page.annots() or []):
            at = a.type[1]
            if at not in ("FreeText", "StrikeOut", "Highlight", "Ink", "Squiggly"):
                continue
            content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
            u = content.upper()
            if at == "FreeText" and not content:
                continue
            # focus conduit / long traffic / item nos
            if at == "FreeText":
                if not (
                    "CONDUIT" in u
                    or "TRAFFIC SIGNAL" in u
                    or "MAST ARM" in u
                    or "680." in u
                    or "DELETE" in u
                    or "REMOVE" in u
                    or "REPLACE" in u
                ):
                    continue
            stroke = (a.colors or {}).get("stroke")
            fill = (a.colors or {}).get("fill")
            r = a.rect
            print(f"[{i}] {at} stroke={stroke} fill={fill} pdf=({r.x0:.0f},{r.y0:.0f})")
            if content:
                print(content[:600])
            else:
                # strikeout may have no content — get covered text
                try:
                    words = page.get_text("words")
                    covered = [
                        w[4]
                        for w in words
                        if w[0] < r.x1 and w[2] > r.x0 and w[1] < r.y1 and w[3] > r.y0
                    ]
                    if covered:
                        print("COVERED:", " ".join(covered)[:400])
                except Exception as ex:
                    print("no content", ex)
