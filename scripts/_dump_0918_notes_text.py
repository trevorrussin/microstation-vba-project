"""Dump NOTES-related text from 09.18 BTS PDF."""
from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
doc = fitz.open(PDF)
for pi in range(doc.page_count):
    page = doc[pi]
    text = page.get_text("text")
    u = text.upper()
    if not any(
        k in u
        for k in (
            "ENSURE THAT EACH CONDUIT",
            "REFER TO DWG. TSN",
            "SIDEWALK BOX",
            "MAST ARM FOOTING",
        )
    ):
        continue
    print(f"==== page {pi+1} ====")
    for ln in text.splitlines():
        lu = ln.upper().strip()
        if not lu:
            continue
        if any(
            k in lu
            for k in (
                "NOTE",
                "REFER TO",
                "ENSURE",
                "FURNISH AND INSTALL THAT",
                "CONDUIT DOES NOT",
                "MAST ARM FOOTING",
                "SIDEWALK BOX",
                "AND DETAILS",
            )
        ):
            print(ln[:140])
