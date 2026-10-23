"""Find NOTES FreeTexts / words in 09.18 BTS PDF."""
from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
doc = fitz.open(PDF)
print("pages", doc.page_count)
keys = (
    "ENSURE",
    "AND DETAILS",
    "SIDEWALK BOX",
    "MAST ARM FOOTING",
    "LEFT JUSTIFY",
    "TSN-01",
    "MORE THAN 9",
)
for pi in range(doc.page_count):
    page = doc[pi]
    for i, a in enumerate(page.annots() or []):
        c = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        cu = c.upper()
        if any(k in cu for k in keys) or (
            "NOTE" in cu and len(c) > 20 and a.type[1] == "FreeText"
        ):
            print(f"p{pi+1}[{i}] {a.type[1]} {c[:220]!r}")
    for w in ("ENSURE", "TSN-01", "SIDEWALK", "FOOTING", "DETAILS", "NOTES"):
        hits = page.search_for(w)
        if hits:
            print(f"p{pi+1} word {w}: {len(hits)}")
