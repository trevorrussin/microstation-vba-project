"""Extract all NOTES-area text from BTS markup PDFs (r1 + 09.18)."""
from pathlib import Path
import fitz

paths = [
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf",
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf",
]

for PDF in paths:
    if not PDF.exists():
        print("MISSING", PDF.name)
        continue
    doc = fitz.open(PDF)
    print("=" * 60, PDF.name)
    for pi in range(doc.page_count):
        page = doc[pi]
        # All FreeTexts
        for i, a in enumerate(page.annots() or []):
            if a.type[1] != "FreeText":
                continue
            c = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
            if not c:
                continue
            r = a.rect
            # left side of page often notes; also any note-like content
            cu = c.upper()
            leftish = r.x0 < 400
            noteish = any(
                k in cu
                for k in (
                    "NOTE",
                    "REFER",
                    "ENSURE",
                    "FOOTING",
                    "SIDEWALK",
                    "CONDUIT",
                    "DETAILS",
                    "FURNISH AND INSTALL THAT",
                    "LEFT JUSTIFY",
                )
            )
            if leftish or noteish:
                print(f"  p{pi+1}[{i}] @({r.x0:.0f},{r.y0:.0f}) {c[:250]!r}")

        # Base text blocks on left third
        blocks = page.get_text("blocks")
        print(f"  --- page {pi+1} left-ish blocks ---")
        for b in blocks:
            x0, y0, x1, y1, text, *_ = b
            if x0 > page.rect.width * 0.45:
                continue
            t = (text or "").strip()
            if not t:
                continue
            if any(
                k in t.upper()
                for k in (
                    "REFER",
                    "ENSURE",
                    "FOOTING",
                    "SIDEWALK",
                    "NOTES",
                    "CONDUIT",
                    "FURNISH",
                    "DETAIL",
                    "STANDARD",
                )
            ) or (t[:2].rstrip(".").isdigit() and len(t) > 10):
                print(f"  block @({x0:.0f},{y0:.0f}) {t[:300]!r}")
