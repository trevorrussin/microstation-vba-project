"""Find NOTES content in r1 and 09.18 BTS PDFs."""
from pathlib import Path
import fitz

paths = [
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf",
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf",
]
for PDF in paths:
    if not PDF.exists():
        print("MISSING", PDF)
        continue
    doc = fitz.open(PDF)
    print("====", PDF.name, "pages", doc.page_count)
    for pi in range(doc.page_count):
        page = doc[pi]
        text = page.get_text("text")
        if "ENSURE THAT EACH CONDUIT" in text.upper() or "REFER TO DWG" in text.upper():
            print(f"  base text hit on page {pi+1}")
            for ln in text.splitlines():
                lu = ln.upper()
                if any(
                    k in lu
                    for k in (
                        "REFER TO",
                        "ENSURE",
                        "FOOTING",
                        "SIDEWALK",
                        "AND DETAILS",
                        "1.",
                        "2.",
                        "3.",
                        "4.",
                    )
                ):
                    if len(ln.strip()) > 5:
                        print("   ", ln[:130])
        for i, a in enumerate(page.annots() or []):
            c = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
            cu = c.upper()
            if any(
                k in cu
                for k in (
                    "ENSURE",
                    "AND DETAILS",
                    "SIDEWALK BOX",
                    "MAST ARM FOOTING",
                    "LEFT JUSTIFY",
                    "REFER TO DWG",
                    "REFER TO TSP",
                    "REFER TO NYCDOT",
                )
            ):
                print(f"  p{pi+1}[{i}] {c[:200]!r}")
