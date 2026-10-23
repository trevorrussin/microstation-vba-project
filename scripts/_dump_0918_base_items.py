"""Show base PDF callout lines containing ITEM NO on 09.18 page 3."""
from pathlib import Path
import fitz

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
page = fitz.open(PDF)[2]
blocks = page.get_text("dict")["blocks"]
n = 0
for b in blocks:
    if b.get("type") != 0:
        continue
    lines = []
    for line in b.get("lines", []):
        t = "".join(s["text"] for s in line.get("spans", []))
        if t.strip():
            lines.append(t.strip())
    text = "\n".join(lines)
    u = text.upper()
    if "ITEM NO" in u and (
        "CONDUIT" in u or "CABLE" in u or "GROUND" in u or "CONDUCTOR" in u
    ):
        n += 1
        bb = b["bbox"]
        print(f"\n--- block @({bb[0]:.0f},{bb[1]:.0f}) ---")
        print(text)
print(f"\nblocks={n}")

# Also raw lines containing 670. or 680.73
print("\n=== all lines with 670. or 680.73 or 731 or 7507 ===")
for b in blocks:
    if b.get("type") != 0:
        continue
    for line in b.get("lines", []):
        t = "".join(s["text"] for s in line.get("spans", [])).strip()
        if any(k in t for k in ("670.", "680.73", "731", "7507", "7305", "7307")):
            bb = line["bbox"]
            print(f"({bb[0]:.0f},{bb[1]:.0f}) {t!r}")
