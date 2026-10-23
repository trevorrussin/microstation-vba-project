import fitz
from pathlib import Path

pdfs = [
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"),
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Bally's_Hutch_09.11.26_Traffic_BTS.pdf"),
    Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.09.26_Traffic_BTS.pdf"),
]

for p in pdfs:
    if not p.exists():
        print("MISSING", p)
        continue
    doc = fitz.open(p)
    print("===", p.name, "pages", doc.page_count)
    if doc.page_count < 3:
        print("  <3 pages")
        continue
    page = doc[2]  # third page (0-based)
    print("  page3 size", page.rect)
    annos = list(page.annots() or [])
    print("  annots", len(annos))
    for i, a in enumerate(annos):
        colors = a.colors or {}
        info = a.info or {}
        content = (info.get("content") or info.get("Content") or "").strip()
        print(
            f"  [{i}] type={a.type} stroke={colors.get('stroke')} "
            f"fill={colors.get('fill')} rect={tuple(round(x, 1) for x in a.rect)}"
        )
        if content:
            print("      content:", repr(content[:400]))
        # Also try get_text from annot
        try:
            t = a.get_text()
            if t and t.strip() and t.strip() != content:
                print("      get_text:", repr(t.strip()[:400]))
        except Exception:
            pass
