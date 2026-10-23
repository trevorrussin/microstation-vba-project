from pathlib import Path
import fitz

p = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
print("exists", p.exists(), p)
doc = fitz.open(p)
print("pages", doc.page_count)
for pi in range(doc.page_count):
    page = doc[pi]
    annos = list(page.annots() or [])
    types = {}
    for a in annos:
        types[a.type[1]] = types.get(a.type[1], 0) + 1
    green = sum(
        1
        for a in annos
        if a.type[1] == "FreeText" and (a.colors or {}).get("stroke") == [0.0, 1.0, 0.0]
    )
    words = page.get_text("words")
    redcov = 0
    red_samples = []
    for a in annos:
        stroke = (a.colors or {}).get("stroke") or []
        if not (stroke and stroke[0] > 0.7 and stroke[1] < 0.35):
            continue
        r = a.rect
        cov = [
            w[4]
            for w in words
            if w[0] < r.x1 and w[2] > r.x0 and w[1] < r.y1 and w[3] > r.y0
        ]
        if cov:
            redcov += 1
            if len(red_samples) < 25:
                red_samples.append((a.type[1], " ".join(cov)[:120], (r.x0, r.y0)))
    print(f"page {pi+1}: annots={len(annos)} {types} greenFT={green} redCoveringWords={redcov}")
    for s in red_samples:
        print(" ", s)

print("\n=== green FreeTexts with CONDUIT or ITEM fragments page-by-page ===")
for pi in range(doc.page_count):
    page = doc[pi]
    print(f"\n--- page {pi+1} ---")
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        u = content.upper()
        if not any(
            k in u
            for k in (
                "CONDUIT",
                "CABLE",
                "680.",
                "APS",
                "MAST",
                "STREETLIGHT",
                "VEHICULAR",
                "ITEM",
            )
        ):
            continue
        print(f"[{i}] {content[:300]!r}")
