"""List page-3 callout work items from r1 BTS PDF (new vs edit)."""
from __future__ import annotations

import fitz
from pathlib import Path

pdf = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf")
page = fitz.open(pdf)[2]

print("=== NEW (conduit / pull cables) ===")
print("=== EDIT (item inserts on existing) ===")
for i, a in enumerate(page.annots() or []):
    if a.type[1] != "FreeText":
        continue
    stroke = (a.colors or {}).get("stroke")
    if stroke != [0.0, 1.0, 0.0]:
        continue
    content = (a.info or {}).get("content") or ""
    content = content.replace("\r", "\n").strip()
    u = content.upper()
    r = a.rect
    # skip table / notes
    if content.startswith("5.") or content.startswith("6."):
        continue
    if " EA " in u or (u.startswith("680.") and "LF" in u and "CONDUIT" not in u and "(" not in content.split("\n")[0]):
        # pay item table rows
        if "SIGNAL CABLE 5 CONDUCTORS" in u or "SIGNAL CABLE 7 CONDUCTORS" in u:
            continue
        if "FURNISH 3 SECTION" in u or "SHAFT EXTENSION" in u and "EA" in u:
            continue
    is_new = (
        "TRAFFIC SIGNAL CONDUIT" in u
        or "TRAFFIC SIGNAL CABLE (" in u
        or "PULL NEW CABLES" in u
    )
    # short item edits
    is_edit = (
        not is_new
        and (
            "680.818" in u
            or "680.5001" in u
            or "680.822" in u
            or "680.8308" in u
            or content.strip() in ("(1)", "(2)", "5")
        )
    )
    if not is_new and not is_edit:
        continue
    kind = "NEW" if is_new else "EDIT"
    print(f"\n[{i}] {kind} pdf=({r.x0:.0f},{r.y0:.0f})-({r.x1:.0f},{r.y1:.0f})")
    print(content[:500])
