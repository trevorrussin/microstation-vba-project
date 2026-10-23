"""Audit red-line small deletions + green EDIT chips vs live callout text."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import fitz

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)

# Known EDIT callout hosts from prior scripts (non-conduit)
HOSTS = [6132, 5474, 6120, 6126, 6129, 6280, 6117, 3770, 6214, 6135]


def walk_tns(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines)
            u = t.upper()
            # skip conduit paragraphs
            if "TRAFFIC SIGNAL CONDUIT" in u:
                return
            if u.startswith("TRAFFIC SIGNAL CABLE") and "680.731314" in u:
                return
            # keep plan callouts with ITEM / APS / STREETLIGHT / MAST / VEHICULAR / PUSH
            keys = (
                "ITEM NO",
                "APS",
                "STREETLIGHT",
                "MAST ARM",
                "VEHICULAR",
                "PUSH BUTTON",
                "CONTROL CABINET",
                "S-1 POLE",
                "DUAL 25",
            )
            if any(k in u for k in keys):
                rng = el.Range
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": lines,
                        "text": t,
                        "y": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "x": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk_tns(ee.Current, out)
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    live = []
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    while ee.MoveNext():
        walk_tns(ee.Current, live)
    by_id = {x["id"]: x for x in live}
    live = list(by_id.values())

    print("\n=== LIVE non-conduit callouts ===")
    for L in sorted(live, key=lambda x: -x["y"]):
        print(f"\nid={L['id']} @({L['x']:.0f},{L['y']:.0f})")
        for ln in L["lines"]:
            print(" ", repr(ln))

    page = fitz.open(PDF)[2]
    words = page.get_text("words")

    print("\n=== RED marks covering small words/numbers ===")
    red_hits = []
    for i, a in enumerate(page.annots() or []):
        stroke = (a.colors or {}).get("stroke") or []
        fill = (a.colors or {}).get("fill") or []
        is_red = False
        for c in (stroke, fill):
            if c and len(c) >= 3 and c[0] > 0.7 and c[1] < 0.35 and c[2] < 0.35:
                is_red = True
        if not is_red:
            continue
        if a.type[1] in ("Ink", "Square") and a.rect.width < 40 and a.rect.height < 40:
            # skip x/y/+ chip marks
            content = ((a.info or {}).get("content") or "").strip()
            if content in ("x", "y", "+", "-", ".", ""):
                continue
        r = a.rect
        pad = 2
        covered = [
            w[4]
            for w in words
            if w[0] < r.x1 + pad
            and w[2] > r.x0 - pad
            and w[1] < r.y1 + pad
            and w[3] > r.y0 - pad
        ]
        if not covered:
            continue
        cov = " ".join(covered)
        # keep if looks like item/qty fragment
        if not re.search(r"680\.|\(\d\)|ITEM|SEE NOTE", cov, re.I):
            if len(cov) > 40:
                continue
        red_hits.append((i, a.type[1], cov, r))
        print(f"[{i}] {a.type[1]} covers: {cov!r} @({r.x0:.0f},{r.y0:.0f})")

    print("\n=== GREEN FreeText EDIT chips (not full conduit paragraphs) ===")
    greens = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        u = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" in u or u.startswith("TRAFFIC SIGNAL CABLE"):
            continue
        if content.startswith("5.") or content.startswith("6."):
            continue
        # table rows
        if " EA " in u and u.strip()[0].isdigit():
            continue
        r = a.rect
        greens.append((i, content, r))
        print(f"[{i}] {content!r} @({r.x0:.0f},{r.y0:.0f})")

    # Expected FINAL from finish script — check live
    FINAL = {
        6132: [
            "VEHICULAR SIGNAL HEADS AND APS PUSH BUTTON",
            "(ITEM NOS. 680.818122011 (1), 680.81822011 (1), 680.8223 (1), 680.8151 (1))",
        ],
        5474: [
            "S-1 POLE WITH VEHICULAR HEAD ",
            "(ITEM NOS. 680.5001, 680.20720011 (1), 680.8223 (1),",
            "680.818122011 (1), 680.81822011 (1))",
            "(SEE NOTE 3)",
        ],
        6120: [
            "STREETLIGHT W/PEDESTRIAN TRAFFIC SIGNALS AND APS PUSH BUTTON",
            "(ITEM NOS. 680.8141 (2), 680.815002 (2), 680.8151 (1), 680.83080011 (1))",
        ],
        6126: [
            "APS W/POST",
            "(ITEM NOS. 680.5001, 680.8152 (1))",
        ],
        6129: [
            "APS W/POST",
            "(ITEM NOS. 680.5001, 680.8152 (1))",
        ],
        # 6280 was swapped to CONTROL CABINET by engineer request
        6280: [
            "APS PUSH BUTTON ON ",
            "EXIST. CONTROL CABINET",
            "(ITEM NO. 680.8151)",
        ],
        6117: [
            "APS W/POST",
            "(ITEM NOS. 680.5001,",
            " 680.8152 (1))",
        ],
        3770: [
            "25' TRAFFIC SIGNAL MAST ARM WITH VEHICULAR",
            "AND PEDESTRIAN TRAFFIC SIGNALS",
            "AND APS PUSH BUTTON",
            "(ITEM NOS. 680.5001, 680.8141 (2), 680.815002 (2),",
            "680.8151 (1), 680.818122011 (4), 680.81822011 (4),",
            "680.82050011, 680.8223 (1), 680.82310011, 680.83720011)",
            "(SEE NOTE 3)",
        ],
        6214: [
            "DUAL 25' TRAFFIC SIGNAL MAST ARM WITH VEHICULAR HEADS",
            "(ITEM NOS. 680.5001, 680.818122011 (5), 680.81822011 (5),",
            "680.82050011, 680.8223, 680.82310011, 680.83720011(2))",
            "(SEE NOTE 3)",
        ],
    }

    print("\n=== EXPECTED FINAL vs LIVE ===")
    for hid, want in FINAL.items():
        L = by_id.get(hid)
        if not L:
            print(f"MISSING host {hid}")
            continue
        got = [ln for ln in L["lines"] if ln.strip() and ln.strip() != " "]
        want_n = [ln for ln in want if ln.strip()]
        ok = got == want_n
        print(f"\n{'OK' if ok else 'DIFF'} id={hid}")
        if not ok:
            print("  LIVE:")
            for ln in got:
                print("   ", repr(ln))
            print("  WANT:")
            for ln in want_n:
                print("   ", repr(ln))


if __name__ == "__main__":
    main()
