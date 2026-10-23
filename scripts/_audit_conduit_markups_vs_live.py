"""Compare: PDF original sheet text vs strikeouts vs green FreeText vs live DGN."""
from __future__ import annotations

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


def norm(s: str) -> str:
    return " ".join(s.replace("\r", "\n").split())


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines)
            u = t.upper()
            if "TRAFFIC SIGNAL CONDUIT" in u or (
                "TRAFFIC SIGNAL CABLE" in u and "680." in u
            ):
                rng = el.Range
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": [ln for ln in lines if ln.strip()],
                        "y": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "x": (float(rng.Low.X) + float(rng.High.X)) / 2,
                        "text": t,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
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
        walk(ee.Current, live)
    by_id = {x["id"]: x for x in live}
    live = list(by_id.values())

    doc = fitz.open(PDF)
    page = doc[2]
    words = page.get_text("words")  # x0,y0,x1,y1,word,block,line,word_no

    print("\n=== ALL annot types on page 3 (conduit-related area) ===")
    for i, a in enumerate(page.annots() or []):
        at = a.type[1]
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        r = a.rect
        stroke = (a.colors or {}).get("stroke")
        # show strikeouts always; FreeText if conduit/cable
        u = content.upper()
        if at in ("StrikeOut", "Highlight", "Squiggly", "Underline", "Ink"):
            covered = [
                w[4]
                for w in words
                if w[0] < r.x1
                and w[2] > r.x0
                and w[1] < r.y1
                and w[3] > r.y0
            ]
            print(f"\n[{i}] {at} stroke={stroke} pdf=({r.x0:.0f},{r.y0:.0f})-({r.x1:.0f},{r.y1:.0f})")
            if content:
                print(" content:", content[:200])
            if covered:
                print(" covered:", " ".join(covered)[:300])
        elif at == "FreeText" and (
            "TRAFFIC SIGNAL CONDUIT" in u
            or u.startswith("TRAFFIC SIGNAL CABLE")
            or "CONDUIT" in u
            and "CABLE" in u
        ):
            print(f"\n[{i}] FreeText GREEN? stroke={stroke} pdf=({r.x0:.0f},{r.y0:.0f})")
            print(content[:500])

    print("\n\n=== LIVE vs nearest green FreeText (full dump) ===")
    greens = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        u = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" not in u and not u.startswith("TRAFFIC SIGNAL CABLE"):
            continue
        lines = [ln for ln in content.splitlines() if ln.strip()]
        greens.append({"idx": i, "lines": lines, "rect": a.rect, "text": "\n".join(lines)})

    for L in sorted(live, key=lambda x: -x["y"]):
        print(f"\n----- LIVE id={L['id']} -----")
        for ln in L["lines"]:
            print(" ", repr(ln))
        # find green with same first line + best conductor overlap
        best = None
        for G in greens:
            if L["lines"] == G["lines"]:
                best = (999, G)
                break
            score = 0
            if L["lines"] and G["lines"] and L["lines"][0] == G["lines"][0]:
                score += 5
            if len(L["lines"]) > 2 and len(G["lines"]) > 2:
                if L["lines"][2] == G["lines"][2]:
                    score += 5
            score += len(set(L["lines"]) & set(G["lines"]))
            if best is None or score > best[0]:
                best = (score, G)
        if best:
            G = best[1]
            same = L["lines"] == G["lines"]
            print(f"  nearest green [{G['idx']}] score={best[0]} SAME={same}")
            if not same:
                print("  GREEN:")
                for ln in G["lines"]:
                    print("   ", repr(ln))
                print("  ONLY IN LIVE:", [x for x in L["lines"] if x not in G["lines"]])
                print("  ONLY IN GREEN:", [x for x in G["lines"] if x not in L["lines"]])

    # Also: extract native PDF text blocks that look like old conduit callouts
    print("\n\n=== PDF page-3 TEXT BLOCKS containing CONDUIT/CABLE (base drawing) ===")
    blocks = page.get_text("dict")["blocks"]
    n = 0
    for b in blocks:
        if b.get("type") != 0:
            continue
        text = ""
        for line in b.get("lines", []):
            line_t = "".join(s["text"] for s in line.get("spans", []))
            text += line_t + "\n"
        u = text.upper()
        if "TRAFFIC SIGNAL CONDUIT" in u or (
            "TRAFFIC SIGNAL CABLE" in u and "CONDUCTOR" in u
        ):
            n += 1
            bb = b["bbox"]
            print(f"\nblock@{bb[0]:.0f},{bb[1]:.0f}:")
            print(text.strip()[:600])
    print(f"\nbase drawing conduit/cable blocks: {n}")


if __name__ == "__main__":
    main()
