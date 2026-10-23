"""Compare 09.18 page-3 conduit FreeTexts vs live fea_tsp TRAFFIC SIGNAL CONDUIT callouts."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import fitz
import numpy as np

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"
ANCHORS = [
    (662.6, 453.2, 1029505.95, 239730.47),
    (649.1, 468.6, 1029501.96, 239722.75),
    (633.7, 468.6, 1029494.05, 239721.88),
    (609.8, 363.9, 1029480.91, 239776.98),
    (624.4, 427.7, 1029488.63, 239743.33),
    (625.7, 443.1, 1029490.00, 239734.42),
]
_A = np.array([[1, px, py] for px, py, _, _ in ANCHORS], float)
_AX = np.linalg.lstsq(_A, np.array([dx for _, _, dx, _ in ANCHORS]), rcond=None)[0]
_AY = np.linalg.lstsq(_A, np.array([dy for _, _, _, dy in ANCHORS]), rcond=None)[0]


def pdf_to_dgn(px, py):
    return float(_AX[0] + _AX[1] * px + _AX[2] * py), float(
        _AY[0] + _AY[1] * px + _AY[2] * py
    )


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines).upper()
            if "TRAFFIC SIGNAL CONDUIT" in t or (
                "TRAFFIC SIGNAL CABLE" in t and "680." in t
            ):
                rng = el.Range
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": [ln.rstrip() for ln in lines],
                        "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def fp(lines):
    """Fingerprint: conduit qty, cable qty, conductor qty tuple."""
    cduit = cable = None
    parts = []
    for ln in lines:
        u = ln.upper().strip()
        m = re.search(r"TRAFFIC SIGNAL CONDUIT\s*\((\d+)\)", u)
        if m:
            cduit = int(m.group(1))
        m = re.search(r"TRAFFIC SIGNAL CABLE\s*\((\d+)\)", u)
        if m:
            cable = int(m.group(1))
        if "CONDUCTORS" in u or "GROUND WIRE" in u or "SINGLE CONDUCTOR" in u:
            qm = re.search(r"\((\d+)\)\s*(\(SEE NOTE \d+\))?$", u)
            if qm:
                kind = (
                    "13c"
                    if "13 CONDUCTORS" in u
                    else "7c"
                    if "7 CONDUCTORS" in u
                    else "5c"
                    if "5 CONDUCTORS" in u
                    else "gnd"
                    if "GROUND" in u
                    else "1c"
                )
                parts.append((kind, int(qm.group(1))))
    return (cduit, cable, tuple(parts))


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)
    by_id = {x["id"]: x for x in live}
    live = sorted(by_id.values(), key=lambda x: -x["cy"])

    print(f"\n=== LIVE conduit/cable callouts ({len(live)}) ===")
    for L in live:
        print(
            f"id={L['id']} cy={L['cy']:.1f} fp={fp(L['lines'])} "
            f"head={L['lines'][0][:50]!r}"
        )

    page = fitz.open(PDF)[2]
    markups = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        if not content:
            continue
        cu = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" not in cu and "TRAFFIC SIGNAL CABLE" not in cu:
            continue
        # skip table-style paste of pay items
        if "SINGLE CONDUCTOR CABLE 10 GAGE" in cu and "LF" in cu:
            continue
        r = a.rect
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        dx, dy = pdf_to_dgn(cx, cy)
        lines = [ln.rstrip() for ln in content.split("\n") if ln.strip()]
        markups.append(
            {
                "i": i,
                "lines": lines,
                "fp": fp(lines),
                "dx": dx,
                "dy": dy,
                "stroke": (a.colors or {}).get("stroke"),
            }
        )

    print(f"\n=== 09.18 FreeTexts with CONDUIT/CABLE ({len(markups)}) ===")
    used = set()
    unmatched_markup = []
    for M in markups:
        # nearest unused live by Y then X
        cands = sorted(
            live,
            key=lambda L: abs(L["cy"] - M["dy"]) + 0.25 * abs(L["cx"] - M["dx"]),
        )
        best = None
        for L in cands:
            if L["id"] in used:
                continue
            best = L
            break
        if best is None:
            unmatched_markup.append(M)
            print(f"[{M['i']}] fp={M['fp']} NO LIVE LEFT")
            continue
        used.add(best["id"])
        same = M["fp"] == fp(best["lines"])
        # also check line-by-line ignore item mid-digit variants
        print(
            f"[{M['i']}] markup_fp={M['fp']} -> id={best['id']} "
            f"live_fp={fp(best['lines'])} match={same} "
            f"dy={abs(best['cy']-M['dy']):.1f}"
        )
        if not same:
            print("  markup lines:")
            for ln in M["lines"][:8]:
                print(f"    {ln}")
            print("  live lines:")
            for ln in best["lines"][:8]:
                print(f"    {ln}")

    unused_live = [L for L in live if L["id"] not in used]
    print(f"\n=== LIVE with no matched markup FreeText ({len(unused_live)}) ===")
    for L in unused_live:
        print(f"id={L['id']} cy={L['cy']:.1f} fp={fp(L['lines'])}")
        for ln in L["lines"][:6]:
            print(f"  {ln}")

    print(f"\n=== MARKUP FreeTexts with no live twin ({len(unmatched_markup)}) ===")
    for M in unmatched_markup:
        print(f"[{M['i']}] fp={M['fp']} dgn=({M['dx']:.0f},{M['dy']:.0f})")
        for ln in M["lines"][:6]:
            print(f"  {ln}")


if __name__ == "__main__":
    main()
