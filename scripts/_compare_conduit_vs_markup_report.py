"""Compare live fea_tsp conduit callouts to r1 FreeTexts + 09.18 digit/qty chips.
Report-only; no edits.
"""
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

PDF_R1 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
PDF18 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"

# 09.18 authority for item mid-digits (engineer-corrected)
ITEM_FIX = [
    (re.compile(r"\(5 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
     "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.7310514)"),
    (re.compile(r"\(7 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
     "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.7310714)"),
    (re.compile(r"\(13 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
     "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314)"),
    (re.compile(r"\(GROUND WIRE 10 AWG\) \(ITEM NO\. 670\.\d+\)"),
     "(GROUND WIRE 10 AWG) (ITEM NO. 670.75070011)"),
]

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
                        "lines": [ln.rstrip() for ln in lines if ln.strip()],
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


def norm_item(line: str) -> str:
    out = line
    for rx, repl in ITEM_FIX:
        m = rx.search(out)
        if m:
            out = out[: m.start()] + repl + out[m.end() :]
    # also fix bare 670.70011 style if present without full pattern
    out = out.replace("670.70011", "670.75070011")
    return out


def norm_lines(lines):
    return [norm_item(ln.rstrip()) for ln in lines if ln.strip()]


def cable_qty(lines):
    for ln in lines:
        m = re.search(r"TRAFFIC SIGNAL CABLE\s*\((\d+)\)", ln.upper())
        if m:
            return int(m.group(1))
    return None


def conduit_qty(lines):
    for ln in lines:
        m = re.search(r"TRAFFIC SIGNAL CONDUIT\s*\((\d+)\)", ln.upper())
        if m:
            return int(m.group(1))
    return None


def conductor_sig(lines):
    parts = []
    for ln in lines:
        u = ln.upper()
        if not any(
            k in u for k in ("CONDUCTORS", "GROUND WIRE", "SINGLE CONDUCTOR")
        ):
            continue
        qm = re.search(r"\((\d+)\)\s*(\(SEE NOTE \d+\))?$", u)
        if not qm:
            continue
        if "13 CONDUCTORS" in u:
            kind = "13c"
        elif "7 CONDUCTORS" in u:
            kind = "7c"
        elif "5 CONDUCTORS" in u:
            kind = "5c"
        elif "SINGLE CONDUCTOR" in u:
            kind = "1c"
        elif "GROUND" in u:
            kind = "gnd"
        else:
            kind = "?"
        note = " NOTE" if qm.group(2) else ""
        parts.append(f"{kind}={qm.group(1)}{note}")
    return parts


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    print("DGN", app.ActiveDesignFile.FullName)
    mr = app.ActiveModelReference
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)
    by = {x["id"]: x for x in live}
    live = sorted(by.values(), key=lambda x: -x["cy"])

    # r1 FreeTexts
    page = fitz.open(PDF_R1)[2]
    markups = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        cu = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" not in cu and "TRAFFIC SIGNAL CABLE" not in cu:
            continue
        if "LF" in cu and "SINGLE CONDUCTOR CABLE 10 GAGE" in cu:
            continue
        r = a.rect
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        # crude pdf->approx using 09.18 anchors (same sheet family)
        dx, dy = pdf_to_dgn(cx, cy)
        lines = [ln.rstrip() for ln in content.split("\n") if ln.strip()]
        # engineer: FreeText [64] CABLE(6) means CABLE(11)
        if i == 64:
            lines = [
                ln.replace("TRAFFIC SIGNAL CABLE (6)", "TRAFFIC SIGNAL CABLE (11)")
                for ln in lines
            ]
        markups.append(
            {
                "i": i,
                "lines": lines,
                "dx": dx,
                "dy": dy,
                "norm": norm_lines(lines),
            }
        )

    print(f"\nLIVE={len(live)}  r1_FreeTexts={len(markups)}")
    used = set()
    mismatches = []
    matches = []

    for M in markups:
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
            mismatches.append(("NO_LIVE", M, None))
            continue
        used.add(best["id"])
        live_n = norm_lines(best["lines"])
        mark_n = M["norm"]
        # compare ignoring trailing spaces; allow SEE NOTE line placement diffs separately
        if live_n == mark_n:
            matches.append((M["i"], best["id"]))
            continue
        # detailed diff
        diffs = []
        # header cable/conduit
        if cable_qty(live_n) != cable_qty(mark_n):
            diffs.append(
                f"CABLE live={cable_qty(live_n)} markup={cable_qty(mark_n)}"
            )
        if conduit_qty(live_n) != conduit_qty(mark_n):
            diffs.append(
                f"CONDUIT live={conduit_qty(live_n)} markup={conduit_qty(mark_n)}"
            )
        ls = conductor_sig(live_n)
        ms = conductor_sig(mark_n)
        if ls != ms:
            diffs.append(f"conductors live={ls} markup={ms}")
        # line-by-line if still unclear
        if not diffs:
            for a, b in zip(live_n, mark_n):
                if a != b:
                    diffs.append(f"line\n    live: {a}\n    mark: {b}")
            if len(live_n) != len(mark_n):
                diffs.append(f"n_lines live={len(live_n)} markup={len(mark_n)}")
        mismatches.append(("DIFF", M, best, diffs, live_n, mark_n))

    unused = [L for L in live if L["id"] not in used]

    print("\n=== MATCH (item-digits normalized) ===")
    for mi, lid in matches:
        print(f"  markup[{mi}] <-> id={lid}")

    print("\n=== MISMATCH vs r1 FreeText (expected diffs) ===")
    for row in mismatches:
        kind = row[0]
        if kind == "NO_LIVE":
            M = row[1]
            print(f"\n  markup[{M['i']}] NO LIVE TWIN")
            for ln in M["norm"]:
                print(f"    mark: {ln}")
            continue
        _, M, best, diffs, live_n, mark_n = row
        print(
            f"\n  markup[{M['i']}] <-> id={best['id']} cy={best['cy']:.1f} "
            f"CABLE live={cable_qty(live_n)} mark={cable_qty(mark_n)}"
        )
        for d in diffs:
            print(f"    ! {d}")
        print("    --- live ---")
        for ln in live_n:
            print(f"      {ln}")
        print("    --- markup (normalized) ---")
        for ln in mark_n:
            print(f"      {ln}")

    print(f"\n=== LIVE with no r1 FreeText twin ({len(unused)}) ===")
    for L in unused:
        print(
            f"  id={L['id']} cy={L['cy']:.1f} CABLE={cable_qty(L['lines'])} "
            f"sig={conductor_sig(L['lines'])}"
        )
        for ln in L["lines"]:
            print(f"      {ln}")

    # 09.18 qty chips that look like conductor qty edits
    print("\n=== 09.18 qty/item chips (context only) ===")
    page18 = fitz.open(PDF18)[2]
    for i, a in enumerate(page18.annots() or []):
        if a.type[1] != "FreeText":
            continue
        c = ((a.info or {}).get("content") or "").strip()
        if not c:
            continue
        if not (
            re.fullmatch(r"\(\d+\)", c)
            or re.fullmatch(r"\d+", c)
            or re.match(r"730514|730714|7310514|7310714", c)
            or "13 CONDUCTORS" in c.upper()
        ):
            continue
        r = a.rect
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        dx, dy = pdf_to_dgn(cx, cy)
        near = sorted(
            live, key=lambda L: abs(L["cy"] - dy) + 0.25 * abs(L["cx"] - dx)
        )[0]
        print(
            f"  [{i}] {c!r} -> nearest id={near['id']} "
            f"CABLE={cable_qty(near['lines'])} dy={abs(near['cy']-dy):.1f}"
        )


if __name__ == "__main__":
    main()
