"""Audit 09.18 page-3 FreeTexts vs live fea_tsp callouts — gaps only."""
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
            t = "\n".join(lines)
            tu = t.upper()
            keys = (
                "TRAFFIC SIGNAL",
                "CONDUIT",
                "CABLE",
                "APS",
                "MAST ARM",
                "PULLBOX",
                "PEDESTRIAN",
                "LUMINAIRE",
                "TRANSFORMER",
                "POLE TYPE",
                "EXTENSION ARM",
                "LENS",
                "BRACKET",
            )
            if any(k in tu for k in keys):
                rng = el.Range
                out.append(
                    {
                        "id": el.ID.Low,
                        "lines": lines,
                        "t": t,
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


def classify(content: str) -> str:
    c = content.strip()
    cu = c.upper()
    if re.fullmatch(r"73\d{4,7}|75070011|70011", c.replace(" ", "")):
        return "item_digit"
    if re.fullmatch(r"\(\d+\)", c) or re.fullmatch(r"\d+", c):
        return "qty_chip"
    if re.match(r"730514|730714|7310514|7310714", c):
        return "item_qty"
    if any(k in cu for k in ("CONDUCTOR", "CABLE", "CONDUIT", "GROUND")):
        return "conduit_text"
    if any(
        k in cu
        for k in (
            "APS",
            "MAST",
            "POLE",
            "PULL",
            "PED",
            "LUMIN",
            "BRACKET",
            "LENS",
            "ARM",
            "TRANSFORMER",
            "SHAFT",
        )
    ):
        return "feature_callout"
    return "other"


def qty_on_line(line: str):
    m = re.search(r"\((\d+)\)\s*(\(SEE NOTE \d+\))?$", line)
    return int(m.group(1)) if m else None


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

    print("\n=== LIVE callouts (signal-ish) ===")
    for L in live:
        head = (L["lines"][0] if L["lines"] else "")[:72]
        print(f"id={L['id']} cy={L['cy']:.1f} cx={L['cx']:.0f} | {head}")

    # Cable header vs sum check
    print("\n=== CABLE header vs conductor+ground sum ===")
    for L in live:
        if not any(ln.upper().startswith("TRAFFIC SIGNAL CABLE") for ln in L["lines"]):
            continue
        header = next(
            ln for ln in L["lines"] if ln.upper().startswith("TRAFFIC SIGNAL CABLE")
        )
        hm = re.search(r"\((\d+)\)", header)
        hqty = int(hm.group(1)) if hm else None
        parts = []
        for ln in L["lines"]:
            u = ln.upper()
            if any(k in u for k in ("CONDUCTORS", "GROUND WIRE")):
                q = qty_on_line(ln)
                if q is not None:
                    parts.append(q)
        s = sum(parts)
        flag = "OK" if hqty == s else "MISMATCH"
        print(f"id={L['id']} header={hqty} sum={s} parts={parts} {flag}")

    page = fitz.open(PDF)[2]
    print("\n=== PAGE3 FreeText by class ===")
    buckets = {k: [] for k in (
        "item_digit",
        "qty_chip",
        "item_qty",
        "conduit_text",
        "feature_callout",
        "other",
    )}
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        if not content:
            continue
        stroke = (a.colors or {}).get("stroke")
        r = a.rect
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        dx, dy = pdf_to_dgn(cx, cy)
        kind = classify(content)
        near = sorted(
            live, key=lambda L: abs(L["cy"] - dy) + 0.25 * abs(L["cx"] - dx)
        )
        nid = near[0]["id"] if near else None
        ndy = abs(near[0]["cy"] - dy) if near else None
        buckets[kind].append((i, content, stroke, dx, dy, nid, ndy))

    for kind, items in buckets.items():
        print(f"\n-- {kind} ({len(items)}) --")
        for i, content, stroke, dx, dy, nid, ndy in items:
            print(
                f"[{i}] stroke={stroke} dgn=({dx:.0f},{dy:.0f}) "
                f"near={nid} dy={ndy:.1f} {content[:100]!r}"
            )

    # Unsure qty chips: ones that are not clearly reflected in nearest callout
    print("\n=== QTY CHIP vs nearest callout (manual review) ===")
    for i, content, stroke, dx, dy, nid, ndy in buckets["qty_chip"] + buckets["item_qty"]:
        L = by_id.get(nid)
        if not L:
            print(f"[{i}] {content!r} NO LIVE NEAR")
            continue
        print(f"[{i}] {content!r} -> id={nid} dy={ndy:.1f}")
        for ln in L["lines"]:
            if any(
                k in ln.upper()
                for k in ("CABLE (", "CONDUCTORS", "GROUND", "CONDUIT")
            ):
                print(f"     {ln}")

    print("\n=== FEATURE / CONDUIT text markups (possible missing callouts) ===")
    for kind in ("feature_callout", "conduit_text", "other"):
        for i, content, stroke, dx, dy, nid, ndy in buckets[kind]:
            # skip pure digit item fixes already done
            if kind == "other" and len(content) < 8:
                continue
            print(f"[{i}] {kind} near={nid} dy={ndy:.1f}")
            print(f"     {content[:200]!r}")


if __name__ == "__main__":
    main()
