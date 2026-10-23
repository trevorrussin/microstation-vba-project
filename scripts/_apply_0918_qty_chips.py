"""Map 09.18 qty chips to live callouts and apply text-only qty fixes."""
from __future__ import annotations

import re
import sys
import time
from pathlib import Path

import fitz
import numpy as np

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Bally's_Hutch_09.18.26_Traffic_BTS.pdf"

# From rebuild script (r1). Re-fit if 09.18 page size differs.
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


def pdf_to_dgn(px: float, py: float) -> tuple[float, float]:
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
                        "el": el,
                        "lines": lines,
                        "x0": float(rng.Low.X),
                        "x1": float(rng.High.X),
                        "y0": float(rng.Low.Y),
                        "y1": float(rng.High.Y),
                        "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                        "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def set_lines(tn_el, lines):
    tn = tn_el.AsTextNodeElement
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    tn_el.Rewrite()


def nearest_callout(live, dx, dy):
    best = None
    for L in live:
        # chip usually just east of RJ text
        d = abs(L["cy"] - dy) + 0.15 * abs(L["cx"] - dx)
        if best is None or d < best[0]:
            best = (d, L)
    return best


def apply_qty_on_line(line: str, new_qty: str) -> str | None:
    """Replace trailing (N) qty before optional SEE NOTE."""
    # Match last (digits) that is a qty, not part of ITEM NO. 680.xxx
    # Patterns: ... 680.730514) (1)   or  ... (1) (SEE NOTE 5)
    m = re.search(r"(\) )?(\(\d+\))(\s*\(SEE NOTE \d+\))?$", line)
    if not m:
        # cable header TRAFFIC SIGNAL CABLE (16)
        m2 = re.match(r"^(TRAFFIC SIGNAL CABLE )\((\d+)\)$", line.strip())
        if m2:
            return f"{m2.group(1)}{new_qty}"
        return None
    # If the (N) sits right after ITEM NO closing paren
    prefix = line[: m.start(2)]
    suffix = line[m.end(2) :]
    return prefix + new_qty + suffix


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    doc = fitz.open(PDF)
    page = doc[2]
    print("PDF", PDF.name, "page3", page.rect)

    # Collect green qty-like chips
    chips = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        if not content:
            continue
        r = a.rect
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        # classify
        kind = None
        payload = content
        if re.fullmatch(r"\(\d+\)", content):
            kind = "qty"
        elif re.fullmatch(r"\d+", content) and len(content) <= 3:
            kind = "qty_bare"  # e.g. 2
            payload = f"({content})"
        elif re.fullmatch(r"730514 \(\d+\)", content):
            kind = "item5_qty"
        elif re.fullmatch(r"730714 \(\d+\)", content):
            kind = "item7_qty"
        elif "13 CONDUCTORS" in content.upper():
            kind = "line13"
        elif re.fullmatch(r"7310514|7310714|75070011|730514|730714", content):
            kind = "item_digits_only"  # already handled
        else:
            continue
        if kind == "item_digits_only":
            continue
        dx, dy = pdf_to_dgn(cx, cy)
        chips.append(
            {
                "idx": i,
                "kind": kind,
                "payload": payload,
                "content": content,
                "pdf": (cx, cy),
                "dgn": (dx, dy),
            }
        )

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)
    by_id = {L["id"]: L for L in live}
    live = list(by_id.values())
    print(f"live={len(live)} chips={len(chips)}")

    # Propose edits per callout
    planned = {L["id"]: list(L["lines"]) for L in live}
    reports = []

    for ch in chips:
        dist, L = nearest_callout(live, *ch["dgn"])
        lines = planned[L["id"]]
        dx, dy = ch["dgn"]
        # pick target line by Y within callout
        # estimate line ys evenly in range
        n = len(lines)
        if n == 0:
            continue
        line_ys = [
            L["y1"] - (i + 0.5) * (L["y1"] - L["y0"]) / n for i in range(n)
        ]
        li = min(range(n), key=lambda i: abs(line_ys[i] - dy))
        target = lines[li]
        new_line = None
        note = ""

        if ch["kind"] in ("qty", "qty_bare"):
            # Prefer conductor/ground/cable-header lines over conduit header
            candidates = []
            for i, ln in enumerate(lines):
                u = ln.upper()
                if "CONDUCTORS" in u or "GROUND WIRE" in u or u.startswith(
                    "TRAFFIC SIGNAL CABLE"
                ):
                    candidates.append((abs(line_ys[i] - dy), i, ln))
            if candidates:
                candidates.sort()
                li = candidates[0][1]
                target = lines[li]
            new_line = apply_qty_on_line(target, ch["payload"])
            note = f"qty->{ch['payload']} on line {li}"
        elif ch["kind"] == "item5_qty":
            # find 5-conductor line
            for i, ln in enumerate(lines):
                if "5 CONDUCTORS" in ln.upper():
                    # set item already fixed; set qty
                    q = re.search(r"\((\d+)\)$", ch["content"].replace(" ", ""))
                    # content like '730514 (1)'
                    qm = re.search(r"\((\d+)\)", ch["content"])
                    qty = f"({qm.group(1)})" if qm else "(1)"
                    base = re.sub(
                        r"\(ITEM NO\. 680\.\d+\)",
                        "(ITEM NO. 680.730514)",
                        ln,
                    )
                    new_line = apply_qty_on_line(base, qty) or base
                    li = i
                    note = f"5-cond item+qty {ch['content']}"
                    break
        elif ch["kind"] == "item7_qty":
            for i, ln in enumerate(lines):
                if "7 CONDUCTORS" in ln.upper():
                    qm = re.search(r"\((\d+)\)", ch["content"])
                    qty = f"({qm.group(1)})" if qm else "(1)"
                    base = re.sub(
                        r"\(ITEM NO\. 680\.\d+\)",
                        "(ITEM NO. 680.730714)",
                        ln,
                    )
                    new_line = apply_qty_on_line(base, qty) or base
                    li = i
                    note = f"7-cond item+qty {ch['content']}"
                    break
        elif ch["kind"] == "line13":
            # full replacement line from markup (fix obvious typos)
            fixed = ch["content"]
            fixed = fixed.replace("680.71314)", "680.731314)")
            fixed = fixed.replace("ITEM NO.", "(ITEM NO.")
            # normalize spacing
            if not fixed.startswith("("):
                fixed = "(" + fixed if "CONDUCTORS" in fixed else fixed
            # markup: '(13 CONDUCTORS, 14 AWG) ITEM NO. 680.71314)  (5)'
            fixed = re.sub(
                r"\(13 CONDUCTORS, 14 AWG\)\s*\(?ITEM NO\.\s*680\.\d+\)\s*\((\d+)\)",
                r"(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (\1)",
                fixed,
            )
            for i, ln in enumerate(lines):
                if "13 CONDUCTORS" in ln.upper():
                    new_line = fixed if "680.731314" in fixed else (
                        f"(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) (5)"
                    )
                    # extract qty from chip
                    qm = re.search(r"\((\d+)\)\s*$", ch["content"])
                    if qm:
                        new_line = f"(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314) ({qm.group(1)})"
                    li = i
                    note = "13-cond line from markup"
                    break

        reports.append(
            {
                "chip": ch,
                "callout": L["id"],
                "dist": dist,
                "li": li,
                "before": target,
                "after": new_line,
                "note": note,
            }
        )
        if new_line and new_line != lines[li]:
            lines[li] = new_line
            planned[L["id"]] = lines

    print("\n=== PLAN ===")
    for r in reports:
        ch = r["chip"]
        print(
            f"chip[{ch['idx']}] {ch['kind']} {ch['content']!r} -> "
            f"id={r['callout']} dist={r['dist']:.1f} | {r['note']}"
        )
        if r["after"] and r["after"] != r["before"]:
            print(f"  - {r['before']!r}")
            print(f"  + {r['after']!r}")
        elif r["after"] is None:
            print(f"  SKIP no apply on {r['before']!r}")
        else:
            print("  (no change)")

    # Apply
    print("\n=== APPLY ===")
    touched = 0
    line_edits = 0
    for L in live:
        before = L["lines"]
        after = planned[L["id"]]
        if before == after:
            continue
        diffs = sum(1 for a, b in zip(before, after) if a != b) + abs(
            len(before) - len(after)
        )
        line_edits += diffs
        touched += 1
        print(f"id={L['id']} diffs={diffs}")
        for a, b in zip(before, after):
            if a != b:
                print(f"  - {a!r}")
                print(f"  + {b!r}")
        set_lines(L["el"], after)

    if touched:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
    print(f"\nSAVED callouts={touched} line_edits={line_edits}")


if __name__ == "__main__":
    main()
