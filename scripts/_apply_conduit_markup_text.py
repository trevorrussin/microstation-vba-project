"""Apply markup FreeText wording onto live fea_tsp conduit callouts (text only)."""
from __future__ import annotations

import sys
import time
from pathlib import Path

import fitz

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)


def norm_lines(lines):
    return [ln.rstrip() for ln in lines if ln.strip()]


def fingerprint(lines):
    cduit = cable = None
    parts = []
    for ln in lines:
        u = ln.upper()
        if u.startswith("TRAFFIC SIGNAL CONDUIT"):
            cduit = ln.strip()
        elif u.startswith("TRAFFIC SIGNAL CABLE"):
            cable = ln.strip()
        elif "CONDUCTORS" in u or "GROUND WIRE" in u or "SINGLE CONDUCTOR" in u:
            parts.append(u.split("(SEE")[0].strip())
    return (cduit, cable, tuple(parts))


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


def walk(el, out, depth=0):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = norm_lines(
                [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            )
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
                        "fp": fingerprint(lines),
                        "y": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "x": (float(rng.Low.X) + float(rng.High.X)) / 2,
                        "depth": depth,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out, depth + 1)
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("ACTIVE", app.ActiveDesignFile.FullName)

    page = fitz.open(PDF)[2]
    markups = []
    for i, a in enumerate(page.annots() or []):
        if a.type[1] != "FreeText":
            continue
        if (a.colors or {}).get("stroke") != [0.0, 1.0, 0.0]:
            continue
        content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
        u = content.upper()
        if "TRAFFIC SIGNAL CONDUIT" not in u and not u.startswith(
            "TRAFFIC SIGNAL CABLE"
        ):
            continue
        lines = norm_lines(content.splitlines())
        markups.append({"idx": i, "lines": lines, "fp": fingerprint(lines)})

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)

    # de-dupe by id (cell walk can re-hit)
    by_id = {}
    for L in live:
        by_id[L["id"]] = L
    live = list(by_id.values())
    print(f"markups={len(markups)} live={len(live)}")

    used_m, used_l = set(), set()
    pairs = []

    live_i = list(enumerate(live))
    for li, L in live_i:
        for mi, M in enumerate(markups):
            if mi in used_m:
                continue
            if L["lines"] == M["lines"]:
                pairs.append((li, mi, "exact"))
                used_m.add(mi)
                used_l.add(li)
                break

    for li, L in live_i:
        if li in used_l:
            continue
        best = None
        for mi, M in enumerate(markups):
            if mi in used_m:
                continue
            if L["fp"][0] and L["fp"][0] == M["fp"][0] and L["fp"][1] == M["fp"][1]:
                s = len(set(L["fp"][2]) & set(M["fp"][2]))
                if best is None or s > best[0]:
                    best = (s, mi)
            elif (
                L["fp"][0] is None
                and M["fp"][0] is None
                and L["fp"][1] == M["fp"][1]
            ):
                s = len(set(L["fp"][2]) & set(M["fp"][2]))
                if best is None or s > best[0]:
                    best = (s, mi)
        if best and best[0] >= 2:
            pairs.append((li, best[1], f"fp:{best[0]}"))
            used_m.add(best[1])
            used_l.add(li)

    print("\n=== MATCHES ===")
    updates = []
    for li, mi, how in pairs:
        L, M = live[li], markups[mi]
        same = L["lines"] == M["lines"]
        status = "OK" if same else "DIFF"
        print(f"{status} live={L['id']} <-> [{M['idx']}] ({how})")
        if not same:
            for a, b in zip(L["lines"], M["lines"]):
                if a != b:
                    print(f"  ~ {a!r}")
                    print(f"  > {b!r}")
            if len(L["lines"]) != len(M["lines"]):
                print(f"  linecount {len(L['lines'])} -> {len(M['lines'])}")
            updates.append((L, M))

    print("\n=== UNMATCHED LIVE ===")
    for li, L in live_i:
        if li not in used_l:
            print(f"id={L['id']} y={L['y']:.1f}")
            for ln in L["lines"]:
                print(" ", repr(ln))

    print("\n=== UNMATCHED MARKUP (no live twin) ===")
    for mi, M in enumerate(markups):
        if mi not in used_m:
            print(f"[{M['idx']}]")
            for ln in M["lines"]:
                print(" ", repr(ln))

    print("\n=== APPLY (SetTextLine only) ===")
    for L, M in updates:
        set_lines(L["el"], M["lines"])
        print(f"UPDATED id={L['id']} from markup [{M['idx']}]")
        for ln in M["lines"]:
            print(" ", ln)

    if updates:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.5)
        print("SAVED", len(updates), "callouts")
    else:
        print("Nothing to change on matched pairs.")


if __name__ == "__main__":
    main()
