"""Content-matched compare: live conduit callouts vs r1 FreeTexts. Report only."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import fitz

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PDF_R1 = Path(r"C:\Users\RussinT\OneDrive - AECOM\Desktop") / "Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"

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
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def norm_item(line: str) -> str:
    out = line.replace("670.70011", "670.75070011")
    for rx, repl in ITEM_FIX:
        m = rx.search(out)
        if m:
            out = out[: m.start()] + repl + out[m.end() :]
    return out.rstrip()


def norm_lines(lines):
    return [norm_item(ln) for ln in lines if ln.strip()]


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


def cond_sig(lines):
    parts = []
    for ln in lines:
        u = ln.upper()
        if not any(k in u for k in ("CONDUCTORS", "GROUND WIRE", "SINGLE CONDUCTOR")):
            continue
        # tolerate broken "(1)(GROUND..." by also searching mid-line
        qm = re.search(r"\((\d+)\)(?:\s*\(SEE NOTE \d+\))?$", u)
        if not qm:
            qm = re.search(r"\((\d+)\)", u)
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
        note = "N" if "SEE NOTE" in u else ""
        parts.append(f"{kind}={qm.group(1)}{note}")
    return tuple(parts)


def fp(lines):
    return (conduit_qty(lines), cable_qty(lines), cond_sig(lines))


def score(mark_n, live_n):
    """Higher = better match. Prefer exact cable qty + conductor overlap."""
    mf, lf = fp(mark_n), fp(live_n)
    s = 0
    if mf[1] is not None and mf[1] == lf[1]:
        s += 100  # cable qty
    if mf[0] is not None and mf[0] == lf[0]:
        s += 20
    # conductor bag overlap
    ms, ls = set(mf[2]), set(lf[2])
    if ms and ls:
        s += 10 * len(ms & ls)
        s -= 5 * len(ms - ls)
        s -= 5 * len(ls - ms)
    if mark_n == live_n:
        s += 500
    return s


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
live = list(by.values())
for L in live:
    L["norm"] = norm_lines(L["lines"])
    L["fp"] = fp(L["norm"])

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
    lines = [ln.rstrip() for ln in content.split("\n") if ln.strip()]
    if i == 64:  # engineer: CABLE(6) means CABLE(11)
        lines = [
            ln.replace("TRAFFIC SIGNAL CABLE (6)", "TRAFFIC SIGNAL CABLE (11)")
            for ln in lines
        ]
    n = norm_lines(lines)
    markups.append({"i": i, "norm": n, "fp": fp(n)})

# greedy best-score matching
pairs = []
for M in markups:
    for L in live:
        pairs.append((score(M["norm"], L["norm"]), M["i"], L["id"]))
pairs.sort(reverse=True)
used_m, used_l = set(), set()
matched = []
for sc, mi, lid in pairs:
    if mi in used_m or lid in used_l:
        continue
    if sc < 50:  # weak — leave unmatched
        continue
    used_m.add(mi)
    used_l.add(lid)
    matched.append((mi, lid, sc))

print(f"\nLIVE={len(live)} markups={len(markups)} paired={len(matched)}")
print("\n=== EXACT / OK MATCHES ===")
mismatches = []
for mi, lid, sc in sorted(matched, key=lambda x: x[0]):
    M = next(m for m in markups if m["i"] == mi)
    L = by[lid]
    if M["norm"] == L["norm"]:
        print(f"  OK  markup[{mi}] <-> id={lid} CABLE={cable_qty(L['norm'])}")
    else:
        mismatches.append((mi, lid, sc, M, L))

print("\n=== CONTENT MISMATCHES (paired by fingerprint) ===")
for mi, lid, sc, M, L in mismatches:
    print(f"\n  markup[{mi}] <-> id={lid} score={sc}")
    print(f"    mark fp={M['fp']}")
    print(f"    live fp={L['fp']}")
    # line diffs
    ml, ll = M["norm"], L["norm"]
    maxlen = max(len(ml), len(ll))
    for k in range(maxlen):
        a = ml[k] if k < len(ml) else "<missing>"
        b = ll[k] if k < len(ll) else "<missing>"
        if a != b:
            print(f"    mark: {a}")
            print(f"    live: {b}")

print("\n=== UNPAIRED MARKUPS ===")
for M in markups:
    if M["i"] not in used_m:
        print(f"  markup[{M['i']}] fp={M['fp']}")
        for ln in M["norm"]:
            print(f"    {ln}")

print("\n=== UNPAIRED LIVE ===")
for L in sorted(live, key=lambda x: -x["cy"]):
    if L["id"] not in used_l:
        print(f"  id={L['id']} cy={L['cy']:.1f} fp={L['fp']}")
        for ln in L["norm"]:
            print(f"    {ln}")
