"""Summarize new vs old offset CSVs; probe AC length vs stationing."""
from __future__ import annotations

import csv
import math
import re
import sys
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "mcp-server"))

import pythoncom
import ms_connect
from win32com.client import Dispatch

OFFSET = "Offset From Alignment Best Fit Line to Photogrammetry (ft)"
NEW = ROOT / "outputs/727-alignment-offset/dashed_sandwich_offset_20260915_183936.csv"
OLD = ROOT / "outputs/727-alignment-offset/dashed_sandwich_offset_20260903_144339.csv"


def parse_sta(s: str) -> float | None:
    m = re.match(r"(\d+)\+(\d+(?:\.\d+)?)", s.strip().replace(" ", ""))
    if not m:
        return None
    return int(m.group(1)) * 100.0 + float(m.group(2))


def load(p: Path):
    return list(csv.DictReader(p.open(encoding="utf-8-sig")))


def stats(rows, label):
    stas = [parse_sta(r["station"]) for r in rows]
    stas = [s for s in stas if s is not None]
    both = sum(1 for r in rows if r.get("north") and r.get("south"))
    one = sum(1 for r in rows if bool(r.get("north")) ^ bool(r.get("south")))
    none = sum(1 for r in rows if not r.get("north") and not r.get("south"))
    offs = []
    for r in rows:
        try:
            if r.get(OFFSET):
                offs.append(float(r[OFFSET]))
        except ValueError:
            pass
    print(f"=== {label} ===")
    print(f" rows={len(rows)} both={both} one={one} blank={none}")
    print(f" sta {min(stas):.2f} .. {max(stas):.2f}")
    if offs:
        print(
            f" offset n={len(offs)} mean={sum(offs)/len(offs):.3f} "
            f"max={max(offs):.3f} >1={sum(1 for o in offs if o>1)} >8={sum(1 for o in offs if o>8)}"
        )
    east = [r for r in rows if (parse_sta(r["station"]) or 0) >= 24850]
    print(
        f" sta>=248+50 rows={len(east)} both={sum(1 for r in east if r.get('north') and r.get('south'))} "
        f"blank={sum(1 for r in east if not r.get('north') and not r.get('south'))}"
    )
    print(" last 12:")
    for r in rows[-12:]:
        print(
            f"  {r['station']:10} pt={r.get('point',''):4} "
            f"N={r.get('north',''):6} S={r.get('south',''):6} "
            f"off={r.get(OFFSET,''):6} bf={r.get('bf_north_south','')}"
        )


stats(load(OLD), "OLD")
stats(load(NEW), "NEW")

checks = ["19+50", "45+50", "75+50", "241+50", "248+50", "249+00", "270+50", "290+50", "299+75"]
by = {r["station"]: r for r in load(NEW)}
print("=== checks NEW ===")
for c in checks:
    r = by.get(c)
    if not r:
        print(c, "MISSING")
    else:
        print(
            c,
            "N",
            r.get("north"),
            "S",
            r.get("south"),
            "off",
            r.get(OFFSET),
            "bf",
            r.get("bf_north_south"),
            "pt",
            r.get("point"),
        )

# Probe AC endpoints / length without full script chaining
pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
adf = app.ActiveDesignFile
mr = adf.DefaultModelReference
lv = adf.Levels("AC")
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
o.ExcludeAllLevels()
o.IncludeLevel(lv)
ee = mr.Scan(o)
xs = []
ys = []
n_line = n_arc = n_cs = 0
while ee.MoveNext():
    e = ee.Current
    t = int(e.Type)
    if t == 12:
        n_cs += 1
    elif t == 16:
        n_arc += 1
    elif t in (3, 4):
        n_line += 1
    try:
        r = e.Range
        xs.extend([float(r.Low.X), float(r.High.X)])
        ys.extend([float(r.Low.Y), float(r.High.Y)])
    except Exception:
        pass
print("AC elem types lines", n_line, "arcs", n_arc, "complex", n_cs)
print("AC bbox X", min(xs), max(xs), "dx", max(xs) - min(xs))
print("AC bbox Y", min(ys), max(ys), "dy", max(ys) - min(ys))

# Station label X span
lv = adf.Levels("AC_C")
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
o.ExcludeAllLevels()
o.IncludeLevel(lv)
ee = mr.Scan(o)
lab = []
while ee.MoveNext():
    e = ee.Current
    try:
        if e.IsTextElement:
            t = e.AsTextElement.Text
            ox = float(e.AsTextElement.Origin.X)
            oy = float(e.AsTextElement.Origin.Y)
        elif e.IsTextNodeElement:
            t = e.AsTextNodeElement.TextLine(1)
            ox = float(e.AsTextNodeElement.Origin.X)
            oy = float(e.AsTextNodeElement.Origin.Y)
        else:
            continue
    except Exception:
        continue
    if re.match(r"^\s*(PC|PT)\b", t, re.I):
        continue
    sta = parse_sta(t)
    if sta is None or sta < 1950:
        continue
    lab.append((sta, ox, oy))
lab.sort()
print("labels", len(lab), "X", lab[0][1], "..", lab[-1][1], "dx", lab[-1][1] - lab[0][1])
print("first label", lab[0], "last", lab[-1])
# labels near havemeyer X~1026444
near = sorted(lab, key=lambda r: abs(r[1] - 1026444))[:5]
print("labels nearest Havemeyer X", near)
