"""Loc01 curb verify gate — run before claiming classify DONE.

Exit 0 = pass (or WARN-only). Exit 1 = FAIL hard checks.
"""
from __future__ import annotations

import math
import sys
import pythoncom
from win32com.client import GetObject

STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]
SHORT = {
    STEEL[0]: "Straight",
    STEEL[1]: "Depressed",
    STEEL[2]: "Corner",
}
OUTERS = {
    "13D80A": "NW",
    "13D25C": "N",
    "13D7DC": "SW",
    "13CEC7": "E",
    "13D2EA": "MidN",
    "13D7BC": "MidS",
}
ISLANDS = [
    {"name": "NW", "center": (1028470.0, 187478.0), "r": 28.0, "expect_dep": True},
    {"name": "N", "center": (1028564.0, 187503.0), "r": 28.0, "expect_dep": True},
    {"name": "SW", "center": (1028486.0, 187421.0), "r": 28.0, "expect_dep": True},
    {"name": "E", "center": (1028589.0, 187445.0), "r": 28.0, "expect_dep": False},
    {"name": "MidN", "center": (1028501.0, 187485.0), "r": 22.0, "expect_dep": False},
    {"name": "MidS", "center": (1028525.0, 187428.0), "r": 22.0, "expect_dep": False},
]
# Depressed LF ballpark when PROP flares exist (from NW-process runs)
DEP_MIN = {"NW": 15.0, "N": 10.0, "SW": 20.0}
STRAIGHT_MIN = {"NW": 8.0}  # NW must keep long red runs


def main() -> int:
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("VERIFY Loc01 curb")
    print("file", doc.FullName)
    fails: list[str] = []
    warns: list[str] = []

    for h, name in OUTERS.items():
        try:
            e = doc.HandleToObject(h)
            print(f"  OUTER OK {name} {h} {e.ObjectName}")
        except Exception:
            fails.append(f"OUTER missing {name} {h}")
            print(f"  OUTER FAIL {name} {h}")

    ms = doc.ModelSpace
    segs: list[tuple[str, float, float, float]] = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer not in SHORT or e.ObjectName != "AcDbLine":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187390 < cy < 187530):
                continue
            segs.append((e.Layer, float(e.Length), cx, cy))
        except Exception:
            continue

    print(f"  steel segs in Loc01 box: {len(segs)}")
    for isl in ISLANDS:
        by = {STEEL[0]: 0.0, STEEL[1]: 0.0, STEEL[2]: 0.0}
        n = {STEEL[0]: 0, STEEL[1]: 0, STEEL[2]: 0}
        for lyr, L, cx, cy in segs:
            if math.hypot(cx - isl["center"][0], cy - isl["center"][1]) <= isl["r"]:
                by[lyr] += L
                n[lyr] += 1
        print(
            f"  {isl['name']:5s}  "
            f"S={by[STEEL[0]]:7.2f}({n[STEEL[0]]:3d})  "
            f"D={by[STEEL[1]]:7.2f}({n[STEEL[1]]:3d})  "
            f"C={by[STEEL[2]]:7.2f}({n[STEEL[2]]:3d})"
        )
        if isl["expect_dep"] and by[STEEL[1]] < DEP_MIN.get(isl["name"], 8.0):
            fails.append(
                f"{isl['name']}: Depressed LF {by[STEEL[1]]:.2f} < min "
                f"{DEP_MIN.get(isl['name'], 8.0)}"
            )
        if isl["name"] in STRAIGHT_MIN and by[STEEL[0]] < STRAIGHT_MIN[isl["name"]]:
            fails.append(
                f"{isl['name']}: Straight LF {by[STEEL[0]]:.2f} < min "
                f"{STRAIGHT_MIN[isl['name']]} (long red runs missing?)"
            )
        if by[STEEL[0]] + by[STEEL[1]] + by[STEEL[2]] < 5.0:
            warns.append(f"{isl['name']}: almost no steel in keep radius")
        # Heuristic: if Corner >> Straight+Depressed on PROP island, likely densify bug
        if isl["expect_dep"] and by[STEEL[2]] > 3.0 * max(by[STEEL[0]], 1.0) and by[STEEL[0]] < 5.0:
            warns.append(
                f"{isl['name']}: Corner dominates with little Straight — check LONG_SEG/RDP"
            )

    for w in warns:
        print("  WARN", w)
    if fails:
        print("FAIL")
        for f in fails:
            print(" ", f)
        return 1
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
