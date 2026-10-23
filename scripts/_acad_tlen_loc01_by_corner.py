"""Loc01 curb LF by corner — print to AutoCAD command line, then wait for typed input."""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import pythoncom
from win32com.client import GetObject
from win32com.client.dynamic import Dispatch

ROOT = Path(__file__).resolve().parents[1]
MAP_PATH = ROOT / "Data" / "loc01-island-to-corner.json"

STEEL = {
    "Straight Steel Faced Concrete Curb (18'' Deep)": "Straight",
    "Depressed Steel Faced Concrete Curb (18'' Deep)": "Depressed",
    "Corner Steel Faced Concrete Curb (18'' Deep)": "Corner",
}

ISLANDS = [
    {"name": "NW", "center": (1028470.0, 187478.0), "r": 28.0},
    {"name": "N", "center": (1028564.0, 187503.0), "r": 28.0},
    {"name": "SW", "center": (1028486.0, 187421.0), "r": 28.0},
    {"name": "E", "center": (1028589.0, 187445.0), "r": 28.0},
    {"name": "MidN", "center": (1028501.0, 187485.0), "r": 22.0},
    {"name": "MidS", "center": (1028525.0, 187428.0), "r": 22.0},
]
ORDER = ["NW", "N", "SW", "E", "MidN", "MidS"]


def bring_forward(app) -> None:
    try:
        import win32con
        import win32gui

        hwnd = int(app.HWND)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def line_len(e) -> float:
    try:
        d = Dispatch(e)
        try:
            return float(d.Length)
        except Exception:
            pass
        s, t = d.StartPoint, d.EndPoint
        return math.hypot(float(t[0]) - float(s[0]), float(t[1]) - float(s[1]))
    except Exception:
        return 0.0


def nearest_island(cx: float, cy: float):
    best = None
    best_d = 1e18
    for isl in ISLANDS:
        d = math.hypot(cx - isl["center"][0], cy - isl["center"][1])
        if d <= isl["r"] and d < best_d:
            best = isl
            best_d = d
    return best


def main() -> int:
    pythoncom.CoInitialize()
    meta = json.loads(MAP_PATH.read_text(encoding="utf-8"))
    island_meta = meta["islands"]

    app = GetObject(None, "AutoCAD.Application")
    doc = app.ActiveDocument
    print("file", doc.FullName)
    app.Visible = True
    bring_forward(app)

    totals: dict[str, dict[str, float]] = {
        i["name"]: {"Straight": 0.0, "Depressed": 0.0, "Corner": 0.0} for i in ISLANDS
    }
    counts: dict[str, dict[str, int]] = {
        i["name"]: {"Straight": 0, "Depressed": 0, "Corner": 0} for i in ISLANDS
    }

    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            kind = STEEL.get(e.Layer)
            if not kind or e.ObjectName != "AcDbLine":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2.0
            cy = (float(mn[1]) + float(mx[1])) / 2.0
            isl = nearest_island(cx, cy)
            if isl is None:
                continue
            L = line_len(e)
            if L <= 0.0:
                continue
            name = isl["name"]
            totals[name][kind] += L
            counts[name][kind] += 1
        except Exception:
            continue

    # Build report — NW first, then all corners
    report_lines: list[str] = [
        "==== Loc01 curb LF by corner ====",
        "",
    ]
    for name in ORDER:
        m = island_meta.get(name, {})
        cid = m.get("corner_id", "?")
        loc = m.get("location_no", "?")
        t = totals[name]
        c = counts[name]
        block = (
            f"{name}  Loc{loc}  CornerID {cid}  |  "
            f"Straight {t['Straight']:.2f} ({c['Straight']})  "
            f"Depressed {t['Depressed']:.2f} ({c['Depressed']})  "
            f"Corner {t['Corner']:.2f} ({c['Corner']})  "
            f"TOTAL {t['Straight']+t['Depressed']+t['Corner']:.2f}"
        )
        report_lines.append(block)
        print(block)

    report_lines.append("")
    report_lines.append(
        "Type corner name (NW/N/SW/E/MidN/MidS) for detail, or Enter to finish:"
    )

    # Dump to AutoCAD command line (no modal alert — user can type)
    util = doc.Utility
    for line in report_lines:
        try:
            util.Prompt(line + "\n")
        except Exception:
            pass

    # Interactive typed prompt on AutoCAD command line
    try:
        typed = util.GetString(
            0,
            "Corner [NW/N/SW/E/MidN/MidS] or Enter: ",
        )
    except Exception as ex:
        print("GetString failed (cancel or busy):", ex)
        typed = ""

    typed = (typed or "").strip()
    print("typed:", repr(typed))

    if typed:
        key = typed.upper().replace(" ", "")
        # allow MidN / MIDN
        alias = {
            "MIDN": "MidN",
            "MIDS": "MidS",
            "NW": "NW",
            "N": "N",
            "SW": "SW",
            "E": "E",
        }
        name = alias.get(key, typed if typed in totals else None)
        if name and name in totals:
            m = island_meta.get(name, {})
            t = totals[name]
            c = counts[name]
            detail = (
                f"\n---- {name} detail ----\n"
                f"Loc {m.get('location_no')}  CornerID {m.get('corner_id')}\n"
                f"{m.get('note', '')}\n"
                f"Straight   {t['Straight']:8.2f} ft  ({c['Straight']} segs)\n"
                f"Depressed  {t['Depressed']:8.2f} ft  ({c['Depressed']} segs)\n"
                f"Corner     {t['Corner']:8.2f} ft  ({c['Corner']} segs)\n"
                f"TOTAL      {t['Straight']+t['Depressed']+t['Corner']:8.2f} ft\n"
            )
            print(detail)
            util.Prompt(detail)
            # Second type-in so they always have a prompt after detail too
            try:
                util.GetString(0, "Press Enter (or type another corner): ")
            except Exception:
                pass
        else:
            util.Prompt(f"\nUnknown corner '{typed}'. Done.\n")

    util.Prompt("\nTLEN Loc01 by-corner done.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
