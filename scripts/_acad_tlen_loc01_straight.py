"""TLEN-equivalent: total length of Loc01 Straight steel curb segments."""
from __future__ import annotations

import sys
import pythoncom
from win32com.client import GetObject

LAYER = "Straight Steel Faced Concrete Curb (18'' Deep)"
# Loc01 keep box (same as verify)
X0, X1 = 1028440.0, 1028620.0
Y0, Y1 = 187390.0, 187530.0


def seg_len(e) -> float:
    name = e.ObjectName
    try:
        if name == "AcDbLine":
            return float(e.Length)
        if name == "AcDbArc":
            return float(e.ArcLength)
        if name in (
            "AcDbCircle",
            "AcDbSpline",
            "AcDbPolyline",
            "AcDb2dPolyline",
            "AcDb3dPolyline",
            "AcDbEllipse",
        ):
            try:
                return float(e.Length)
            except Exception:
                return 0.0
        if "FeatureLine" in name:
            try:
                return float(e.Length)
            except Exception:
                return 0.0
    except Exception:
        return 0.0
    return 0.0


def in_loc01(e) -> bool:
    try:
        mn, mx = e.GetBoundingBox()
        cx = (float(mn[0]) + float(mx[0])) / 2.0
        cy = (float(mn[1]) + float(mx[1])) / 2.0
        return X0 < cx < X1 and Y0 < cy < Y1
    except Exception:
        return False


def main() -> int:
    pythoncom.CoInitialize()
    app = GetObject(None, "AutoCAD.Application")
    doc = app.ActiveDocument
    print("TLEN Loc01 Straight")
    print("file", doc.FullName)

    tl = 0.0
    n = 0
    by_type: dict[str, int] = {}
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != LAYER:
                continue
            if not in_loc01(e):
                continue
            L = seg_len(e)
            tl += L
            n += 1
            by_type[e.ObjectName] = by_type.get(e.ObjectName, 0) + 1
        except Exception:
            continue

    msg = f"Total length of selected objects is {tl:.4f}"
    print("layer", LAYER)
    print("count", n)
    print("by_type", by_type)
    print(msg)
    try:
        doc.Utility.Prompt("\n" + msg + "\n")
    except Exception as ex:
        print("prompt_err", ex)
    return 0


if __name__ == "__main__":
    sys.exit(main())
