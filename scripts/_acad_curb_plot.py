"""Plot curb QA from DWG geometry (no screen capture / focus steal).

Draws, near --center: curb polylines/FLs (gray, handle labels), ramp flare solids
(yellow), and steel layers (Straight red / Depressed green / Corner blue).
Usage: python scripts/_acad_curb_plot.py --name 11-12 --center X Y --half 30 --out outputs/qa/x.png
"""
from __future__ import annotations

import argparse
import math
import sys
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pythoncom
from win32com.client import GetObject
from win32com.client.dynamic import Dispatch as DD

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _acad_reclassify_loc01_v7_keep as v7

STEEL = {"Straight": "red", "Depressed": "limegreen", "Corner": "blue"}
CURB_LAYERS = ("C-TOPO-FEAT", "PRO-S-CURB-LINE", "PROP-S-CURB-LINE", "PRO-S-CURB-DLINE")
RAMP_LAYERS = ("PROP-D-ROAD-RAMPS", "PRO-S-CURB-LINE", "PROP-S-CURB-LINE", "S-CURB-DROP")


def coords_xy(e):
    on = e.ObjectName
    if on == "AeccDbFeatureLine":
        return None
    c = list(e.Coordinates)
    step = 2 if on == "AcDbPolyline" else 3
    return [(c[i], c[i + 1]) for i in range(0, len(c) - 1, step)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", required=True)
    ap.add_argument("--center", nargs=2, type=float, required=True)
    ap.add_argument("--half", type=float, default=30)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    doc = [acad.Documents.Item(i) for i in range(acad.Documents.Count)
           if a.name.lower() in acad.Documents.Item(i).Name.lower()][0]
    cx, cy, h = a.center[0], a.center[1], a.half
    fig, ax = plt.subplots(figsize=(10, 10))
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = DD(ms.Item(i))
            mn, mx = e.GetBoundingBox()
            if mx[0] < cx - h or mn[0] > cx + h or mx[1] < cy - h or mn[1] > cy + h:
                continue
            L, on = e.Layer, e.ObjectName
            if on == "AcDbLine" and "Steel Faced" in L:
                s, t = e.StartPoint, e.EndPoint
                col = STEEL[L.split()[0]]
                ax.plot([s[0], t[0]], [s[1], t[1]], color=col, lw=3, zorder=5)
            elif L in CURB_LAYERS and on in ("AcDbPolyline", "AcDb3dPolyline", "AeccDbFeatureLine", "AcDbLine", "AcDbArc"):
                if on in ("AcDbLine", "AcDbArc", "AcDbPolyline"):
                    import _acad_batch_classify_locations as bc
                    pts = bc._poly_pts(doc, e.Handle)
                else:
                    pts = coords_xy(e)
                if pts is None:
                    pts = [(p[0], p[1]) for p in v7.sample_fl_vlax(doc, e.Handle, step=0.5)]
                xs, ys = zip(*pts)
                ax.plot(xs, ys, color="gray", lw=1, zorder=3)
                ax.text(xs[len(xs) // 2], ys[len(ys) // 2], e.Handle, fontsize=7, color="black")
            elif L in RAMP_LAYERS and on == "AcDbBlockReference":
                rd = v7.ramp_flare_data(doc, e.Handle)
                for f in rd["flares"]:
                    ax.fill([p[0] for p in f], [p[1] for p in f], color="yellow", ec="olive", zorder=4)
                if rd["flares"]:
                    p = e.InsertionPoint
                    ax.text(p[0], p[1], rd["name"], fontsize=7, color="darkorange")
        except Exception:
            continue
    ax.set_xlim(cx - h, cx + h); ax.set_ylim(cy - h, cy + h); ax.set_aspect("equal")
    ax.set_title(f"{doc.Name}  red=Straight green=Depressed blue=Corner gray=curb yellow=flare")
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    fig.savefig(a.out, dpi=110, bbox_inches="tight")
    print("saved", a.out)


if __name__ == "__main__":
    main()
