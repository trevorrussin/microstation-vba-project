"""Map remaining qty chips [15],[38],[52] via PDF leaders + host centroids."""
from __future__ import annotations

import math
import sys
from pathlib import Path

import fitz
import pythoncom

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import ms_connect

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference

pdf = Path(
    r"C:\Users\RussinT\OneDrive - AECOM\Desktop\Ballys_Hutch_09.11.26_Traffic_r1_BTS.pdf"
)
page = fitz.open(pdf)[2]
annots = list(page.annots() or [])


def annot_verts(a):
    try:
        v = a.vertices
        if v:
            # pymupdf may return flat floats or list of points
            if isinstance(v[0], (int, float)):
                return [(float(v[i]), float(v[i + 1])) for i in range(0, len(v) - 1, 2)]
            out = []
            for p in v:
                if isinstance(p, (list, tuple)) and len(p) >= 2:
                    out.append((float(p[0]), float(p[1])))
            if out:
                return out
    except Exception:
        pass
    r = a.rect
    return [(r.x0, r.y0), (r.x1, r.y1)]


for idx in (15, 38, 52, 0, 43, 51, 44, 45, 47, 49, 58):
    a = annots[idx]
    r = a.rect
    content = ((a.info or {}).get("content") or "").replace("\r", "\n").strip()
    print(
        f"CHIP [{idx}] {content!r} rect=({r.x0:.0f},{r.y0:.0f})-({r.x1:.0f},{r.y1:.0f})"
    )
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    near = []
    for i, b in enumerate(annots):
        t = b.type[1]
        if t not in ("Line", "PolyLine", "Ink"):
            continue
        br = b.rect
        dist = math.hypot((br.x0 + br.x1) / 2 - cx, (br.y0 + br.y1) / 2 - cy)
        if dist > 100:
            continue
        verts = annot_verts(b)
        near.append((dist, i, t, verts))
    near.sort()
    for dist, i, t, verts in near[:6]:
        print(f"  d={dist:.0f} leader[{i}] {t} verts={[(round(x), round(y)) for x, y in verts]}")

HOSTS = [
    3770,
    3778,
    3786,
    5386,
    5474,
    6117,
    6120,
    6126,
    6129,
    6132,
    6135,
    6214,
    6280,
    6792,
    6813,
]
print("\nHOST text centroids + full item lines:")
for hid in HOSTS:
    e = mr.GetElementByID64(hid)
    tn = None
    if e.IsTextNodeElement:
        tn = e
    elif int(e.Type) == 2:
        ee = e.AsComplexElement.GetSubElements()
        while ee.MoveNext():
            if ee.Current.IsTextNodeElement:
                tn = ee.Current
                break
    if not tn:
        continue
    r = tn.Range
    cx = (r.Low.X + r.High.X) / 2
    cy = (r.Low.Y + r.High.Y) / 2
    lines = [
        tn.AsTextNodeElement.TextLine(i)
        for i in range(1, tn.AsTextNodeElement.TextLinesCount + 1)
    ]
    print(f"\n=== {hid} centroid=({cx:.1f},{cy:.1f}) ===")
    for L in lines:
        print(" ", L)
