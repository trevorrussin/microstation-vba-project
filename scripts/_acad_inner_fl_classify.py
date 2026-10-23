"""Sample AECC feature-line points via vlax-curve; rewrite classified segs on INNER curb."""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

OUT = "C:/Users/RussinT/AppData/Local/Temp/fl_sample.txt"
FL_HANDLES = ["13D822", "13D824"]
FLARE_HANDLES = ["13DABE", "13DABD", "13D9E6", "13D9E7"]
HEX_HANDLES = ["13DAB9", "13D9E2"]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLOR_STRAIGHT = 1
COLOR_DEPRESSED = 3
COLOR_CORNER = 5
PAD = 0.15
STEP = 0.25  # sample spacing ft for densifying, then break at zone changes + key verts


def bbox_of(doc, handle, pad=0.0):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return (
        float(mn[0]) - pad,
        float(mn[1]) - pad,
        float(mx[0]) + pad,
        float(mx[1]) + pad,
    )


def inside(box, x, y):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def classify(mx, my, flares, hexes):
    if any(inside(b, mx, my) for b in flares):
        return LAYER_CORNER
    if any(inside(b, mx, my) for b in hexes):
        return LAYER_DEPRESSED
    return LAYER_STRAIGHT


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    if os.path.exists(OUT):
        os.remove(OUT)
    out_lisp = OUT.replace("\\", "/")

    # Sample both feature lines densely via vlax-curve
    lisp = f"""
(progn
  (vl-load-com)
  (setq f (open "{out_lisp}" "w"))
  (foreach h (list "13D822" "13D824")
    (setq e (handent h))
    (write-line (strcat "HANDLE " h) f)
    (if (and e (vlax-curve-getEndParam e))
      (progn
        (setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e)))
        (write-line (strcat "LEN " (rtos len 2 6)) f)
        (setq d 0.0)
        (while (<= d len)
          (setq p (vlax-curve-getPointAtDist e d))
          (if p
            (write-line
              (strcat (rtos d 2 4) "," (rtos (car p) 2 6) "," (rtos (cadr p) 2 6) ","
                      (rtos (if (caddr p) (caddr p) 0.0) 2 6)) f)
          )
          (setq d (+ d 0.25))
        )
        ; always include end
        (setq p (vlax-curve-getPointAtDist e len))
        (if p
          (write-line
            (strcat (rtos len 2 4) "," (rtos (car p) 2 6) "," (rtos (cadr p) 2 6) ","
                    (rtos (if (caddr p) (caddr p) 0.0) 2 6)) f)
        )
        (write-line "====" f)
      )
      (write-line "NOCURVE" f)
    )
  )
  (close f)
  (princ "SAMPLED")
)
"""
    doc.SendCommand(lisp.replace("\n", " ") + " ")
    time.sleep(2.5)

    if not os.path.exists(OUT):
        raise SystemExit("no sample file")
    text = open(OUT, encoding="utf-8", errors="replace").read()
    print(text[:1500])
    if "NOCURVE" in text and "LEN " not in text:
        raise SystemExit("vlax-curve not supported on feature lines")

    # parse
    fl_pts = {}
    cur = None
    for line in text.splitlines():
        if line.startswith("HANDLE "):
            cur = line.split()[1]
            fl_pts[cur] = []
        elif line.startswith("LEN ") or line == "====" or line == "NOCURVE":
            continue
        elif cur and "," in line:
            parts = line.split(",")
            if len(parts) >= 4:
                fl_pts[cur].append(
                    (float(parts[1]), float(parts[2]), float(parts[3]), float(parts[0]))
                )

    for h, pts in fl_pts.items():
        print(f"{h}: {len(pts)} samples")

    # Delete previously created classified lines on the three layers
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName == "AcDbLine" and e.Layer in (
                LAYER_STRAIGHT,
                LAYER_DEPRESSED,
                LAYER_CORNER,
            ):
                to_del.append(e.Handle)
        except Exception:
            continue
    print(f"removing prior classified lines: {len(to_del)}")
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass

    flares = [bbox_of(doc, h, PAD) for h in FLARE_HANDLES]
    hexes = [bbox_of(doc, h, PAD) for h in HEX_HANDLES]

    # Ensure layers
    for name, col in (
        (LAYER_STRAIGHT, COLOR_STRAIGHT),
        (LAYER_DEPRESSED, COLOR_DEPRESSED),
        (LAYER_CORNER, COLOR_CORNER),
    ):
        found = None
        for i in range(doc.Layers.Count):
            if doc.Layers.Item(i).Name == name:
                found = doc.Layers.Item(i)
                break
        if found is None:
            found = doc.Layers.Add(name)
        found.Color = col

    totals = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    counts = {LAYER_STRAIGHT: 0, LAYER_DEPRESSED: 0, LAYER_CORNER: 0}

    # For each FL: build segments between consecutive samples, merge runs of same class
    for h, samples in fl_pts.items():
        if len(samples) < 2:
            continue
        # merge consecutive samples into classified polyline runs
        runs = []  # each: (layer, [pts])
        for i in range(len(samples) - 1):
            p0 = samples[i]
            p1 = samples[i + 1]
            mx = 0.5 * (p0[0] + p1[0])
            my = 0.5 * (p0[1] + p1[1])
            layer = classify(mx, my, flares, hexes)
            if not runs or runs[-1][0] != layer:
                runs.append([layer, [p0, p1]])
            else:
                runs[-1][1].append(p1)

        print(f"\n{h} runs={len(runs)}")
        for layer, pts in runs:
            # emit LINE per consecutive pair in run (individual segments as requested)
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                length = math.hypot(b[0] - a[0], b[1] - a[1])
                if length < 1e-6:
                    continue
                sp = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (a[0], a[1], a[2]))
                ep = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (b[0], b[1], b[2]))
                line = ms.AddLine(sp, ep)
                line.Layer = layer
                line.Color = 256
                totals[layer] += length
                counts[layer] += 1
            print(f"  {layer.split()[0]:9s} pts={len(pts)} span~{sum(math.hypot(pts[i+1][0]-pts[i][0], pts[i+1][1]-pts[i][1]) for i in range(len(pts)-1)):.2f}")

        # Hide/delete original feature line? User asked to explode it.
        # Since we can't true-explode, delete FL after replacing with segs.
        try:
            doc.HandleToObject(h).Delete()
            print(f"deleted feature line {h}")
        except Exception as ex:
            print(f"could not delete {h}: {ex}")

    doc.Regen(1)
    print("\n===== QUANTITIES (linear feet) — INNER curb =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"  {lyr}")
        print(f"    segments={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"  TOTAL LF={sum(totals.values()):.2f}")


if __name__ == "__main__":
    main()
