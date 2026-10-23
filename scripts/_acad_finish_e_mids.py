"""Finish E + MidS INNER only; do not clear existing classified; never touch outers."""
from __future__ import annotations

import importlib.util
import pythoncom
from win32com.client import Dispatch

spec = importlib.util.spec_from_file_location(
    "v5", r"c:\repos\microstation-vba-project\scripts\_acad_reclassify_loc01_v5_inner.py"
)
v5 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(v5)

OUTER = v5.OUTER_HANDLES


def ped_ramp_as_ramp(doc, handle):
    br = doc.HandleToObject(handle)
    mn, mx = br.GetBoundingBox()
    corners = [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]
    return {
        "handle": handle,
        "name": "PED_RAMP",
        "flares": [],
        "dws": corners,
        "area": None,
    }


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    for h in OUTER:
        try:
            doc.HandleToObject(h)
        except Exception:
            print("ABORT outer missing", h)
            return

    totals = {k: 0.0 for k in v5.COLORS}
    counts = {k: 0 for k in v5.COLORS}
    claimed = set()

    ramps = [
        ped_ramp_as_ramp(doc, "133817"),
        v5.hatch_as_ramp(doc, "138B88", "MidS-trail"),
    ]
    # also try short face FL near E if full inner has no hit
    jobs = [
        ("fl", "13CED2", "E"),
        ("fl", "13CED4", "E-face"),
        ("fl", "13D7BF", "MidS"),
        ("fl", "13D7C4", "MidS-face"),
    ]
    for kind, h, label in jobs:
        if h in OUTER:
            print("REFUSE", h)
            continue
        try:
            e = doc.HandleToObject(h)
        except Exception:
            print("missing", h)
            continue
        print(f"\n=== sample {label} {h} ===")
        pts = v5.sample_fl_vlax(doc, h)
        print("n", len(pts))
        if len(pts) < 3:
            continue
        ok, used = v5.classify_clean(ms, pts, ramps, totals, counts, label, claimed)
        if ok:
            claimed.update(used)
            if h not in OUTER:
                try:
                    e.Delete()
                    print("deleted INNER", h)
                except Exception as ex:
                    print("del", ex)

    print("\n=== OUTER intact? ===")
    for h in sorted(OUTER):
        try:
            doc.HandleToObject(h)
            print("OK", h)
        except Exception:
            print("MISSING", h)

    print("\nadded LF", {k.split()[0]: round(v, 2) for k, v in totals.items()})
    print("claimed", claimed)
    doc.Regen(1)


if __name__ == "__main__":
    main()
