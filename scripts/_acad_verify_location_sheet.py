"""Per-sheet verify gate for HWP24QC1 Location_*_SGD.dwg (Screenshots sandbox).

PASS requires (deterministic):
  1. No xref path under pwworking / d017*
  2. Plan Location xref resolves under Screenshots
  3. Steel layers exist with total LF > 0 (Straight+Depressed+Corner)
  4. If ramp/flare blocks present on known layers → Depressed LF > 0

Exit 0 = PASS, 1 = FAIL, 2 = open/COM error.

Usage:
  python scripts/_acad_verify_location_sheet.py --file Screenshots/HWP24QC1_Location_10_SGD.dwg
  python scripts/_acad_verify_location_sheet.py --active
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

import pythoncom
from pywintypes import com_error
from win32com.client import GetObject, Dispatch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _acad_site_zone as site_zone

SHOTS = r"C:\repos\microstation-vba-project\Screenshots"
BUSY = -2147418111

STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]
LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER = STEEL

RAMP_LAYERS = (
    "PROP-D-ROAD-RAMPS",
    "PRO-S-CURB-LINE",
    "PROP-S-CURB-LINE",
    "S-CURB-DROP",
)


def retry(fn, tries=40, delay=0.4):
    last = None
    for _ in range(tries):
        try:
            return fn()
        except com_error as e:
            last = e
            if e.args and e.args[0] == BUSY:
                time.sleep(delay)
                continue
            raise
    raise last  # type: ignore[misc]


def is_forbidden_path(p: str) -> bool:
    n = (p or "").replace("/", "\\").lower()
    return (
        "pwworking" in n
        or "\\d017" in n
        or n.startswith("..\\d017")
        or "/d017" in n.replace("\\", "/")
    )


def list_xrefs(doc):
    out = []
    for j in range(retry(lambda: doc.Blocks.Count)):
        try:
            b = retry(lambda j=j: doc.Blocks.Item(j))
            if not retry(lambda: b.IsXRef):
                continue
            out.append((retry(lambda: b.Name), retry(lambda: (b.Path or ""))))
        except Exception:
            continue
    return out


def is_plan_xref(name: str, path: str) -> bool:
    blob = f"{name}|{path}".replace("/", "\\")
    if "T4318" in blob:
        return False
    return "HWP24QC1_Location" in name or "Location_" in name


def abs_path(host: str, path: str) -> str:
    if not path:
        return ""
    if os.path.isabs(path):
        return os.path.normcase(os.path.abspath(path))
    return os.path.normcase(
        os.path.abspath(os.path.normpath(os.path.join(os.path.dirname(host), path)))
    )


def steel_lf(doc) -> dict[str, float]:
    totals = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    ms = retry(lambda: doc.ModelSpace)
    n = retry(lambda: ms.Count)
    for i in range(n):
        try:
            e = retry(lambda i=i: ms.Item(i))
            lyr = retry(lambda: e.Layer)
            if lyr not in totals:
                continue
            try:
                from win32com.client.dynamic import Dispatch as DD  # early-bound lacks Length
                totals[lyr] += float(retry(lambda: DD(e).Length))
            except Exception:
                # some curve types expose Length differently; count as present
                totals[lyr] += 0.01
        except Exception:
            continue
    return totals


def count_rampish(doc, cells=None) -> int:
    """Ramp/flare signal: PROP-D-ROAD-RAMPS ents, or BlockReference on curb-drop layers.

    cells: survey density grid (site_zone) — off-site template/foreign blocks ignored.
    """
    n = 0
    ms = retry(lambda: doc.ModelSpace)
    count = retry(lambda: ms.Count)
    for i in range(count):
        try:
            e = retry(lambda i=i: ms.Item(i))
            lyr = retry(lambda: e.Layer)
            oname = retry(lambda: e.ObjectName)
            if lyr != "PROP-D-ROAD-RAMPS" and not (
                lyr in ("PRO-S-CURB-LINE", "PROP-S-CURB-LINE", "S-CURB-DROP")
                and "BlockReference" in oname
            ):
                continue
            if cells is not None:
                mn, mx = retry(lambda: e.GetBoundingBox())
                if not site_zone.on_site(cells, (mn[0] + mx[0]) / 2, (mn[1] + mx[1]) / 2):
                    continue
            n += 1
        except Exception:
            continue
    return n


def verify_doc(doc, host_path: str | None = None) -> dict:
    host = host_path or retry(lambda: doc.FullName)
    fails: list[str] = []
    warns: list[str] = []
    xrefs = list_xrefs(doc)
    forbidden = [(n, p) for n, p in xrefs if is_forbidden_path(p)]
    if forbidden:
        fails.append(
            "forbidden_xref_paths:"
            + ",".join(f"{n}=>{p}" for n, p in forbidden)
        )

    plans = [(n, p) for n, p in xrefs if is_plan_xref(n, p)]
    if not plans:
        fails.append("missing_plan_location_xref")
    else:
        for n, p in plans:
            ap = abs_path(host, p)
            if is_forbidden_path(ap) or is_forbidden_path(p):
                fails.append(f"plan_xref_forbidden:{n}=>{p}")
            elif "screenshots" not in ap.lower():
                fails.append(f"plan_xref_not_screenshots:{n}=>{p}")
            elif not os.path.exists(ap if os.path.isabs(p) else abs_path(host, p)):
                # resolve
                rp = p if os.path.isabs(p) else abs_path(host, p)
                if not os.path.exists(rp):
                    fails.append(f"plan_xref_unresolved:{n}=>{p}")

    totals = steel_lf(doc)
    total_steel = sum(totals.values())
    if total_steel <= 0.05:
        fails.append("no_steel_layers_or_lf_zero")

    cells = site_zone.survey_cells(doc, log=lambda m: None)
    if cells is None:
        fails.append("no_survey_xref_site_zone_uncertain")
    ramp_n = count_rampish(doc, cells)
    dep = totals[LAYER_DEPRESSED]
    if ramp_n > 0 and dep <= 0.05:
        fails.append(f"ramps_present_depressed_zero:ramps={ramp_n}")

    status = "PASSED" if not fails else "BLOCKED"
    return {
        "file": os.path.basename(host),
        "path": host,
        "status": status,
        "fails": fails,
        "warns": warns,
        "xrefs": [{"name": n, "path": p} for n, p in xrefs],
        "plan_xrefs": [{"name": n, "path": p} for n, p in plans],
        "steel_lf": {
            "Straight": round(totals[LAYER_STRAIGHT], 2),
            "Depressed": round(totals[LAYER_DEPRESSED], 2),
            "Corner": round(totals[LAYER_CORNER], 2),
            "total": round(total_steel, 2),
        },
        "rampish_count": ramp_n,
    }


def get_acad():
    try:
        return GetObject(None, "AutoCAD.Application")
    except Exception:
        a = Dispatch("AutoCAD.Application")
        a.Visible = True
        return a


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="Path to Location_*_SGD.dwg")
    ap.add_argument("--active", action="store_true", help="Verify ActiveDocument")
    ap.add_argument("--json", action="store_true", help="Print JSON only")
    args = ap.parse_args()
    if not args.file and not args.active:
        ap.error("pass --file PATH or --active")

    pythoncom.CoInitialize()
    acad = retry(lambda: get_acad())
    opened = False
    try:
        if args.active:
            doc = retry(lambda: acad.ActiveDocument)
            host = retry(lambda: doc.FullName)
        else:
            path = os.path.abspath(args.file)
            if not os.path.exists(path):
                print(json.dumps({"status": "ESCALATED", "fails": [f"missing:{path}"]}))
                return 2
            # reuse if open
            doc = None
            for i in range(retry(lambda: acad.Documents.Count)):
                d = retry(lambda i=i: acad.Documents.Item(i))
                if os.path.normcase(retry(lambda: d.FullName)) == os.path.normcase(path):
                    doc = d
                    break
            if doc is None:
                doc = retry(lambda: acad.Documents.Open(path))
                opened = True
            host = path
        result = verify_doc(doc, host)
    except Exception as e:
        result = {
            "file": args.file or "(active)",
            "status": "ESCALATED",
            "fails": [f"com_error:{e}"],
            "warns": [],
        }
        if not args.json:
            print("VERIFY ESCALATED", e)
        else:
            print(json.dumps(result, indent=2))
        return 2
    finally:
        if opened:
            try:
                retry(lambda: doc.Close(False))
            except Exception:
                pass

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"VERIFY {result['file']} → {result['status']}")
        print(f"  steel={result.get('steel_lf')}")
        print(f"  rampish={result.get('rampish_count')}")
        for f in result.get("fails", []):
            print(f"  FAIL {f}")
        for w in result.get("warns", []):
            print(f"  WARN {w}")
    return 0 if result.get("status") == "PASSED" else 1


if __name__ == "__main__":
    sys.exit(main())
