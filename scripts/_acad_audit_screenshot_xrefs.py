"""Report which Screenshots HWP24QC1 Location DWGs have unresolved xrefs / thin hosts."""
from __future__ import annotations

import glob
import os
import time

import pythoncom
from win32com.client import Dispatch, GetObject

SCREEN = r"c:\repos\microstation-vba-project\Screenshots"
KEEP = ("01-02", "RECOVERED", "FROM_AUTOSAVE")


def main() -> None:
    pythoncom.CoInitialize()
    try:
        acad = GetObject(None, "AutoCAD.Application")
    except Exception:
        acad = Dispatch("AutoCAD.Application")
    acad.Visible = True

    # close non-keep leftovers first
    for _ in range(6):
        closed = False
        try:
            n = acad.Documents.Count
        except Exception:
            time.sleep(1)
            continue
        for i in range(n - 1, -1, -1):
            try:
                d = acad.Documents.Item(i)
                name = d.Name
            except Exception:
                continue
            if any(k.upper() in name.upper() for k in KEEP):
                continue
            if "Location_" in name and "_SGD" in name.upper():
                try:
                    d.Close(False)
                    closed = True
                except Exception:
                    pass
                time.sleep(0.3)
        if not closed:
            break

    paths = sorted(
        p
        for p in glob.glob(os.path.join(SCREEN, "HWP24QC1_Location_*.dwg"))
        if "CLASSIFIED" not in os.path.basename(p).upper()
        and "01-02" not in os.path.basename(p)
    )
    print(f"scanning {len(paths)} Screenshots Location DWGs\n")
    rows = []
    for path in paths:
        name = os.path.basename(path)
        size_mb = os.path.getsize(path) / (1024 * 1024)
        doc = None
        try:
            doc = acad.Documents.Open(path, False)  # read-only prefer False for COM
            time.sleep(0.4)
            ms = doc.ModelSpace.Count
            layers = doc.Layers.Count
            xrefs = []
            missing = []
            for i in range(doc.Blocks.Count):
                try:
                    b = doc.Blocks.Item(i)
                    if not b.IsXRef:
                        continue
                    xp = b.Path or ""
                    cand = os.path.normpath(os.path.join(os.path.dirname(path), xp))
                    ok = os.path.isfile(cand)
                    xrefs.append((b.Name, xp, ok, cand))
                    if not ok:
                        missing.append((b.Name, xp))
                except Exception:
                    continue
            status = "OK"
            if missing:
                status = "MISSING_XREF"
            elif ms < 200 and size_mb < 1.5:
                status = "THIN_HOST"
            rows.append(
                {
                    "file": name,
                    "status": status,
                    "mb": round(size_mb, 2),
                    "ms": ms,
                    "layers": layers,
                    "nxref": len(xrefs),
                    "nmiss": len(missing),
                    "missing": missing,
                }
            )
            flag = status
            print(
                f"{flag:12s}  ms={ms:5d}  layers={layers:3d}  "
                f"{size_mb:5.2f}MB  xrefs={len(xrefs)} miss={len(missing)}  {name}"
            )
            for bn, xp in missing[:4]:
                print(f"             missing {bn} <- {xp}")
        except Exception as ex:
            rows.append(
                {
                    "file": name,
                    "status": "OPEN_FAIL",
                    "mb": round(size_mb, 2),
                    "ms": None,
                    "layers": None,
                    "nxref": None,
                    "nmiss": None,
                    "missing": [],
                    "error": str(ex),
                }
            )
            print(f"OPEN_FAIL     {size_mb:5.2f}MB  {name}  {ex}")
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
                time.sleep(0.6)

    miss = [r for r in rows if r["status"] in ("MISSING_XREF", "THIN_HOST", "OPEN_FAIL")]
    ok = [r for r in rows if r["status"] == "OK"]
    print(f"\n===== SUMMARY ok={len(ok)} problem={len(miss)} total={len(rows)} =====")
    print("PROBLEM FILES:")
    for r in miss:
        print(f"  {r['status']:12s}  {r['file']}  ms={r['ms']}  mb={r['mb']}  miss={r['nmiss']}")

    out = os.path.join(
        r"c:\repos\microstation-vba-project\outputs", "curb-batch-xref-audit.txt"
    )
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"ok={len(ok)} problem={len(miss)} total={len(rows)}\n")
        for r in rows:
            f.write(
                f"{r['status']}\t{r['file']}\tms={r['ms']}\tmb={r['mb']}\t"
                f"xrefs={r['nxref']}\tmiss={r['nmiss']}\n"
            )
            for bn, xp in r.get("missing") or []:
                f.write(f"  MISSING\t{bn}\t{xp}\n")
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
