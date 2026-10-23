"""Restore color-coded SGD hosts from Screenshots backup/classified, then
reattach Screenshots Location_XX xrefs (no pwworking/d017).

Steel lived on SGD hosts; refreshing from d017 wiped it. Backup still has steel.
"""
from __future__ import annotations

import glob
import os
import shutil
import sys
import time

import pythoncom
from pywintypes import com_error
from win32com.client import GetObject, Dispatch, VARIANT

SHOTS = r"C:\repos\microstation-vba-project\Screenshots"
BACKUP_GLOB = os.path.join(SHOTS, "_sgd_backup_before_restore_*")
CLASSIFIED = os.path.join(SHOTS, "classified")
BUSY = -2147418111
ORIGIN = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0.0, 0.0, 0.0))

STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]


def retry(fn, tries=50, delay=0.4):
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


def is_forbidden(p: str) -> bool:
    n = (p or "").replace("/", "\\").lower()
    return "pwworking" in n or "\\d017" in n or n.startswith("..\\d017")


def latest_backup_dir() -> str | None:
    dirs = sorted(glob.glob(BACKUP_GLOB), key=os.path.getmtime, reverse=True)
    return dirs[0] if dirs else None


def find_restore_source(sgd_name: str) -> str | None:
    """Prefer classified/, then newest _sgd_backup_* copy."""
    c = os.path.join(CLASSIFIED, sgd_name)
    if os.path.exists(c):
        return c
    bdir = latest_backup_dir()
    if bdir:
        p = os.path.join(bdir, sgd_name)
        if os.path.exists(p):
            return p
    return None


def match_location(sgd_name: str) -> str | None:
    tok = sgd_name.replace("_SGD.dwg", "").replace("HWP24QC1_Location_", "", 1)
    specials = {
        "05-08": "HWP24QC1_Location_05_08.dwg",
        "19": "HWP24QC1_Location_19-20.dwg",
        "20": "HWP24QC1_Location_19-20.dwg",
        "90_91": "HWP24QC1_Location_90-91.dwg",
        "89_96": "HWP24QC1_Location_89_96-98.dwg",
    }
    cands = []
    if tok in specials:
        cands.append(specials[tok])
    cands += [
        f"HWP24QC1_Location_{tok}.dwg",
        f"HWP24QC1_Location_{tok.replace('-', '_')}.dwg",
        f"HWP24QC1_Location_{tok.replace('_', '-')}.dwg",
    ]
    for c in cands:
        p = os.path.join(SHOTS, c)
        if os.path.exists(p):
            return p
    return None


def has_steel_layers_file(path: str, acad) -> bool:
    """Quick open check for steel layer names (not LF)."""
    doc = retry(lambda: acad.Documents.Open(path))
    try:
        names = set()
        for i in range(retry(lambda: doc.Layers.Count)):
            try:
                names.add(retry(lambda i=i: doc.Layers.Item(i).Name))
            except Exception:
                pass
        return all(s in names for s in STEEL)
    finally:
        retry(lambda: doc.Close(False))


def attach_plan(doc, loc_path: str) -> str:
    block = os.path.splitext(os.path.basename(loc_path))[0]
    # repath existing plan xref or attach
    for j in range(retry(lambda: doc.Blocks.Count)):
        try:
            b = retry(lambda j=j: doc.Blocks.Item(j))
            if not retry(lambda: b.IsXRef):
                continue
            name = retry(lambda: b.Name)
            path = retry(lambda: b.Path or "")
            if "T4318" in name or "T4318" in path:
                # survey → Screenshots if present
                base = os.path.basename(path.replace("/", "\\")) if path else f"{name}.dwg"
                if not base.lower().endswith(".dwg"):
                    base = f"{name}.dwg"
                surv = os.path.join(SHOTS, base)
                if os.path.exists(surv) and not is_forbidden(surv):
                    if os.path.normcase(path) != os.path.normcase(surv):
                        retry(lambda: setattr(b, "Path", surv) or True)
                continue
            if "Location" in name:
                exact = os.path.join(SHOTS, name + ".dwg")
                target = exact if os.path.exists(exact) else loc_path
                if is_forbidden(target):
                    continue
                retry(lambda: setattr(b, "Path", target) or True)
                return f"repath:{name}"
        except Exception:
            continue
    retry(
        lambda: doc.ModelSpace.AttachExternalReference(
            loc_path, block, ORIGIN, 1.0, 1.0, 1.0, 0.0, True
        )
    )
    return f"attach:{block}"


def restore_one(acad, sgd_name: str, *, force_backup: bool = False) -> dict:
    dest = os.path.join(SHOTS, sgd_name)
    src = find_restore_source(sgd_name)
    if not src:
        return {"file": sgd_name, "ok": False, "error": "no_backup_or_classified"}
    loc = match_location(sgd_name)
    if not loc:
        return {"file": sgd_name, "ok": False, "error": "no_screenshots_location"}

    # close if open
    for i in range(retry(lambda: acad.Documents.Count) - 1, -1, -1):
        d = retry(lambda i=i: acad.Documents.Item(i))
        if os.path.normcase(retry(lambda: d.FullName)) == os.path.normcase(dest):
            retry(lambda: d.Close(False))

    shutil.copy2(src, dest)
    doc = retry(lambda: acad.Documents.Open(dest))
    try:
        note = attach_plan(doc, loc)
        retry(lambda: doc.Save())
        return {
            "file": sgd_name,
            "ok": True,
            "from": src,
            "location": loc,
            "xref": note,
        }
    finally:
        retry(lambda: doc.Close(True))


def list_sgd_names() -> list[str]:
    return sorted(
        f
        for f in os.listdir(SHOTS)
        if f.endswith("_SGD.dwg")
        and f.startswith("HWP24QC1_Location_")
        and "01-02" not in f
        and "Location_01" not in f
    )


def main(argv: list[str] | None = None) -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="Substring e.g. 05-08 or 10")
    ap.add_argument("--all-with-backup", action="store_true")
    args = ap.parse_args(argv)

    names = list_sgd_names()
    if args.only:
        names = [n for n in names if args.only.upper() in n.upper()]
    elif args.all_with_backup:
        names = [n for n in names if find_restore_source(n)]
    else:
        ap.error("pass --only SUBSTR or --all-with-backup")

    pythoncom.CoInitialize()
    try:
        acad = GetObject(None, "AutoCAD.Application")
    except Exception:
        acad = Dispatch("AutoCAD.Application")
        acad.Visible = True

    ok = fail = 0
    for name in names:
        print(f"--- restore {name}")
        row = restore_one(acad, name)
        print(" ", row)
        if row.get("ok"):
            ok += 1
        else:
            fail += 1
    print(f"DONE restore ok={ok} fail={fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
