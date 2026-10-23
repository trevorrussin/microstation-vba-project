"""
Attach Screenshots Location_XX.dwg as xref inside each Screenshots Location_XX_SGD.dwg.

PASS gate: plan xref resolves to Screenshots Location_*.dwg.
Screenshots-only — never read/write pwworking or d017* for xrefs or copies.
Survey uses T4318*.dwg already present in Screenshots.
"""
from __future__ import annotations

import os
import sys
import time

import pythoncom
from pywintypes import com_error
from win32com.client import GetObject, VARIANT

SHOTS = r"C:\repos\microstation-vba-project\Screenshots"
BUSY = -2147418111
ORIGIN = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0.0, 0.0, 0.0))


def is_loc01(name: str) -> bool:
    return (
        "Location_01" in name
        or "01-02_50-52" in name
        or "01-02,50-52" in name
    )


def retry(fn, tries=80, delay=0.4):
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


def match_location(sgd_name: str, loc_files: dict[str, str]) -> str | None:
    tok = sgd_name.replace("_SGD.dwg", "").replace("HWP24QC1_Location_", "", 1)
    if tok in ("19", "20"):
        return loc_files.get("HWP24QC1_Location_19-20.dwg")
    specials = {
        "05-08": "HWP24QC1_Location_05_08.dwg",
        "90_91": "HWP24QC1_Location_90-91.dwg",
        "89_96": "HWP24QC1_Location_89_96-98.dwg",
    }
    if tok in specials and specials[tok] in loc_files:
        return loc_files[specials[tok]]
    for c in (
        f"HWP24QC1_Location_{tok}.dwg",
        f"HWP24QC1_Location_{tok.replace('-', '_')}.dwg",
        f"HWP24QC1_Location_{tok.replace('_', '-')}.dwg",
    ):
        if c in loc_files:
            return loc_files[c]
    return None


def list_xrefs(doc):
    out = []
    for j in range(retry(lambda: doc.Blocks.Count)):
        try:
            b = retry(lambda j=j: doc.Blocks.Item(j))
            if not retry(lambda: b.IsXRef):
                continue
            out.append((retry(lambda: b.Name), retry(lambda: (b.Path or "")), b))
        except Exception:
            continue
    return out


def is_plan(name: str, path: str) -> bool:
    blob = f"{name}|{path}".replace("/", "\\")
    if "T4318" in blob:
        return False
    return ("HWP24QC1_Location" in name) or ("Location_" in name)


def is_survey(name: str, path: str) -> bool:
    return "T4318" in f"{name}|{path}".replace("/", "\\")


def abs_xref(host: str, path: str) -> str:
    if not path:
        return ""
    if os.path.isabs(path):
        return os.path.normcase(os.path.abspath(path))
    return os.path.normcase(
        os.path.abspath(os.path.normpath(os.path.join(os.path.dirname(host), path)))
    )


def plan_ok(host: str, doc, want: str) -> bool:
    if is_forbidden_path(want):
        return False
    want_n = os.path.normcase(os.path.abspath(want))
    for name, path, _ in list_xrefs(doc):
        if not is_plan(name, path) or is_forbidden_path(path):
            continue
        got = abs_xref(host, path)
        if got == want_n and os.path.exists(want):
            return True
        exact = os.path.join(SHOTS, name + ".dwg")
        if os.path.exists(exact) and got == os.path.normcase(os.path.abspath(exact)):
            return True
    return False


def thaw_all(doc) -> int:
    n = 0
    try:
        count = int(doc.Layers.Count)
    except Exception:
        return 0
    for i in range(count):
        try:
            ly = doc.Layers.Item(i)
            try:
                if not ly.LayerOn:
                    ly.LayerOn = True
                    n += 1
            except Exception:
                pass
            try:
                if ly.Freeze:
                    ly.Freeze = False
                    n += 1
            except Exception:
                pass
        except Exception:
            pass
    return n


def screenshots_survey_files() -> dict[str, str]:
    out: dict[str, str] = {}
    for fn in os.listdir(SHOTS):
        if fn.startswith("T4318") and fn.lower().endswith(".dwg"):
            out[fn] = os.path.join(SHOTS, fn)
    return out


def fix_one(acad, host: str, want_loc: str, survey_files: dict[str, str]) -> tuple[bool, str]:
    if is_forbidden_path(want_loc) or not want_loc.lower().startswith(SHOTS.lower()):
        return False, "refuse_non_screenshots_location"
    doc = retry(lambda: acad.Documents.Open(host))
    notes: list[str] = []
    try:
        xrefs = list_xrefs(doc)
        plans = [(n, p, b) for n, p, b in xrefs if is_plan(n, p)]
        survs = [(n, p, b) for n, p, b in xrefs if is_survey(n, p)]

        if plans:
            for name, path, b in plans:
                target = want_loc
                exact = os.path.join(SHOTS, name + ".dwg")
                if os.path.exists(exact):
                    target = exact
                if is_forbidden_path(target):
                    notes.append(f"refuse:{name}")
                    continue
                if abs_xref(host, path) != os.path.normcase(os.path.abspath(target)):
                    retry(lambda: setattr(b, "Path", target) or True)
                    notes.append(f"repath:{name}")
        else:
            block = os.path.splitext(os.path.basename(want_loc))[0]
            retry(
                lambda: doc.ModelSpace.AttachExternalReference(
                    want_loc, block, ORIGIN, 1.0, 1.0, 1.0, 0.0, True
                )
            )
            notes.append(f"attach:{block}")

        for name, path, b in survs:
            base = os.path.basename(path.replace("/", "\\")) if path else f"{name}.dwg"
            if not base.lower().endswith(".dwg"):
                base = f"{name}.dwg"
            target = survey_files.get(base) or survey_files.get(f"{name}.dwg")
            if not target:
                for k, v in survey_files.items():
                    if name in k:
                        target = v
                        break
            if not target:
                notes.append(f"survey_missing:{name}")
                continue
            if is_forbidden_path(target):
                notes.append(f"refuse_survey:{name}")
                continue
            if abs_xref(host, path) != os.path.normcase(os.path.abspath(target)):
                retry(lambda: setattr(b, "Path", target) or True)
                notes.append(f"survey:{name}")

        notes.append(f"thaw:{thaw_all(doc)}")
        ok = plan_ok(host, doc, want_loc)
        retry(lambda: doc.Save())
        return ok, ",".join(notes)
    finally:
        try:
            retry(lambda: doc.Close(True))
        except Exception:
            try:
                retry(lambda: doc.Close(False))
            except Exception:
                pass


def close_shots_sgds(acad):
    for i in range(retry(lambda: acad.Documents.Count) - 1, -1, -1):
        d = retry(lambda i=i: acad.Documents.Item(i))
        full = retry(lambda: d.FullName)
        name = retry(lambda: d.Name)
        if "Screenshots" in full and name.endswith("_SGD.dwg"):
            try:
                retry(lambda: d.Close(True))
            except Exception:
                retry(lambda: d.Close(False))


def main() -> int:
    survey_files = screenshots_survey_files()
    loc_files = {
        f: os.path.join(SHOTS, f)
        for f in os.listdir(SHOTS)
        if f.startswith("HWP24QC1_Location_")
        and f.endswith(".dwg")
        and "_SGD" not in f
    }

    pythoncom.CoInitialize()
    acad = retry(lambda: GetObject(None, "AutoCAD.Application"))
    close_shots_sgds(acad)

    sgds = sorted(
        f for f in os.listdir(SHOTS) if f.endswith("_SGD.dwg") and not is_loc01(f)
    )
    print(f"FIX_COUNT={len(sgds)} survey_in_shots={len(survey_files)}")
    ok_n = fail_n = 0
    fails: list[str] = []

    for fn in sgds:
        host = os.path.join(SHOTS, fn)
        want = match_location(fn, loc_files)
        print(f"--- {fn}")
        if not want:
            print("  FAIL NO_LOCATION")
            fail_n += 1
            fails.append(fn)
            continue
        print(f"  want {os.path.basename(want)}")
        try:
            ok, notes = fix_one(acad, host, want, survey_files)
            print(f"  {'PASS' if ok else 'FAIL'} {notes}")
            if ok:
                ok_n += 1
            else:
                fail_n += 1
                fails.append(fn)
        except Exception as e:
            print(f"  FAIL {e}")
            fail_n += 1
            fails.append(fn)

    print("==== SAMPLE_VERIFY ====")
    for sample, want_name in (
        ("HWP24QC1_Location_05-08_SGD.dwg", "HWP24QC1_Location_05_08.dwg"),
        ("HWP24QC1_Location_88_SGD.dwg", "HWP24QC1_Location_88.dwg"),
        ("HWP24QC1_Location_95_SGD.dwg", "HWP24QC1_Location_95.dwg"),
    ):
        host = os.path.join(SHOTS, sample)
        want = os.path.join(SHOTS, want_name)
        if not os.path.exists(host):
            print(f"  SKIP missing {sample}")
            continue
        doc = retry(lambda: acad.Documents.Open(host))
        try:
            ok = plan_ok(host, doc, want)
            xrefs = [f"{n}=>{p}" for n, p, _ in list_xrefs(doc) if is_plan(n, p)]
            bad = [p for _, p, _ in list_xrefs(doc) if is_forbidden_path(p)]
            print(
                f"  {sample} {'PASS' if ok and not bad else 'FAIL'} "
                f"layers={retry(lambda: doc.Layers.Count)} plan={xrefs} forbidden={bad}"
            )
            if not ok or bad:
                fail_n += 1
        finally:
            retry(lambda: doc.Close(False))

    print(f"DONE ok={ok_n} fail={fail_n}")
    if fails:
        print("FAILS " + ", ".join(fails))
    return 0 if fail_n == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
