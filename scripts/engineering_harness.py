"""Engineering harness — single entrypoint for Loc curb classify sandbox.

Screenshots-only. Enforces Failure Protocol states on disk.
Evidence-proportional sampling by default.

Commands:
  python scripts/engineering_harness.py status
  python scripts/engineering_harness.py verify --sample 5
  python scripts/engineering_harness.py verify --all
  python scripts/engineering_harness.py verify --sheet 10
  python scripts/engineering_harness.py restore-steel --sample 5
  python scripts/engineering_harness.py restore-steel --all-with-backup
  python scripts/engineering_harness.py classify --sheet 10
  python scripts/engineering_harness.py classify --sample 3
  python scripts/engineering_harness.py classify --all

Does NOT advance past BLOCKED. Does NOT touch pwworking/d017*.
"""
from __future__ import annotations

import argparse
import json
import os
import random
import subprocess
import sys
import time
from datetime import datetime

import pythoncom
from pywintypes import com_error
from win32com.client import GetObject, Dispatch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _acad_verify_location_sheet as verify_mod
import _acad_restore_sgd_steel as restore_mod
import _eng_harness_manifest as manifest

SHOTS = r"C:\repos\microstation-vba-project\Screenshots"
SCRIPTS = os.path.dirname(os.path.abspath(__file__))
BUSY = -2147418111


def log(msg: str) -> None:
    print(msg, flush=True)


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


def get_acad():
    try:
        return GetObject(None, "AutoCAD.Application")
    except Exception:
        a = Dispatch("AutoCAD.Application")
        a.Visible = True
        return a


def list_sgd() -> list[str]:
    return sorted(
        f
        for f in os.listdir(SHOTS)
        if f.endswith("_SGD.dwg")
        and f.startswith("HWP24QC1_Location_")
        and "01-02" not in f
        and "Location_01" not in f
    )


def pick_sample(names: list[str], n: int, seed: int = 30) -> list[str]:
    """Representative + random + known high-risk (not only first-N)."""
    if n <= 0 or n >= len(names):
        return list(names)
    prefer = []
    for key in ("05-08", "09", "10", "11-12", "58-60", "95"):
        for f in names:
            if key.upper() in f.upper() and f not in prefer:
                prefer.append(f)
    rng = random.Random(seed)
    rest = [f for f in names if f not in prefer]
    rng.shuffle(rest)
    out = []
    # edges: first/last alphabetically among full set
    if names:
        out.append(names[0])
        if names[-1] not in out:
            out.append(names[-1])
    for f in prefer + rest:
        if f not in out:
            out.append(f)
        if len(out) >= n:
            break
    return out[:n]


def ensure_not_blocked(data: dict, *, allow_same: str | None = None) -> None:
    blocked = manifest.any_blocked(data)
    if blocked and blocked != allow_same:
        raise SystemExit(
            f"BLOCKED: {blocked} — Failure Protocol: do not advance. "
            f"Diagnose/fix that sheet (or mark ESCALATED) before NEXT.\n"
            f"manifest: {manifest.HARNESS_MANIFEST}"
        )


def cmd_status(_: argparse.Namespace) -> int:
    data = manifest.load()
    c = manifest.counts(data)
    log(f"harness manifest: {manifest.HARNESS_MANIFEST}")
    log(f"updated: {data.get('updated')}")
    log(f"counts: {c}")
    blocked = manifest.any_blocked(data)
    if blocked:
        log(f"last_blocked: {blocked}")
        log(json.dumps(data["sheets"].get(blocked, {}), indent=2))
    else:
        log("no BLOCKED sheets")
    # sample-based note
    n = sum(c.values())
    if n:
        log(
            f"(sheet rows tracked={n}; conclusion scope is whatever commands "
            f"last wrote — see each sheet.status)"
        )
    return 0


def _verify_one(acad, file_name: str, data: dict) -> dict:
    path = os.path.join(SHOTS, file_name)
    manifest.set_sheet(data, file_name, "VERIFYING")
    manifest.save(data)
    opened = False
    doc = None
    try:
        for i in range(retry(lambda: acad.Documents.Count)):
            d = retry(lambda i=i: acad.Documents.Item(i))
            if os.path.normcase(retry(lambda: d.FullName)) == os.path.normcase(path):
                doc = d
                break
        if doc is None:
            doc = retry(lambda: acad.Documents.Open(path))
            opened = True
        result = verify_mod.verify_doc(doc, path)
    except Exception as e:
        result = {
            "file": file_name,
            "status": "ESCALATED",
            "fails": [f"com_error:{e}"],
        }
    finally:
        if opened and doc is not None:
            try:
                retry(lambda: doc.Close(False))
            except Exception:
                pass

    st = result.get("status", "BLOCKED")
    if st == "PASSED":
        manifest.set_sheet(data, file_name, "PASSED", detail=result)
    elif st == "ESCALATED":
        manifest.set_sheet(data, file_name, "ESCALATED", detail=result)
    else:
        manifest.set_sheet(data, file_name, "BLOCKED", detail=result)
    manifest.save(data)
    return result


def cmd_verify(args: argparse.Namespace) -> int:
    names = list_sgd()
    if args.sheet:
        names = [f for f in names if args.sheet.upper() in f.upper()]
        if not names:
            log(f"no sheet match {args.sheet}")
            return 2
    elif args.all:
        pass
    else:
        n = args.sample if args.sample else 5
        names = pick_sample(names, n)
        log(
            f"EVIDENCE: sample n={len(names)} (not exhaustive). "
            f"Expand only if mixed / high-stakes / --all."
        )

    data = manifest.load()
    pythoncom.CoInitialize()
    acad = retry(lambda: get_acad())
    results = []
    for fn in names:
        log(f"--- verify {fn}")
        r = _verify_one(acad, fn, data)
        log(f"  {r.get('status')} steel={r.get('steel_lf')} fails={r.get('fails')}")
        results.append(r)
        if r.get("status") == "BLOCKED" and not args.continue_on_fail:
            log("BATCH STOP — BLOCKED (Failure Protocol)")
            break

    passed = sum(1 for r in results if r.get("status") == "PASSED")
    blocked = sum(1 for r in results if r.get("status") == "BLOCKED")
    esc = sum(1 for r in results if r.get("status") == "ESCALATED")
    log(f"DONE verify passed={passed} blocked={blocked} escalated={esc}")
    if not args.all and len(results) < len(list_sgd()):
        log(
            f"SAMPLE-BASED: inspected {len(results)} of {len(list_sgd())} SGD hosts. "
            f"Did not open the rest because "
            + (
                "results were consistent PASSED."
                if blocked == 0 and esc == 0 and passed == len(results)
                else "stop-on-block or mixed — expand if needed."
            )
        )
    return 0 if blocked == 0 and esc == 0 else 1


def cmd_restore(args: argparse.Namespace) -> int:
    data = manifest.load()
    names = list_sgd()
    if args.sheet:
        names = [f for f in names if args.sheet.upper() in f.upper()]
    elif args.all_with_backup:
        names = [f for f in names if restore_mod.find_restore_source(f)]
    else:
        n = args.sample if args.sample else 5
        # prefer those with backup
        with_bak = [f for f in names if restore_mod.find_restore_source(f)]
        names = pick_sample(with_bak or names, n)
        log(f"EVIDENCE: restore sample n={len(names)}")

    pythoncom.CoInitialize()
    acad = retry(lambda: get_acad())
    ok = fail = 0
    for fn in names:
        log(f"--- restore {fn}")
        row = restore_mod.restore_one(acad, fn)
        log(f"  {row}")
        if not row.get("ok"):
            fail += 1
            manifest.set_sheet(data, fn, "ESCALATED", detail=row)
            manifest.save(data)
            if not args.continue_on_fail:
                break
            continue
        # verify after restore
        r = _verify_one(acad, fn, data)
        log(f"  post-verify {r.get('status')} {r.get('fails')}")
        if r.get("status") == "PASSED":
            ok += 1
        else:
            fail += 1
            if not args.continue_on_fail:
                break
    log(f"DONE restore+verify ok={ok} fail={fail}")
    return 0 if fail == 0 else 1


def cmd_classify(args: argparse.Namespace) -> int:
    data = manifest.load()
    ensure_not_blocked(data, allow_same=None)

    # Build batch args — always Screenshots SGD via existing batch script
    batch = os.path.join(SCRIPTS, "_acad_batch_classify_locations.py")
    cmd = [sys.executable, "-u", batch]
    if args.sheet:
        cmd += ["--only", args.sheet]
    elif args.all:
        cmd += ["--all"]
    else:
        n = args.sample if args.sample else 3
        names = pick_sample(list_sgd(), n)
        log(f"EVIDENCE: classify sample {names}")
        # batch --only is substring; run one-by-one for sample control
        rc_all = 0
        for fn in names:
            ensure_not_blocked(data)
            manifest.set_sheet(data, fn, "IN_PROGRESS")
            manifest.save(data)
            only = fn.replace("HWP24QC1_Location_", "").replace("_SGD.dwg", "")
            one = [sys.executable, "-u", batch, "--only", only]
            if args.force:
                one.append("--force")
            log(f"RUN {' '.join(one)}")
            rc = subprocess.call(one, cwd=os.path.dirname(SCRIPTS))
            # verify after classify
            pythoncom.CoInitialize()
            acad = retry(lambda: get_acad())
            r = _verify_one(acad, fn, data)
            log(f"  post-classify verify {r.get('status')}")
            if r.get("status") != "PASSED" or rc != 0:
                manifest.set_sheet(
                    data,
                    fn,
                    "BLOCKED",
                    detail={"classify_rc": rc, "verify": r},
                )
                manifest.save(data)
                log("BATCH STOP — BLOCKED")
                return 1
            rc_all = rc_all or rc
        log(
            f"SAMPLE-BASED classify: {len(names)} of {len(list_sgd())} sheets. "
            f"Consistent PASSED — not running --all unless asked."
        )
        return rc_all

    if args.force:
        cmd.append("--force")
    # mark targets in_progress
    targets = list_sgd()
    if args.sheet:
        targets = [f for f in targets if args.sheet.upper() in f.upper()]
    for fn in targets:
        manifest.set_sheet(data, fn, "IN_PROGRESS")
    manifest.save(data)

    log(f"RUN {' '.join(cmd)}")
    rc = subprocess.call(cmd, cwd=os.path.dirname(SCRIPTS))

    # After batch, verify sheets that batch touched (from curb-batch-manifest)
    batch_man = os.path.join(
        r"C:\repos\microstation-vba-project\outputs", "curb-batch-manifest.json"
    )
    pythoncom.CoInitialize()
    acad = retry(lambda: get_acad())
    if os.path.exists(batch_man):
        with open(batch_man, encoding="utf-8") as f:
            bm = json.load(f)
        for row in bm.get("summary", []):
            fn = row.get("file")
            if not fn:
                continue
            if row.get("status") in ("verify_fail", "error"):
                manifest.set_sheet(data, fn, "BLOCKED", detail=row)
            elif row.get("status") in ("ok", "skip_done"):
                r = _verify_one(acad, fn, data)
                if r.get("status") != "PASSED":
                    log(f"  harness verify BLOCKED after batch ok: {fn}")
            manifest.save(data)
        if bm.get("blocked"):
            bf = bm["blocked"].get("file")
            if bf:
                manifest.set_sheet(data, bf, "BLOCKED", detail=bm["blocked"])
                manifest.save(data)
                return 1
    return rc


def cmd_mark(args: argparse.Namespace) -> int:
    """Engineer override after manual QA (does not skip verify by default)."""
    data = manifest.load()
    names = [f for f in list_sgd() if args.sheet.upper() in f.upper()]
    if not names:
        log(f"no sheet match {args.sheet}")
        return 2
    for fn in names:
        if args.passed:
            if not args.force:
                log(f"refusing mark PASSED without --force (run verify first): {fn}")
                return 2
            manifest.set_sheet(
                data,
                fn,
                "PASSED",
                detail={"marked_by": "engineer", "when": datetime.now().isoformat()},
            )
            log(f"marked PASSED {fn}")
        elif args.escalated:
            manifest.set_sheet(
                data,
                fn,
                "ESCALATED",
                detail={"reason": args.reason or "engineer"},
            )
            log(f"marked ESCALATED {fn}")
        elif args.clear_blocked:
            row = data.get("sheets", {}).get(fn, {})
            if row.get("status") == "BLOCKED":
                manifest.set_sheet(data, fn, "NOT_STARTED", detail={"cleared_block": True})
                log(f"cleared BLOCKED → NOT_STARTED {fn}")
    manifest.save(data)
    return 0


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="engineering_harness",
        description="Engineering harness for Screenshots Loc curb classify",
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("status", help="Show harness manifest state")
    p.set_defaults(func=cmd_status)

    p = sub.add_parser("verify", help="Deterministic per-sheet done gate")
    p.add_argument("--sample", type=int, default=0, help="Sample size (default 5 if no --all/--sheet)")
    p.add_argument("--all", action="store_true")
    p.add_argument("--sheet", help="Substring e.g. 10 or 05-08")
    p.add_argument("--continue-on-fail", action="store_true")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser(
        "restore-steel",
        help="Restore SGD from backup/classified + reattach Screenshots Location xref",
    )
    p.add_argument("--sample", type=int, default=0)
    p.add_argument("--all-with-backup", action="store_true")
    p.add_argument("--sheet", help="Substring")
    p.add_argument("--continue-on-fail", action="store_true")
    p.set_defaults(func=cmd_restore)

    p = sub.add_parser("classify", help="Run curb classify under harness gates")
    p.add_argument("--sample", type=int, default=0, help="Default 3 if no --all/--sheet")
    p.add_argument("--all", action="store_true")
    p.add_argument("--sheet", help="Substring")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_classify)

    p = sub.add_parser("mark", help="Engineer override of sheet state")
    p.add_argument("--sheet", required=True)
    p.add_argument("--passed", action="store_true")
    p.add_argument("--escalated", action="store_true")
    p.add_argument("--clear-blocked", action="store_true")
    p.add_argument("--force", action="store_true", help="Required with --passed")
    p.add_argument("--reason", default="")
    p.set_defaults(func=cmd_mark)

    return ap


def main(argv: list[str] | None = None) -> int:
    ap = build_parser()
    args = ap.parse_args(argv)
    return int(args.func(args) or 0)


if __name__ == "__main__":
    sys.exit(main())
