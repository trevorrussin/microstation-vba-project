"""
Batch MicroStation key-in probe + registry promote.

Why this exists
---------------
Growing command-registry.tsv to cover MicroStation's key-in surface cannot
be done one conversational turn at a time. Parallel *research* agents can
append candidate rows to Data/keyin-candidates.tsv; a single probe process
must hit MicroStation serially (one COM session — parallel SendKeyin races
tool state and is unsafe).

Workflow
--------
1. Harvest (parallel-safe): append rows to Data/keyin-candidates.tsv
2. Probe  (serial):         python scripts/keyin_probe.py
3. Promote:                 python scripts/keyin_promote.py

Kinds
-----
settings / view / lock  -> OK probe result can become verified-headless-safe
tool / datapoint / dialog / file -> always unsafe-blocked (or skipped); never verified
"""
from __future__ import annotations

import argparse
import json
import re
import time
from collections import Counter
from datetime import date
from pathlib import Path

import pythoncom
import win32com.client

ROOT = Path(r"c:\repos\microstation-vba-project")
DEFAULT_CANDIDATES = ROOT / "Data" / "keyin-candidates.tsv"
DEFAULT_REGISTRY = ROOT / "Data" / "command-registry.tsv"
DEFAULT_RESULTS = ROOT / "Bridge" / "keyin-probe-batch.json"

SAFE_KINDS = {"settings", "view", "lock"}
UNSAFE_KINDS = {"tool", "datapoint", "dialog", "file"}
# Never SendKeyin these — side effects too large (open file, reload refs, etc.)
SKIP_EXECUTE_KINDS = {"file", "dialog"}  # never SendKeyin — open UI / mutate files

# Scratch file only — never probe against a real project DGN.
REQUIRED_TEST_DGN = "DELETE.dgn"

# Key-ins that returned OK but have crashed MicroStation mid/after batch,
# or blocked SendKeyin indefinitely (PRINT LEVELS hung a live batch).
# Match against uppercased keyin text; skip live SendKeyin entirely.
SKIP_EXECUTE_KEYIN_RES = (
    re.compile(r"^MDL\s+(SILENT)?LOAD\b"),
    re.compile(r"^MDL\s+UNLOAD\b"),
    re.compile(r"^MDL\s+LIST\b"),
    re.compile(r"^EXPAND\s+KEYIN\b"),
    re.compile(r"\bDESIGN\s+HISTORY\b"),
    re.compile(r"^NAMED\s+VIEW\s+(ATTACH|DETACH)\s+ALL\b"),
    re.compile(r"^PRINT\s+LEVELS\b"),
    re.compile(r"^PLOT\s+LEVELS\b"),
    re.compile(r"^FACET\s+EDGES\b"),
    re.compile(r"^SET\s+FACET\s+EDGES\b"),
    re.compile(r"^HIDDEN\s+EDGES\b"),
    re.compile(r"^VISIBLE\s+EDGES\b"),
    re.compile(r"^PUBLISHDGN\s+PUBLISH\b"),
    re.compile(r"^PRINTORGANIZER\b"),
    re.compile(r"^IPLOT\b"),
    re.compile(r"^BATCHPROCESS\b"),
    re.compile(r"^PRINT\s+(VIEW|SCALE|PAPERSIZE|ORIENTATION|AREA|PENMAP|RASTER|REFERENCE|MONOCHROME|GRAYSCALE|TRUECOLOR|COLLATE|COPIES|DRIVER|DESTINATION|EXECUTE|SUBMIT)\b"),
    # OpenRoads/civil display often hangs without civil apps loaded
    re.compile(r"^CIVIL\s+DISPLAY\b"),
    re.compile(r"^VIEW\s+CIVIL\b"),
    re.compile(r"^VIEW\s+GEOMETRY\b"),
    re.compile(r"^GEOMETRY\s+DISPLAY\b"),
    re.compile(r"^SUPERELEVATION\b"),
    re.compile(r"^CORRIDOR\s+DISPLAY\b"),
    re.compile(r"^PROFILE\s+DISPLAY\b"),
    re.compile(r"^ALIGNMENT\s+DISPLAY\b"),
    re.compile(r"^FEATURE\s+DEFINITION\b"),
    re.compile(r"^CIVIL\s+LABEL"),
    re.compile(r"^OPENROADS\b"),
    re.compile(r"^CIVIL\s+GEOMETRY\b"),
)


def _should_skip_execute_keyin(keyin: str) -> str | None:
    """Return reason if this keyin must not be SendKeyin'd during probe."""
    ku = keyin.strip().upper()
    for rx in SKIP_EXECUTE_KEYIN_RES:
        if rx.search(ku):
            return f"crash-risk skip — matched {rx.pattern}"
    return None


def _ms_alive(app) -> bool:
    """Cheap heartbeat — if COM is dead after a keyin, stop the batch."""
    try:
        _ = app.ActiveDesignFile.Name
        return True
    except Exception:
        return False


def _read_tsv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    raw = path.read_text(encoding="utf-8", errors="replace")
    lines = []
    for line in raw.splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        lines.append(line)
    if not lines:
        return []
    header = lines[0].split("\t")
    rows = []
    for line in lines[1:]:
        cols = line.split("\t")
        row = {header[i]: (cols[i] if i < len(cols) else "") for i in range(len(header))}
        rows.append(row)
    return rows


def _existing_op_names(registry: Path) -> set[str]:
    names: set[str] = set()
    for row in _read_tsv_rows(registry):
        if row.get("opName"):
            names.add(row["opName"])
    return names


def _existing_keyins(registry: Path) -> set[str]:
    """Normalize recipe KEYIN:text for duplicate detection."""
    found: set[str] = set()
    for row in _read_tsv_rows(registry):
        recipe = row.get("recipeLines", "")
        for step in recipe.split("|"):
            step = step.strip()
            if step.upper().startswith("KEYIN:"):
                found.add(step[6:].strip().upper())
            elif step.upper().startswith("COMMAND:"):
                found.add(step[8:].strip().upper())
    return found


def probe(candidates_path: Path, results_path: Path, skip_existing: bool) -> list[dict]:
    candidates = _read_tsv_rows(candidates_path)
    if not candidates:
        raise SystemExit(f"no candidates in {candidates_path}")

    registry = DEFAULT_REGISTRY
    known_ops = _existing_op_names(registry)
    known_keyins = _existing_keyins(registry)

    pythoncom.CoInitialize()
    app = win32com.client.GetObject(Class="MicroStationDGN.Application")
    try:
        active_name = app.ActiveDesignFile.Name
    except Exception as e:
        raise SystemExit(
            f"No design file active in MicroStation ({e}). "
            f"Open {REQUIRED_TEST_DGN} and retry."
        )
    if active_name.upper() != REQUIRED_TEST_DGN.upper():
        raise SystemExit(
            f"Refusing to probe: active file is '{active_name}', "
            f"required scratch file is '{REQUIRED_TEST_DGN}'. "
            f"Switch to DELETE.dgn (e.g. c:\\pwworking\\usny\\d0119562\\DELETE.dgn) and retry."
        )
    print(f"probing against {app.ActiveDesignFile.FullName}")
    q = app.CadInputQueue

    results: list[dict] = []
    for row in candidates:
        op = (row.get("opName") or "").strip()
        keyin = (row.get("keyin") or "").strip()
        kind = (row.get("kind") or "settings").strip().lower()
        if not op or not keyin:
            continue
        if skip_existing and (op in known_ops or keyin.upper() in known_keyins):
            results.append({
                "opName": op, "keyin": keyin, "kind": kind,
                "requiredParams": row.get("requiredParams", ""),
                "source": row.get("source", ""), "notes": row.get("notes", ""),
                "verdict": "SKIP_EXISTS", "dt": 0.0, "err": "",
            })
            print(f"SKIP    exists  {op}")
            continue

        if kind in SKIP_EXECUTE_KINDS:
            results.append({
                "opName": op, "keyin": keyin, "kind": kind,
                "requiredParams": row.get("requiredParams", ""),
                "source": row.get("source", ""), "notes": row.get("notes", ""),
                "verdict": "UNSAFE", "dt": 0.0,
                "err": f"not executed — kind={kind}",
            })
            print(f"UNSAFE  skip-exec [{kind:9}] {op}")
            continue

        skip_reason = _should_skip_execute_keyin(keyin)
        if skip_reason:
            results.append({
                "opName": op, "keyin": keyin, "kind": kind,
                "requiredParams": row.get("requiredParams", ""),
                "source": row.get("source", ""), "notes": row.get("notes", ""),
                "verdict": "UNSAFE", "dt": 0.0, "err": skip_reason,
            })
            print(f"UNSAFE  skip-exec [crashrisk] {op}  |  {keyin}")
            continue

        t0 = time.perf_counter()
        err = ""
        try:
            # Must call SendKeyin on this thread — COM is STA; a worker-thread
            # timeout wrapper returns instant bogus errors and can poison the
            # apartment. Hang risk is handled by skip-exec for dialog/file and
            # by keeping batches free of known-blocking keyins.
            q.SendKeyin(keyin)
            try:
                q.SendReset()
            except Exception:
                pass
            try:
                app.CommandState.StartDefaultCommand()
            except Exception as e2:
                err = f"neutralize:{e2}"
        except Exception as e:
            err = str(e)
        dt = time.perf_counter() - t0
        hard_err = bool(err) and not err.startswith("neutralize")

        if kind in UNSAFE_KINDS:
            verdict = "UNSAFE"
        elif hard_err:
            verdict = "ERR"
        elif dt > 8:
            verdict = "SLOW"
        else:
            verdict = "OK"

        rec = {
            "opName": op, "keyin": keyin, "kind": kind,
            "requiredParams": row.get("requiredParams", ""),
            "source": row.get("source", ""), "notes": row.get("notes", ""),
            "verdict": verdict, "dt": round(dt, 3), "err": err[:120],
        }
        results.append(rec)
        print(f"{verdict:8} {dt:6.3f}s  [{kind:9}] {op}  |  {keyin}")

        if not _ms_alive(app):
            print(f"FATAL   MicroStation COM died after {op} — stopping batch")
            rec["err"] = (rec.get("err") or "") + " | COM dead after this keyin"
            results_path.parent.mkdir(parents=True, exist_ok=True)
            results_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
            raise SystemExit(
                f"MicroStation crashed or disconnected after probing {op} ({keyin}). "
                f"Partial results written to {results_path}."
            )

    results_path.parent.mkdir(parents=True, exist_ok=True)
    results_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
    counts = Counter(r["verdict"] for r in results)
    print("---")
    print(dict(counts), "wrote", results_path)
    return results


def _op_name_to_recipe(keyin: str, required_params: str) -> str:
    """Build KEYIN: recipe; if requiredParams listed, replace trailing value tokens with {param}."""
    params = [p for p in required_params.split("|") if p.strip()]
    if not params:
        return f"KEYIN:{keyin}"
    # Heuristic: last N whitespace-separated tokens become placeholders when
    # candidate keyin was written with sample values. Prefer explicit {param}
    # already in the keyin string.
    if "{" in keyin:
        return f"KEYIN:{keyin}"
    parts = keyin.split()
    if len(parts) <= len(params):
        # e.g. AA=0 style
        if "=" in keyin and len(params) == 1:
            prefix = keyin.split("=", 1)[0]
            return f"KEYIN:{prefix}={{{params[0]}}}"
        return f"KEYIN:{keyin}"
    head = parts[: len(parts) - len(params)]
    tail = [f"{{{p}}}" for p in params]
    return "KEYIN:" + " ".join(head + tail)


def promote(results_path: Path, registry_path: Path, dry_run: bool = False) -> int:
    results = json.loads(results_path.read_text(encoding="utf-8"))
    known = _existing_op_names(registry_path)
    today = date.today().isoformat()
    rows: list[str] = []

    for r in results:
        op = r["opName"]
        if op in known:
            continue
        verdict = r["verdict"]
        kind = r["kind"]
        keyin = r["keyin"]
        req = r.get("requiredParams", "")
        src = r.get("source", "") or f"batch probe {today}"
        notes = r.get("notes", "")
        dt = r.get("dt", 0)

        if verdict == "OK" and kind in SAFE_KINDS:
            status = "verified-headless-safe"
            promoted = today
            note = (notes + " " if notes else "") + f"live batch probe {today} ({dt}s)"
        elif verdict in ("UNSAFE", "OK") and kind in UNSAFE_KINDS:
            # OK+unsafe kind still blocked; UNSAFE same
            status = "unsafe-blocked"
            promoted = ""
            note = (notes + " " if notes else "") + f"batch probe {today}: kind={kind}, not promotable"
        elif verdict == "SLOW" and kind in SAFE_KINDS:
            status = "needs-testing"
            promoted = ""
            note = (notes + " " if notes else "") + f"slow return ({dt}s) — manual review"
        elif verdict == "ERR":
            status = "needs-testing"
            promoted = ""
            note = (notes + " " if notes else "") + f"probe error: {r.get('err','')}"
        else:
            continue  # SKIP_EXISTS etc.

        recipe = _op_name_to_recipe(keyin, req)
        # COMMAND: for bare tool names that look like commands without KEYIN settings form
        if kind == "tool" and not keyin.upper().startswith("DIALOG"):
            recipe = f"COMMAND:{keyin}"

        row = "\t".join([
            op, "keyin_recipe", status, recipe, "", req, "", "N", "N",
            src, today, promoted, note.strip(),
        ])
        rows.append(row)
        known.add(op)
        print(f"PROMOTE {status:28} {op}")

    if dry_run:
        print(f"dry-run: would add {len(rows)} rows")
        return len(rows)

    if not rows:
        print("nothing to promote")
        return 0

    text = registry_path.read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n") + "\n"
    text = text + "\n".join(rows) + "\n"
    registry_path.write_bytes(text.replace("\n", "\r\n").encode("utf-8"))
    print(f"appended {len(rows)} rows -> {registry_path}")
    return len(rows)


def main() -> None:
    ap = argparse.ArgumentParser(description="Batch probe / promote MicroStation key-ins")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_probe = sub.add_parser("probe", help="Serially probe candidates against live MicroStation")
    p_probe.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    p_probe.add_argument("--out", type=Path, default=DEFAULT_RESULTS)
    p_probe.add_argument("--skip-existing", action="store_true", default=True)
    p_probe.add_argument("--no-skip-existing", action="store_false", dest="skip_existing")

    p_promote = sub.add_parser("promote", help="Append probe results into command-registry.tsv")
    p_promote.add_argument("--results", type=Path, default=DEFAULT_RESULTS)
    p_promote.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_promote.add_argument("--dry-run", action="store_true")

    p_run = sub.add_parser("run", help="probe then promote in one shot")
    p_run.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    p_run.add_argument("--out", type=Path, default=DEFAULT_RESULTS)
    p_run.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_run.add_argument("--dry-run", action="store_true")
    p_run.add_argument("--skip-existing", action="store_true", default=True)
    p_run.add_argument("--no-skip-existing", action="store_false", dest="skip_existing")

    args = ap.parse_args()
    if args.cmd == "probe":
        probe(args.candidates, args.out, args.skip_existing)
    elif args.cmd == "promote":
        promote(args.results, args.registry, args.dry_run)
    elif args.cmd == "run":
        probe(args.candidates, args.out, args.skip_existing)
        promote(args.out, args.registry, args.dry_run)


if __name__ == "__main__":
    main()
