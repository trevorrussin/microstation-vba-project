"""
Batch probe + promote for drawing recipes (COMMAND + DATAPOINT + RESET).

Unlike keyin_batch.py (settings/view/lock only), this script:
  - Runs full recipeLines sequences matching WZTCCommandRegistry.RunRecipeStep
  - Passes only when graphical element count increases by expectedNewElements
  - Uses DELETE.dgn only; 3s hard timeout per recipe step

Workflow
--------
1. Edit Data/recipe-candidates.tsv
2. python scripts/recipe_batch.py run
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import date
from pathlib import Path

import pythoncom
import win32com.client

ROOT = Path(r"c:\repos\microstation-vba-project")
DEFAULT_CANDIDATES = ROOT / "Data" / "recipe-candidates.tsv"
DEFAULT_REGISTRY = ROOT / "Data" / "command-registry.tsv"
DEFAULT_RESULTS = ROOT / "Bridge" / "recipe-probe-batch.json"

REQUIRED_TEST_DGN = "DELETE.dgn"
STEP_TIMEOUT_SEC = 3.0


def _read_tsv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    raw = path.read_text(encoding="utf-8", errors="replace")
    lines = []
    for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        lines.append(line.rstrip("\r"))
    if not lines:
        return []
    header = [h.strip() for h in lines[0].split("\t")]
    rows = []
    for line in lines[1:]:
        cols = [c.strip("\r") for c in line.split("\t")]
        row = {header[i]: (cols[i].strip() if i < len(cols) else "") for i in range(len(header))}
        rows.append(row)
    return rows


def _parse_sample_params(sample: str) -> dict[str, str]:
    """x1=0|y1=0|x2=10|y2=0 -> dict."""
    out: dict[str, str] = {}
    for part in sample.split("|"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        k, v = part.split("=", 1)
        out[k.strip()] = v.strip()
    return out


def _substitute_params(recipe: str, params: dict[str, str]) -> str:
    result = recipe
    for k, v in params.items():
        result = result.replace("{" + k + "}", v)
    return result


def _check_close_out_guard(recipe: str) -> str | None:
    steps = [s.strip() for s in recipe.split("|") if s.strip()]
    has_command = any(s.upper().startswith("COMMAND:") for s in steps)
    has_dp = any(s.upper().startswith("DATAPOINT:") for s in steps)
    has_reset = any(s.upper() == "RESET" for s in steps)
    if has_command and not (has_dp and has_reset):
        return "COMMAND: without both DATAPOINT: and RESET"
    return None


def _ms_alive(app) -> bool:
    try:
        _ = app.ActiveDesignFile.Name
        return True
    except Exception:
        return False


def _graphical_element_count(app) -> int:
    """Count non-graphical-excluded elements via ElementScanCriteria (same as VBA)."""
    model = app.ActiveModelReference
    try:
        sc = win32com.client.Dispatch("MicroStationDGN.ElementScanCriteria")
    except Exception:
        # Fallback typelib progid variants
        sc = win32com.client.Dispatch("Bentley.MicroStation.ElementScanCriteria")
    sc.ExcludeNonGraphical()
    ee = model.Scan(sc)
    n = 0
    while ee.MoveNext():
        n += 1
    return n


def _one_step(step: str) -> int:
    """Child: run one recipe step on DELETE.dgn. Exit 0=ok, 2=wrong file, 1=err."""
    pythoncom.CoInitialize()
    app = win32com.client.GetObject(Class="MicroStationDGN.Application")
    try:
        active_name = app.ActiveDesignFile.Name
    except Exception as e:
        print(f"no design file: {e}", flush=True)
        return 2
    if active_name.upper() != REQUIRED_TEST_DGN.upper():
        print(f"wrong file: {active_name}", flush=True)
        return 2

    q = app.CadInputQueue
    upper = step.upper()
    try:
        if upper.startswith("KEYIN:"):
            q.SendKeyin(step[6:])
        elif upper.startswith("COMMAND:"):
            q.SendCommand(step[8:])
        elif upper.startswith("SETCELL:"):
            # Match VBA SetCExpressionValue path via active-cell keyin
            q.SendKeyin("AC=" + step[8:])
        elif upper.startswith("DATAPOINT:"):
            coords = step[10:]
            parts = [p.strip() for p in coords.split(",")]
            if len(parts) < 2:
                print("DATAPOINT needs x,y", flush=True)
                return 1
            pt = win32com.client.Record("MicroStationDGN.Point3d")
            # Point3d may need manual construction
            try:
                pt.X = float(parts[0])
                pt.Y = float(parts[1])
                pt.Z = float(parts[2]) if len(parts) >= 3 else 0.0
            except Exception:
                # Some installs expose Point3d as a plain tuple via helper
                pt = app.Point3dFromXYZ(
                    float(parts[0]),
                    float(parts[1]),
                    float(parts[2]) if len(parts) >= 3 else 0.0,
                )
            q.SendDataPoint(pt, 1)
        elif upper == "RESET":
            q.SendReset()
        elif upper == "DEFAULTCOMMAND":
            app.CommandState.StartDefaultCommand()
        else:
            print(f"unknown step: {step}", flush=True)
            return 1
    except Exception as e:
        print(f"err:{e}", flush=True)
        return 1
    return 0


def _run_step_subprocess(step: str, timeout: float = STEP_TIMEOUT_SEC) -> tuple[str, float, str]:
    t0 = time.perf_counter()
    try:
        proc = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), "_one_step", step],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        dt = time.perf_counter() - t0
        err = (proc.stderr or proc.stdout or "").strip()[:120]
        if proc.returncode == 0:
            return "OK", dt, err
        if proc.returncode == 2:
            return "WRONG_FILE", dt, err or "not DELETE.dgn"
        return "ERR", dt, err or f"exit {proc.returncode}"
    except subprocess.TimeoutExpired:
        dt = time.perf_counter() - t0
        return "HANG", dt, f"step exceeded {timeout}s"


def _make_point3d(app, x: float, y: float, z: float = 0.0):
    """Build a Point3d compatible with SendDataPoint on this install."""
    try:
        return app.Point3dFromXYZ(x, y, z)
    except Exception:
        pass
    try:
        pt = win32com.client.Record("MicroStationDGN.Point3d")
        pt.X = x
        pt.Y = y
        pt.Z = z
        return pt
    except Exception:
        pass
    # Last resort: dynamic dispatch object with X/Y/Z
    class _P:
        pass

    p = _P()
    p.X = x
    p.Y = y
    p.Z = z
    return p


def _run_step_inline(app, step: str) -> None:
    """Same-thread step (used when subprocess Point3d is awkward). Prefer for DATAPOINT."""
    q = app.CadInputQueue
    upper = step.upper()
    if upper.startswith("KEYIN:"):
        q.SendKeyin(step[6:])
    elif upper.startswith("COMMAND:"):
        q.SendCommand(step[8:])
    elif upper.startswith("SETCELL:"):
        q.SendKeyin("AC=" + step[8:])
    elif upper.startswith("DATAPOINT:"):
        parts = [p.strip() for p in step[10:].split(",")]
        if len(parts) < 2:
            raise RuntimeError("DATAPOINT needs x,y")
        pt = _make_point3d(
            app,
            float(parts[0]),
            float(parts[1]),
            float(parts[2]) if len(parts) >= 3 else 0.0,
        )
        q.SendDataPoint(pt, 1)
    elif upper == "RESET":
        q.SendReset()
    elif upper == "DEFAULTCOMMAND":
        app.CommandState.StartDefaultCommand()
    else:
        raise RuntimeError(f"unknown step: {step}")


def probe(candidates_path: Path, results_path: Path) -> list[dict]:
    candidates = _read_tsv_rows(candidates_path)
    if not candidates:
        raise SystemExit(f"no candidates in {candidates_path}")

    pythoncom.CoInitialize()
    app = win32com.client.GetObject(Class="MicroStationDGN.Application")
    try:
        active_name = app.ActiveDesignFile.Name
    except Exception as e:
        raise SystemExit(f"No design file active ({e}). Open {REQUIRED_TEST_DGN}.")
    if active_name.upper() != REQUIRED_TEST_DGN.upper():
        raise SystemExit(
            f"Refusing to probe: active file is '{active_name}', "
            f"required '{REQUIRED_TEST_DGN}'."
        )
    print(
        f"recipe probe against {app.ActiveDesignFile.FullName} "
        f"(step timeout={STEP_TIMEOUT_SEC}s)",
        flush=True,
    )

    results: list[dict] = []
    for row in candidates:
        op = (row.get("opName") or "").strip()
        recipe = (row.get("recipeLines") or "").strip()
        sample = (row.get("sampleParams") or "").strip()
        source = (row.get("source") or "").strip()
        notes = (row.get("notes") or "").strip()
        req = (row.get("requiredParams") or "").strip()
        try:
            expected = int((row.get("expectedNewElements") or "1").strip() or "1")
        except ValueError:
            expected = 1
        if not op or not recipe:
            continue

        guard = _check_close_out_guard(recipe)
        if guard:
            rec = {
                "opName": op, "recipeLines": recipe, "requiredParams": req,
                "sampleParams": sample, "expectedNewElements": expected,
                "source": source, "notes": notes,
                "verdict": "GUARD", "before": 0, "after": 0, "dt": 0.0,
                "err": guard,
            }
            results.append(rec)
            print(f"GUARD   {op}  |  {guard}", flush=True)
            continue

        params = _parse_sample_params(sample)
        expanded = _substitute_params(recipe, params)
        if "{" in expanded and "}" in expanded:
            leftover = re.findall(r"\{[^}]+\}", expanded)
            rec = {
                "opName": op, "recipeLines": recipe, "requiredParams": req,
                "sampleParams": sample, "expectedNewElements": expected,
                "source": source, "notes": notes,
                "verdict": "ERR", "before": 0, "after": 0, "dt": 0.0,
                "err": f"unsubstituted placeholders: {leftover}",
            }
            results.append(rec)
            print(f"ERR     {op}  unsubstituted {leftover}", flush=True)
            continue

        try:
            before = _graphical_element_count(app)
        except Exception as e:
            raise SystemExit(f"element count failed before {op}: {e}")

        t0 = time.perf_counter()
        step_err = ""
        hang = False
        steps = [s.strip() for s in expanded.split("|") if s.strip()]
        for step in steps:
            upper = step.upper()
            # COMMAND/KEYIN can hang — use subprocess timeout.
            # DATAPOINT/RESET/DEFAULTCOMMAND stay in-process (Point3d COM).
            if upper.startswith("COMMAND:") or upper.startswith("KEYIN:") or upper.startswith("SETCELL:"):
                hint, _dt, err = _run_step_subprocess(step, STEP_TIMEOUT_SEC)
                if hint == "WRONG_FILE":
                    results_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
                    raise SystemExit(f"Active file left {REQUIRED_TEST_DGN}: {err}")
                if hint == "HANG":
                    hang = True
                    step_err = err
                    break
                if hint == "ERR":
                    step_err = err
                    break
                # Rebind app after subprocess
                try:
                    app = win32com.client.GetObject(Class="MicroStationDGN.Application")
                except Exception as e:
                    step_err = f"COM rebind failed: {e}"
                    break
            else:
                try:
                    _run_step_inline(app, step)
                except Exception as e:
                    step_err = str(e)
                    break

        # Always try to neutralize after recipe
        try:
            app.CadInputQueue.SendReset()
        except Exception:
            pass
        try:
            app.CommandState.StartDefaultCommand()
        except Exception:
            pass

        dt = time.perf_counter() - t0
        alive = _ms_alive(app)
        after = before
        if alive:
            try:
                after = _graphical_element_count(app)
            except Exception as e:
                step_err = (step_err + f" | count-after: {e}").strip(" |")

        delta = after - before
        if hang:
            verdict = "HANG"
        elif not alive:
            verdict = "FATAL"
        elif step_err:
            verdict = "ERR"
        elif delta >= expected:
            verdict = "PASS"
        else:
            verdict = "NO_ELEMENT"
            if not step_err:
                step_err = f"expected +{expected} elements, got delta={delta} ({before}->{after})"

        rec = {
            "opName": op, "recipeLines": recipe, "requiredParams": req,
            "sampleParams": sample, "expectedNewElements": expected,
            "source": source, "notes": notes,
            "verdict": verdict, "before": before, "after": after,
            "delta": delta, "dt": round(dt, 3), "err": step_err[:200],
        }
        results.append(rec)
        print(
            f"{verdict:10} {dt:6.3f}s  +{delta} ({before}->{after})  {op}",
            flush=True,
        )
        if verdict == "FATAL":
            results_path.parent.mkdir(parents=True, exist_ok=True)
            results_path.write_text(json.dumps(results, indent=2), encoding="utf-8")
            raise SystemExit(f"MicroStation died after {op}")

        results_path.parent.mkdir(parents=True, exist_ok=True)
        results_path.write_text(json.dumps(results, indent=2), encoding="utf-8")

    counts = Counter(r["verdict"] for r in results)
    print("---", flush=True)
    print(dict(counts), "wrote", results_path, flush=True)
    return results


def _registry_rows(path: Path) -> tuple[str, list[str], list[list[str]]]:
    text = path.read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [l for l in text.split("\n") if l.strip()]
    header = lines[0].split("\t")
    rows = [l.split("\t") for l in lines[1:]]
    return lines[0], header, rows


def promote(results_path: Path, registry_path: Path, dry_run: bool = False) -> int:
    results = json.loads(results_path.read_text(encoding="utf-8"))
    _, header, rows = _registry_rows(registry_path)
    i_op = header.index("opName")
    i_safety = header.index("safetyStatus")
    i_recipe = header.index("recipeLines")
    i_req = header.index("requiredParams")
    i_creates = header.index("createsElements")
    i_own = header.index("ownElementOnly")
    i_src = header.index("sourceRefs")
    i_added = header.index("addedDate")
    i_promoted = header.index("promotedDate")
    i_notes = header.index("notes")
    i_cat = header.index("category")

    today = date.today().isoformat()
    by_op = {r[i_op]: (idx, r) for idx, r in enumerate(rows)}
    changed = 0
    appended: list[list[str]] = []

    for r in results:
        op = r["opName"]
        verdict = r["verdict"]
        recipe = r["recipeLines"]
        req = r.get("requiredParams", "")
        src = r.get("source", "") or f"recipe probe {today}"
        notes = r.get("notes", "")
        dt = r.get("dt", 0)
        delta = r.get("delta", 0)

        if verdict == "PASS":
            status = "verified-headless-safe"
            promoted = today
            note = (
                f"{notes} ".strip()
                + f" recipe probe {today}: PASS delta=+{delta} ({dt}s)"
            ).strip()
        elif verdict == "HANG":
            status = "unsafe-blocked"
            promoted = ""
            note = (f"{notes} ".strip() + f" recipe probe {today}: HANG — {r.get('err','')}").strip()
        elif verdict in ("NO_ELEMENT", "ERR", "GUARD"):
            status = "needs-testing"
            promoted = ""
            note = (
                f"{notes} ".strip()
                + f" recipe probe {today}: {verdict} — {r.get('err','')}"
            ).strip()
        else:
            continue

        if op in by_op:
            idx, row = by_op[op]
            # Pad row to header length
            while len(row) < len(header):
                row.append("")
            row[i_safety] = status
            row[i_recipe] = recipe
            if req:
                row[i_req] = req
            row[i_creates] = "Y"
            row[i_own] = row[i_own] or "N"
            if src and src not in row[i_src]:
                row[i_src] = (row[i_src] + ";" + src).strip(";") if row[i_src] else src
            if promoted:
                row[i_promoted] = promoted
            row[i_notes] = note
            rows[idx] = row
            changed += 1
            print(f"UPDATE {status:28} {op}", flush=True)
        else:
            new = [""] * len(header)
            new[i_op] = op
            new[i_cat] = "keyin_recipe"
            new[i_safety] = status
            new[i_recipe] = recipe
            new[i_req] = req
            new[i_creates] = "Y"
            new[i_own] = "N"
            new[i_src] = src
            new[i_added] = today
            new[i_promoted] = promoted
            new[i_notes] = note
            appended.append(new)
            changed += 1
            print(f"APPEND {status:28} {op}", flush=True)

    if dry_run:
        print(f"dry-run: would change {changed} rows", flush=True)
        return changed

    if changed == 0:
        print("nothing to promote", flush=True)
        return 0

    all_rows = rows + appended
    out_lines = ["\t".join(header)] + ["\t".join(r) for r in all_rows]
    registry_path.write_bytes(("\r\n".join(out_lines) + "\r\n").encode("utf-8"))
    print(f"wrote {changed} updates -> {registry_path}", flush=True)
    return changed


def main() -> None:
    if len(sys.argv) >= 2 and sys.argv[1] == "_one_step":
        step = " ".join(sys.argv[2:]).strip()
        if not step:
            raise SystemExit(1)
        raise SystemExit(_one_step(step))

    ap = argparse.ArgumentParser(description="Probe/promote drawing recipes")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_probe = sub.add_parser("probe")
    p_probe.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    p_probe.add_argument("--out", type=Path, default=DEFAULT_RESULTS)

    p_promote = sub.add_parser("promote")
    p_promote.add_argument("--results", type=Path, default=DEFAULT_RESULTS)
    p_promote.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_promote.add_argument("--dry-run", action="store_true")

    p_run = sub.add_parser("run")
    p_run.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    p_run.add_argument("--out", type=Path, default=DEFAULT_RESULTS)
    p_run.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_run.add_argument("--dry-run", action="store_true")

    args = ap.parse_args()
    if args.cmd == "probe":
        probe(args.candidates, args.out)
    elif args.cmd == "promote":
        promote(args.results, args.registry, args.dry_run)
    elif args.cmd == "run":
        probe(args.candidates, args.out)
        promote(args.out, args.registry, args.dry_run)


if __name__ == "__main__":
    main()
