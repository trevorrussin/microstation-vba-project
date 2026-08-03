"""
Audit verified-headless-safe KEYINs for the ZOOM_* false-OK failure mode.

Root cause (2026-08-02): keyin_batch._one_keyin SendKeyin'd then immediately
SendReset()'d, so a key-in that arms a "select point/view" tool returned OK
in ~0.08s and got promoted — the Reset cancelled the pending prompt before
the probe could see it.

This audit does NOT Reset until after inspection:
  1. StartDefaultCommand (baseline CommandName, usually 'Element Selection')
  2. SendKeyin(recipe)
  3. Read CommandState.CommandName + MessageCenter.StatusPrompt
  4. Then SendReset + StartDefaultCommand

Verdicts
--------
CLEAN  — CommandName still equals baseline after the key-in (no tool left armed)
ARMED  — CommandName changed (tool waiting for view/point/pick) → false-OK risk
SKIP   — placeholder recipe, crash-risk skip list, or not a plain KEYIN:
ERR    — SendKeyin / COM error
HANG   — reserved (not used in-process; keep Reset path short)

Apply mode (--apply) updates Data/command-registry.tsv:
  ARMED  + was verified-headless-safe → unsafe-blocked
  CLEAN  + was needs-testing (prior precautionary downgrade) → verified-headless-safe
           only when --restore-clean is also passed

Scratch file only: refuses to run unless ActiveDesignFile is DELETE.dgn.
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

ROOT = Path(r"c:\repos\microstation-vba-project")
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "mcp-server"))
import keyin_batch  # noqa: E402 -- SKIP_EXECUTE_KEYIN_RES / REQUIRED_TEST_DGN
import ms_connect  # noqa: E402

DEFAULT_REGISTRY = ROOT / "Data" / "command-registry.tsv"
DEFAULT_RESULTS = ROOT / "Bridge" / "keyin-false-ok-audit.json"

# Families most likely to leave a pick/view prompt. Used by --scope suspect.
SUSPECT_OP_RE = re.compile(
    r"(ZOOM|WINDOW_|FIT_|PAN_|MOVE_VIEW|ROTATE_VIEW|UPDATE_VIEW|"
    r"VIEW_TOP|VIEW_BOTTOM|VIEW_FRONT|VIEW_BACK|VIEW_LEFT|VIEW_RIGHT|"
    r"VIEW_ISO|VIEW_NORTH|VIEW_CAMERA|ACCUDRAW_ROTATE|NAMEDVIEW|SAVEDVIEW|"
    r"SELVIEW|WINDOW_)",
    re.I,
)

AUDIT_NOTE_TAG = "false-OK audit 2026-08-02"


def _read_tsv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    raw = path.read_text(encoding="utf-8", errors="replace")
    lines = [ln for ln in raw.splitlines() if ln.strip() and not ln.strip().startswith("#")]
    header = [h.strip().strip("\r") for h in lines[0].split("\t")]
    rows: list[dict[str, str]] = []
    for line in lines[1:]:
        cols = [c.strip("\r") for c in line.split("\t")]
        row = {header[i]: (cols[i] if i < len(cols) else "") for i in range(len(header))}
        rows.append(row)
    return header, rows


def _keyin_from_recipe(recipe: str) -> str | None:
    recipe = (recipe or "").strip()
    if not recipe.upper().startswith("KEYIN:"):
        return None
    if "|" in recipe:
        return None
    body = recipe[6:].strip()
    if not body or "{" in body:
        return None
    return body


def _select_rows(rows: list[dict[str, str]], scope: str) -> list[dict[str, str]]:
    out = []
    for row in rows:
        status = row.get("safetyStatus", "")
        op = row.get("opName", "")
        keyin = _keyin_from_recipe(row.get("recipeLines", ""))
        if keyin is None:
            continue
        if scope == "verified":
            if status != "verified-headless-safe":
                continue
        elif scope == "needs-testing":
            if status != "needs-testing":
                continue
        elif scope == "suspect":
            if status not in ("verified-headless-safe", "needs-testing"):
                continue
            if not SUSPECT_OP_RE.search(op) and not SUSPECT_OP_RE.search(keyin):
                continue
        elif scope == "all-gated":
            if status not in ("verified-headless-safe", "needs-testing"):
                continue
        else:
            raise SystemExit(f"unknown scope {scope!r}")
        out.append(row)
    return out


def _reset(app) -> None:
    try:
        app.CadInputQueue.SendReset()
    except Exception:
        pass
    try:
        app.CommandState.StartDefaultCommand()
    except Exception:
        pass
    time.sleep(0.05)


def _snapshot(app) -> dict[str, str]:
    cs = app.CommandState
    mc = app.MessageCenter
    return {
        "commandName": str(cs.CommandName or ""),
        "statusCommand": str(mc.StatusCommand or ""),
        "statusPrompt": str(mc.StatusPrompt or ""),
        "statusMessage": str(mc.StatusMessage or ""),
        "statusWarning": str(mc.StatusWarning or ""),
    }


SENDKEYIN_TIMEOUT_SEC = 3.0


def _one_keyin_no_reset(keyin: str) -> int:
    """Child entry: SendKeyin only (no Reset) so parent can inspect CommandName.
    Exit 0=ok, 2=wrong file, 1=err."""
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app()
    try:
        active_name = app.ActiveDesignFile.Name
    except Exception as e:
        print(f"no design file: {e}", flush=True)
        return 2
    if active_name.upper() != keyin_batch.REQUIRED_TEST_DGN.upper():
        print(f"wrong file: {active_name}", flush=True)
        return 2
    try:
        app.CadInputQueue.SendKeyin(keyin)
    except Exception as e:
        print(f"err:{e}", flush=True)
        return 1
    return 0


def _sendkeyin_no_reset(keyin: str, timeout: float = SENDKEYIN_TIMEOUT_SEC) -> tuple[str, float, str]:
    t0 = time.perf_counter()
    try:
        proc = subprocess.run(
            [sys.executable, str(Path(__file__).resolve()), "_one_keyin_no_reset", keyin],
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
        return "HANG", dt, f"SendKeyin exceeded {timeout}s"


def _classify(baseline_cmd: str, after: dict[str, str]) -> tuple[str, str]:
    """Return (verdict, note). Empty CommandName after NULL/NOCOMMAND is CLEAN
    (command cleared), not ARMED."""
    cmd = (after.get("commandName") or "").strip()
    if not cmd or cmd == baseline_cmd:
        return "CLEAN", ""
    prompt = after.get("statusPrompt") or ""
    return "ARMED", (
        f"left CommandName={cmd!r} prompt={prompt!r} "
        f"(baseline={baseline_cmd!r})"
    )


def audit(registry: Path, results_path: Path, scope: str, limit: int = 0) -> list[dict]:
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app()
    try:
        active = app.ActiveDesignFile.Name
    except Exception as e:
        raise SystemExit(f"no design file: {e}")
    if active.upper() != keyin_batch.REQUIRED_TEST_DGN.upper():
        raise SystemExit(
            f"Refusing audit: active file is {active!r}, need "
            f"{keyin_batch.REQUIRED_TEST_DGN}"
        )

    _, rows = _read_tsv(registry)
    targets = _select_rows(rows, scope)
    if limit > 0:
        targets = targets[:limit]

    # Resume from checkpoint if present.
    results: list[dict] = []
    done_ops: set[str] = set()
    baseline: dict[str, str] | None = None
    if results_path.exists():
        try:
            prior = json.loads(results_path.read_text(encoding="utf-8"))
            if prior.get("scope") == scope and isinstance(prior.get("results"), list):
                results = prior["results"]
                done_ops = {r["opName"] for r in results if r.get("opName")}
                baseline = prior.get("baseline")
                # Reclassify empty-CommandName ARMED → CLEAN (NULL/NOCOMMAND).
                for r in results:
                    if r.get("verdict") == "ARMED" and r.get("after"):
                        v, note = _classify(
                            (baseline or {}).get("commandName", "Element Selection"),
                            r["after"],
                        )
                        r["verdict"] = v
                        if note:
                            r["note"] = note
                        elif "note" in r:
                            del r["note"]
                print(f"resuming with {len(done_ops)} prior results", flush=True)
        except Exception:
            results = []
            done_ops = set()
            baseline = None

    print(
        f"false-OK audit against {app.ActiveDesignFile.FullName} "
        f"scope={scope} n={len(targets)} done={len(done_ops)}",
        flush=True,
    )

    _reset(app)
    if not baseline:
        baseline = _snapshot(app)
    baseline_cmd = baseline["commandName"]
    print(f"baseline CommandName={baseline_cmd!r}", flush=True)

    for i, row in enumerate(targets, 1):
        op = row["opName"]
        if op in done_ops:
            continue
        keyin = _keyin_from_recipe(row["recipeLines"]) or ""
        prior_status = row.get("safetyStatus", "")
        rec: dict = {
            "opName": op,
            "keyin": keyin,
            "priorStatus": prior_status,
            "i": i,
        }

        skip = keyin_batch._should_skip_execute_keyin(keyin)
        if skip:
            rec["verdict"] = "SKIP"
            rec["err"] = skip
            results.append(rec)
            done_ops.add(op)
            print(f"SKIP  {op}: {skip}", flush=True)
            continue

        _reset(app)
        hint, dt, err = _sendkeyin_no_reset(keyin)
        rec["dt"] = round(dt, 3)
        if hint == "WRONG_FILE":
            raise SystemExit(f"wrong file during {op}: {err}")
        if hint == "HANG":
            rec["verdict"] = "HANG"
            rec["err"] = err
            print(f"HANG  {op}: {err}", flush=True)
            _reset(app)
            # Rebind after hang.
            try:
                pythoncom.CoInitialize()
                app = ms_connect.get_microstation_app()
                _reset(app)
            except Exception as e:
                print(f"WARN  post-hang rebind failed: {e}", flush=True)
        elif hint == "ERR":
            rec["verdict"] = "ERR"
            rec["err"] = err
            print(f"ERR   {op}: {err}", flush=True)
            _reset(app)
        else:
            time.sleep(0.08)
            after = _snapshot(app)
            rec["after"] = after
            verdict, note = _classify(baseline_cmd, after)
            rec["verdict"] = verdict
            if note:
                rec["note"] = note
            print(f"{verdict:5} {op}" + (f": {note}" if note else ""), flush=True)
            _reset(app)

        results.append(rec)
        done_ops.add(op)

        if len(results) % 25 == 0 or len(done_ops) >= len(targets):
            results_path.parent.mkdir(parents=True, exist_ok=True)
            results_path.write_text(json.dumps({
                "baseline": baseline,
                "scope": scope,
                "auditedAt": date.today().isoformat(),
                "results": results,
            }, indent=2), encoding="utf-8")
            print(f"... checkpoint {len(results)}/{len(targets)}", flush=True)

        if not keyin_batch._ms_alive(app):
            raise SystemExit(f"COM died after {op}; partial results in {results_path}")

    results_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "baseline": baseline,
        "scope": scope,
        "auditedAt": date.today().isoformat(),
        "results": results,
    }
    results_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print("---", flush=True)
    print(dict(Counter(r["verdict"] for r in results)), "wrote", results_path, flush=True)
    return results


def _strip_prior_audit_notes(notes: str) -> str:
    """Remove earlier precautionary / this-audit note clauses so re-apply is idempotent."""
    if not notes:
        return ""
    parts = [p.strip() for p in notes.split(" | ")]
    keep = []
    drop_substrings = (
        AUDIT_NOTE_TAG,
        "Downgraded 2026-08-02",
        "Activate-and-wait:",
        "Likely datapoint tool",
        "Waits for element pick",
        "No-op / misleading if nothing selected",
        "Requires an active fence",
        "May wait for interaction; same KEYIN probe gap",
        "precautionarily downgraded 2026-08-02",
        "CONFIRMED LIVE 2026-08-02 broken",
        "same probe/guard gap",
        "Prefer adjust_view",
    )
    for p in parts:
        if any(s in p for s in drop_substrings):
            continue
        keep.append(p)
    return " | ".join(keep).strip()


def apply_results(
    results_path: Path,
    registry_path: Path,
    restore_clean: bool,
    dry_run: bool,
) -> None:
    payload = json.loads(results_path.read_text(encoding="utf-8"))
    results = payload["results"]
    header, rows = _read_tsv(registry_path)
    by_op = {r["opName"]: r for r in rows}
    today = date.today().isoformat()
    armed_n = restore_n = 0

    for rec in results:
        op = rec["opName"]
        row = by_op.get(op)
        if not row:
            continue
        verdict = rec["verdict"]
        notes = _strip_prior_audit_notes(row.get("notes", ""))

        if verdict == "ARMED" and row.get("safetyStatus") in (
            "verified-headless-safe", "needs-testing",
        ):
            detail = rec.get("note") or (
                f"CommandName={rec.get('after', {}).get('commandName', '')!r}"
            )
            new_note = (
                f"{AUDIT_NOTE_TAG}: ARMED — {detail}. "
                f"Old keyin_batch probe SendReset masked this; "
                f"refusing headless execution."
            )
            row["safetyStatus"] = "unsafe-blocked"
            row["promotedDate"] = ""
            row["notes"] = (notes + " | " if notes else "") + new_note
            armed_n += 1
            print(f"APPLY ARMED->unsafe-blocked  {op}", flush=True)

        elif (
            verdict == "CLEAN"
            and restore_clean
            and row.get("safetyStatus") == "needs-testing"
        ):
            new_note = (
                f"{AUDIT_NOTE_TAG}: CLEAN — CommandName stayed at baseline "
                f"after SendKeyin (no tool left armed). Restored to "
                f"verified-headless-safe."
            )
            row["safetyStatus"] = "verified-headless-safe"
            row["promotedDate"] = today
            row["notes"] = (notes + " | " if notes else "") + new_note
            restore_n += 1
            print(f"APPLY CLEAN->verified  {op}", flush=True)

    print(f"armed_blocked={armed_n} restored_clean={restore_n} dry_run={dry_run}", flush=True)
    if dry_run:
        return

    lines = ["\t".join(header)]
    for row in rows:
        lines.append("\t".join(row.get(h, "") for h in header))
    registry_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {registry_path}", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_audit = sub.add_parser("audit", help="Live COM audit (DELETE.dgn)")
    p_audit.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_audit.add_argument("--out", type=Path, default=DEFAULT_RESULTS)
    p_audit.add_argument(
        "--scope",
        choices=("suspect", "verified", "needs-testing", "all-gated"),
        default="suspect",
        help="suspect=view/zoom/fit/window family; verified=all verified "
             "KEYINs; all-gated=verified+needs-testing",
    )
    p_audit.add_argument("--limit", type=int, default=0, help="Cap rows (0=all)")

    p_apply = sub.add_parser("apply", help="Write audit verdicts into registry")
    p_apply.add_argument("--results", type=Path, default=DEFAULT_RESULTS)
    p_apply.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_apply.add_argument(
        "--restore-clean",
        action="store_true",
        help="Promote needs-testing CLEAN rows back to verified-headless-safe",
    )
    p_apply.add_argument("--dry-run", action="store_true")

    p_run = sub.add_parser("run", help="audit then apply")
    p_run.add_argument("--registry", type=Path, default=DEFAULT_REGISTRY)
    p_run.add_argument("--out", type=Path, default=DEFAULT_RESULTS)
    p_run.add_argument("--scope", default="suspect",
                       choices=("suspect", "verified", "needs-testing", "all-gated"))
    p_run.add_argument("--limit", type=int, default=0)
    p_run.add_argument("--restore-clean", action="store_true")
    p_run.add_argument("--dry-run", action="store_true")

    args = ap.parse_args()
    if args.cmd == "audit":
        audit(args.registry, args.out, args.scope, args.limit)
    elif args.cmd == "apply":
        apply_results(args.results, args.registry, args.restore_clean, args.dry_run)
    elif args.cmd == "run":
        audit(args.registry, args.out, args.scope, args.limit)
        apply_results(args.out, args.registry, args.restore_clean, args.dry_run)


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    if len(sys.argv) >= 2 and sys.argv[1] == "_one_keyin_no_reset":
        keyin = " ".join(sys.argv[2:])
        raise SystemExit(_one_keyin_no_reset(keyin))
    main()
