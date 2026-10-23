"""Completeness checks for the QA gate: change set, candidate-standard dispositions, requirement coverage.

changes(task, ctx)        CAD-CHG-01  modified / deleted existing elements must be declared (--may-modify)
dispositions(task, ctx)   STD-DISP-01 every standards-search candidate is marked applies / does not apply
report(task, ctx)         COV-01      applicable requirements (qa/applicability.yaml, derived from the measured
                                      change set) with their verification status; never claims full compliance
"""
from __future__ import annotations

import os
import re

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
MAX_IDS = 25


def _res(status, observed="", expected="", evidence=None, fix=""):
    return {"status": status, "observed": observed, "expected": expected, "evidence": (evidence or [])[:MAX_IDS], "fix": fix}


def registry():
    with open(os.path.join(ROOT, "qa", "applicability.yaml"), encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------- CAD-CHG-01
def changes(task, ctx):
    ch = ctx.get("changes")
    if ch is None:
        return _res("WARN", "no pre-edit snapshot: modified or deleted existing elements cannot be detected",
                    "pre/post change set", fix="Start the task before editing so `qa_gate start` can snapshot the model.")
    if ch.get("error"):
        return _res("UNCERTAIN", ch["error"], "complete pre/post change set",
                    fix="Re-run with a smaller model or raise max_rows; a truncated read cannot prove nothing else changed.")
    nc, nm, nd = len(ch["created"]), len(ch["modified"]), len(ch["deleted"])
    summary = f"{nc} created, {nm} modified, {nd} deleted"
    if not (nm or nd):
        return _res("PASS", summary)
    declared = {str(i).strip() for i in ctx.get("declared_ids") or []}  # ids the agent explicitly touched (chat path)
    mod = [r for r in ch["modified"] if str(r["elementId"]) not in declared]
    dele = {i: o for i, o in ch["deleted"].items() if str(i) not in declared}
    ev = [f"modified {r['elementId']} ({', '.join(ch['modified_fields'].get(str(r['elementId']), []))})" for r in mod]
    ev += [f"deleted {i} {o.get('type')} on {o.get('level')} {o.get('text', '')[:24]}" for i, o in dele.items()]
    if not ev:
        return _res("PASS", summary + " (all changes to existing elements are ones the agent explicitly touched)")
    if task.get("may_modify"):
        return _res("PASS", summary + " (engineer-declared edit of existing elements; modified ones are checked below)", evidence=ev)
    return _res("FAIL", f"{summary}: existing elements changed but the task did not declare it",
                "only new elements, unless the task declares edits to existing ones", ev,
                "If the request was to edit/move/delete existing elements, record `qa_gate start --may-modify`. "
                "Otherwise undo the unintended change.")


# ---------------------------------------------------------------- STD-DISP-01
def dispositions(task, ctx):
    cands = task.get("candidates") or []
    if not cands:
        return _res("NA", "no standards candidates recorded at start")
    open_ = [c["cite"] for c in cands if c.get("applies") not in ("yes", "no") or len((c.get("reason") or "").strip()) < 10]
    if open_:
        return _res("UNCERTAIN", f"{len(open_)} of {len(cands)} candidate standards sections not dispositioned",
                    "each candidate marked applies / does not apply with a reason", open_,
                    'qa_gate dispose --cite "<text of the cite>" --applies yes|no --reason "why (>=10 chars)"')
    yes = sum(1 for c in cands if c["applies"] == "yes")
    return _res("PASS", f"{len(cands)} candidates dispositioned ({yes} apply, {len(cands) - yes} do not)")


# ---------------------------------------------------------------- COV-01
def _kind_matches(kind, el, model_name):
    m = kind.get("match") or {}
    for k, v in m.items():
        if k == "deleted":
            return False  # handled separately
        if k == "model":
            if not re.search(v, model_name or "", re.I):
                return False
        elif k == "cell":
            if not re.search(v, el.get("cellName") or "", re.I):
                return False
        elif not re.search(v, str(el.get(k) or ""), re.I):
            return False
    return True


def applicable(ctx, reg=None):
    """-> (requirements {req_id: (req, [kind ids])}, unmapped element kinds {description: count})."""
    reg = reg or registry()
    pats = [re.compile(p, re.I) for p in (ctx.get("rules") or {}).get("ignore_levels", [])]
    els = [e for e in ctx.get("elements") or [] if not any(p.search(e.get("level") or "") for p in pats)]
    model = ctx.get("model_name") or ""
    reqs, unmapped = {}, {}
    for kind in reg["kinds"]:
        if (kind.get("match") or {}).get("deleted"):
            if (ctx.get("changes") or {}).get("deleted"):
                for r in kind["requirements"]:
                    reqs.setdefault(r["id"], (r, []))[1].append(kind["id"])
            continue
        if any(_kind_matches(kind, e, model) for e in els):
            for r in kind["requirements"]:
                reqs.setdefault(r["id"], (r, []))[1].append(kind["id"])
    specific = [k for k in reg["kinds"]
                if k["id"] != "K-ALL" and not (k.get("match") or {}).get("deleted") and "model" not in (k.get("match") or {})]
    for e in els:
        if not any(_kind_matches(k, e, model) for k in specific):
            key = f"{e.get('type')} on {e.get('level')}"
            unmapped[key] = unmapped.get(key, 0) + 1
    return reqs, unmapped


def _cite(r):
    pb = f"playbook {r['playbook']}" if r.get("playbook") else ""
    if not r.get("doc"):
        return pb or "task rule"
    return f"{r['doc']} §{r['section']}" + (f" p.{r['page']}" if r.get("page") else "") + (f" [{pb}]" if pb else "")


def report(task, ctx):
    reg = registry()
    results = ctx.get("results") or {}
    reqs, unmapped = applicable(ctx, reg)
    rows, notes = [], []
    for rid, (r, kinds) in sorted(reqs.items()):
        if r["check"] == "manual":
            status = "engineer-review"
        else:
            st = (results.get(r["check"]) or {}).get("status")
            status = {"PASS": "verified", "FAIL": "FAILED"}.get(st, "unverified")
        rows.append({"requirement": rid, "req": r["req"], "source": _cite(r), "check": r["check"], "status": status,
                     "kinds": sorted(set(kinds))})
        if status == "engineer-review":
            notes.append(f"engineer review: {r['req']}  [{_cite(r)}]")
        elif status == "unverified":
            notes.append(f"unverified: {r['req']}  [{_cite(r)}]  (check {r['check']} did not return PASS)")
    gaps = []
    for c in task.get("candidates") or []:
        if c.get("applies") != "yes":
            continue
        if not any(r.get("doc") and r["doc"] in c.get("doc", "") and r.get("section") and r["section"] in c.get("cite", "")
                   for r, _ in reqs.values()):
            gaps.append(f"candidate marked applicable but no registry check covers it: {c['cite']}")
    ctx["coverage"] = {"registry_version": reg.get("version"), "requirements": rows,
                       "unmapped_element_kinds": unmapped, "candidate_gaps": gaps}
    n = len(rows)
    v = sum(1 for r in rows if r["status"] == "verified")
    bad = sum(1 for r in rows if r["status"] == "FAILED")
    man = sum(1 for r in rows if r["status"] == "engineer-review")
    unv = sum(1 for r in rows if r["status"] == "unverified")
    obs = f"{v} of {n} applicable requirements verified automatically; {man} need engineer review; {unv} unverified; {bad} failed"
    extra = notes + gaps + [f"no registry entry beyond the universal rules for: {k} x{c}" for k, c in unmapped.items()]
    if n == 0:
        return _res("WARN", "no changed elements, so no applicable requirements were derived", evidence=extra)
    if extra:
        return _res("WARN", obs + ". PASS here means the implemented checks passed, not full NYSDOT/DDC compliance.",
                    evidence=extra, fix="Review the listed items yourself, or add an automated check and registry link for them.")
    return _res("PASS", obs)
