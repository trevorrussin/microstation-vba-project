"""QA gate: task card -> standards + failure modes up front -> checks -> PASS/FAIL/UNCERTAIN record.

The gate of record for every drafting / CAD / engineering task in this repo. The agent
does not decide its own work is correct: `run` re-reads the result from the file,
evaluates requirements-as-code (qa/criteria.yaml) and the task's own expectations,
renders the result, gets an independent visual review, and writes
outputs/qa/tasks/<id>/latest.json. Stop hooks (scripts/qa_stop_hook.py) refuse to
let a session end while an open task is not PASS.

  python scripts/qa_gate.py start --type cad.drafting --desc "Place 619-311 channelizing" --extent 1000,2000,1500,2300
  python scripts/qa_gate.py expect --id E1 --desc "drums along taper" --where level=TWZCD --count ">=10"
  python scripts/qa_gate.py cite --doc hdm/chapt_16.pdf --page 41 --quote "..."
  python scripts/qa_gate.py run                      # exit 0 PASS / 1 FAIL / 2 UNCERTAIN
  python scripts/qa_gate.py review --ok --checked "taper length, drum spacing, no strays"
  python scripts/qa_gate.py close                    # only when latest run PASSED
  python scripts/qa_gate.py escalate --reason "..."  # hand to engineer (stop hook then allows stop)
  python scripts/qa_gate.py status

Task types: cad.drafting (MicroStation), cad.classify.curb (AutoCAD SGD), engineering.calc, generic.
"""
from __future__ import annotations

import argparse
import json
import glob
import os
import re
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
TASKS = os.path.join(ROOT, "outputs", "qa", "tasks")
TYPES = ("cad.drafting", "cad.classify.curb", "engineering.calc", "generic")
BLOCKING = {"FAIL", "UNCERTAIN"}


def _ops():
    """wztc_ops with a bridge attached (CLI uses the MCP bridge; chat_driver already set its own)."""
    import wztc_ops
    if getattr(wztc_ops, "_bridge", None) is None:
        from bridge_client import bridge
        wztc_ops.set_bridge(bridge)
    return wztc_ops


def _yaml(path):
    import yaml
    with open(os.path.join(ROOT, path), encoding="utf-8") as f:
        return yaml.safe_load(f)


def _tdir(tid):
    return os.path.join(TASKS, tid)


def load_task(tid=None):
    if tid:
        p = os.path.join(_tdir(tid), "task.json")
        if not os.path.exists(p):
            raise SystemExit(f"no task {tid}")
        return json.load(open(p, encoding="utf-8"))
    open_tasks = [t for t in all_tasks() if t["status"] == "in_progress"]
    if not open_tasks:
        raise SystemExit("no open task - run `qa_gate start` first")
    return sorted(open_tasks, key=lambda t: t["created"])[-1]


def all_tasks():
    out = []
    if os.path.isdir(TASKS):
        for d in os.listdir(TASKS):
            p = os.path.join(TASKS, d, "task.json")
            if os.path.exists(p):
                try:
                    out.append(json.load(open(p, encoding="utf-8")))
                except Exception:
                    pass
    return out


def save_task(t):
    os.makedirs(_tdir(t["id"]), exist_ok=True)
    with open(os.path.join(_tdir(t["id"]), "task.json"), "w", encoding="utf-8") as f:
        json.dump(t, f, indent=1)


def criteria_for(ttype):
    return [r for r in _yaml("qa/criteria.yaml") if ttype in r["applies"]]


def failure_modes_for(ttype):
    return [r for r in _yaml("qa/failure_modes.yaml") if ttype in r["applies"]]


# ------------------------------------------------------------------ change set (pre/post-edit snapshot)
SNAP_MAX = 20000


def take_snapshot(wztc_ops):
    """Every element in the active model (see qa/changeset.py). None when the read was truncated."""
    from qa import changeset
    r = wztc_ops.get_element_cad_props(min_id=0, ids="", max_rows=SNAP_MAX)
    return None if r.get("truncated") else changeset.snapshot(r["rows"])


def save_snapshot(t, snap):
    os.makedirs(_tdir(t["id"]), exist_ok=True)
    with open(os.path.join(_tdir(t["id"]), "snapshot.json"), "w", encoding="utf-8") as f:
        json.dump(snap, f)
    t["snapshot_n"] = len(snap)


def _load_snapshot(t):
    p = os.path.join(_tdir(t["id"]), "snapshot.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


# ------------------------------------------------------------------ start
def cmd_start(a):
    if a.type not in TYPES:
        raise SystemExit(f"--type must be one of {TYPES}")
    tid = a.id or (datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + re.sub(r"[^a-z0-9]+", "-", a.desc.lower())[:40].strip("-"))
    t = {"id": tid, "type": a.type, "desc": a.desc, "files": a.files or [], "status": "in_progress",
         "created": time.time(), "created_iso": datetime.now().isoformat(timespec="seconds"),
         "extent": [float(x) for x in a.extent.split(",")] if a.extent else None,
         "ids": a.ids or "", "text_heights": [float(x) for x in a.text_heights.split(",")] if a.text_heights else None,
         "is2d": not a.is3d, "expect": [], "citations": [], "calcs": [], "runs": 0, "fail_streak": 0,
         "may_modify": bool(getattr(a, "may_modify", False)), "candidates": []}
    notes = []
    if a.type == "cad.drafting" and not a.is3d:
        try:
            st = _ops().describe_drawing_state()
            if str(st.get("is3D", "")).lower() == "true":
                t["is2d"] = False
                notes.append("active model is 3D - Z=0 rule not applied")
        except Exception:
            pass
        save_task(t)
    if a.type == "cad.drafting" and not a.ids:
        try:
            wztc_ops = _ops()
            b = wztc_ops.get_max_element_id()
            t["baseline_max_id"] = b["maxId"]
            notes.append(f"baseline: max element id {b['maxId']:.0f} ({b['count']} elements) - checks cover elements created after this")
        except Exception as ex:
            t["baseline_max_id"] = None
            notes.append(f"WARNING baseline not recorded ({ex}); pass --ids at run time or restart with MicroStation open")
    if a.type == "cad.drafting" and not getattr(a, "no_snapshot", False):
        try:
            _ops()
            import op_recorder
            t["oplog_offset"] = op_recorder.log_size()  # edit ops recorded after this belong to the task
        except Exception:
            pass
        try:
            snap = take_snapshot(_ops())
            if snap is None:
                notes.append("WARNING pre-edit snapshot truncated; modified/deleted existing elements cannot be detected")
            else:
                save_snapshot(t, snap)
                notes.append(f"pre-edit snapshot: {len(snap)} elements (modified / deleted existing elements will be detected)")
        except Exception as ex:
            notes.append(f"WARNING pre-edit snapshot not taken ({ex}); only new elements will be detected")
    save_task(t)
    print(f"QA task {tid} started ({a.type})")
    for n in notes:
        print("  " + n)
    print("\nCRITERIA this task must pass:")
    for r in criteria_for(a.type):
        print(f"  {r['id']}: {r['req']}  [{r['source'][:70]}]")
    print("\nKNOWN FAILURE MODES for this task type (plan against these):")
    for f in failure_modes_for(a.type):
        print(f"  {f['id']}: {f['mode']}  -> detected by {f['detection'][:60]}")
    if not a.no_standards:
        try:
            import standards_index as si
            hits = si.search(a.desc, k=5, rerank=True)
            t["candidates"] = [{"cite": h["cite"], "doc": h["doc"], "applies": None, "reason": ""} for h in hits]
            try:  # the engineers' working procedures (Standards/Tutorials playbooks) are candidates too
                import tutorial_kb
                for pb in tutorial_kb.match_playbooks(a.desc):
                    t["candidates"].append({"cite": f"PLAYBOOK {pb['topic']}: {pb['title']}", "doc": f"playbook:{pb['topic']}",
                                            "applies": None, "reason": ""})
            except Exception:
                pass
            save_task(t)
            print("\nSTANDARDS likely governing this task (verify; cite with `qa_gate cite`):")
            for h in hits:
                print(f"  {h['cite']}  [{h['doc']}]")
        except Exception as ex:
            print(f"\n(standards search unavailable: {ex})")
    if t["candidates"]:
        print("\nDISPOSITION each candidate above: qa_gate dispose --cite \"<cite text>\" --applies yes|no --reason \"...\"")
    print("\nNEXT: add measurable expectations (`qa_gate expect ...`) before drafting.")
    return 0


def cmd_expect(a):
    t = load_task(a.task)
    row = {"id": a.id, "desc": a.desc}
    if getattr(a, "rel", None):
        if a.rel not in ("offset", "length", "angle") or a.value is None or not a.a:
            raise SystemExit("--rel offset|length|angle needs --a ID [--b ID] --value N [--tol T]")
        row.update(rel=a.rel, a=a.a, b=a.b, value=float(a.value), tol=float(a.tol or 0.05))
    elif a.count:
        row["where"] = dict(kv.split("=", 1) for kv in (a.where or []))
        row["count"] = a.count
    else:
        row["value"], row["target"] = a.value, a.target
    t["expect"] = [e for e in t["expect"] if e["id"] != a.id] + [row]
    save_task(t)
    print(f"{t['id']}: expectation {a.id} recorded ({len(t['expect'])} total)")
    return 0


def cmd_cite(a):
    t = load_task(a.task)
    t["citations"].append({"doc": a.doc, "page": a.page, "quote": a.quote, "req": a.req or ""})
    save_task(t)
    print(f"{t['id']}: citation recorded ({len(t['citations'])} total)")
    return 0


def cmd_calc(a):
    t = load_task(a.task)
    t["calcs"] = [c for c in t["calcs"] if c["id"] != a.id] + [{
        "id": a.id, "fn": a.fn, "inputs": json.loads(a.inputs), "result": a.result, "units": a.units, "criterion": a.criterion}]
    save_task(t)
    print(f"{t['id']}: calc {a.id} recorded")
    return 0


# ------------------------------------------------------------------ run
def _gather(t, a, ctx):
    tdir = _tdir(t["id"])
    lib_path = os.path.join(ROOT, "Standards", "derived", "nysdot-levels.tsv")
    if os.path.exists(lib_path):
        ctx["level_library"] = {ln.split("	", 1)[0] for ln in open(lib_path, encoding="utf-8")
                                if ln.strip() and not ln.startswith("#") and not ln.startswith("name	")}
    ctx.update(task_dir=tdir, rules=_yaml("qa/project_rules.yaml"), is2d=t.get("is2d", True),
               extent=t.get("extent"), text_heights=t.get("text_heights"), images=list(a.image or []))
    if t["type"] == "cad.drafting":
        try:
            wztc_ops = _ops()
            ids = a.ids or t.get("ids") or ""
            mn = t.get("baseline_max_id")
            if not ids and mn is None:
                raise RuntimeError("no baseline and no --ids")
            r = wztc_ops.get_element_cad_props(min_id=mn or 0, ids=ids, max_rows=3000)
            snap = _load_snapshot(t)
            if snap is not None:
                from qa import changeset
                allr = wztc_ops.get_element_cad_props(min_id=0, ids="", max_rows=SNAP_MAX)
                if allr.get("truncated"):
                    ctx["changes"] = {"error": f"post-edit read truncated at {SNAP_MAX} elements"}
                else:
                    ch = changeset.diff(snap, allr["rows"])
                    ctx["changes"] = ch
                    if not ids:  # created + modified existing elements are what the checks cover
                        r = {"rows": ch["created"] + ch["modified"], "truncated": False}
            ctx["elements"] = r["rows"]
            ctx["declared_ids"] = [i for i in ids.split(",") if i.strip()]
            if "oplog_offset" in t:
                import op_recorder
                ctx["oplog"] = op_recorder.read_since(t["oplog_offset"])
            rel_ids = {str(e[k]) for e in t.get("expect") or [] if e.get("rel") for k in ("a", "b") if e.get(k)}
            ctx["geom"] = {}
            for rid in rel_ids:
                try:
                    ctx["geom"][rid] = wztc_ops.get_element_vertices(rid).get("path_vertices") or []
                except Exception:
                    pass
            try:
                ctx["model_state"] = wztc_ops.describe_drawing_state()
                ctx["model_name"] = str(ctx["model_state"].get("modelName", ""))
            except Exception:
                ctx["model_state"], ctx["model_name"] = {}, ""
            if r.get("truncated"):
                ctx["elements_error"] = "read-back truncated at 3000 elements"
            levels = sorted({e.get("level") for e in r["rows"] if e.get("level")})
            if levels:
                defs = wztc_ops.list_level_defs(names=",".join(levels)).get("rows") or []
                ctx["level_defs"] = {d["name"]: d for d in defs}
            if not t.get("extent") and r["rows"]:
                xs = [float(e["rangeLowX"]) for e in r["rows"]] + [float(e["rangeHighX"]) for e in r["rows"]]
                ys = [float(e["rangeLowY"]) for e in r["rows"]] + [float(e["rangeHighY"]) for e in r["rows"]]
                ctx["render_extent"] = [min(xs), min(ys), max(xs), max(ys)]
            new_ids = {e["elementId"] for e in r["rows"]}
            if any("text" in (e.get("type") or "").lower() for e in r["rows"]):
                try:  # the sheet's existing text = the style authority for new text (qa/checks/cad_rules.text_style)
                    ref = wztc_ops.get_element_cad_props(min_id=0, ids="TEXT", max_rows=3000)
                    ctx["reference_texts"] = [e for e in ref["rows"] if e["elementId"] not in new_ids]
                except Exception as ex:
                    ctx["reference_texts_error"] = str(ex)
            # CAD-TXT-04: actual segment paths (not range AABB) for text-vs-line clearance
            line_paths = {}
            for e in r["rows"]:
                typ = (e.get("type") or "").lower()
                if "text" in typ or "cell" in typ or "dimension" in typ:
                    continue
                if not any(k in typ for k in ("line", "curve", "arc", "shape", "string", "complex")):
                    continue
                try:
                    v = wztc_ops.get_element_vertices(str(e["elementId"]))
                    pts = v.get("path_vertices") or v.get("vertices") or []
                    if len(pts) >= 2:
                        line_paths[str(e["elementId"])] = [
                            (float(p[0]), float(p[1])) if not isinstance(p, dict)
                            else (float(p.get("x", p.get("X"))), float(p.get("y", p.get("Y"))))
                            for p in pts
                        ]
                except Exception:
                    pass
            ctx["line_paths"] = line_paths
            from qa.checks import visual
            png = os.path.join(tdir, f"render-{t['runs']+1}.png")
            visual.render_elements(r["rows"], wztc_ops.get_element_vertices, png, f"{t['id']} - elements checked ({'ids given' if (a.ids or t.get('ids')) else 'created after baseline'})",
                                   extent=t.get("extent"), reference_texts=ctx.get("reference_texts"))
            ctx["images"].insert(0, png)
            try:
                import view_capture
                cap = view_capture.capture_microstation(os.path.join(tdir, f"capture-{t['runs']+1}.png"))
                ctx["images"].append(str(cap))
            except Exception:
                pass
        except Exception as ex:
            ctx["elements"] = []
            ctx["elements_error"] = f"element read-back failed: {ex}"
            ctx["render_error"] = str(ex)


def cmd_run(a):
    t = load_task(a.task)
    ctx: dict = {}
    _gather(t, a, ctx)
    import importlib
    results = []
    rows = criteria_for(t["type"])
    vis_row = None
    cov_rows = []
    ctx["results"] = {}
    for r in rows:
        if r["check"] == "visual.review":
            vis_row = r
            continue
        if r["check"] == "coverage.report":
            cov_rows.append(r)
            continue
        mod, fn = r["check"].split(".")
        try:
            res = getattr(importlib.import_module(f"qa.checks.{mod}"), fn)(t, ctx)
        except Exception as ex:
            res = {"status": "UNCERTAIN", "observed": f"checker error: {ex}", "expected": "", "evidence": [], "fix": ""}
        res.update(id=r["id"], req=r["req"], source=r["source"], blocking=r.get("blocking", True))
        results.append(res)
        ctx["results"][r["id"]] = res
    if vis_row:
        from qa.checks import visual
        if ctx.get("images") and not a.no_vision:
            vr = visual.vision_review(ctx["images"], t, failure_modes_for(t["type"]),
                                      [r for r in rows if r["id"].startswith("CAD-")])
            vr["t"] = time.time()
            json.dump(vr, open(os.path.join(_tdir(t["id"]), "vision-review.json"), "w"), indent=1)
        res = visual.review(t, ctx)
        res.update(id=vis_row["id"], req=vis_row["req"], source=vis_row["source"], blocking=True)
        results.append(res)
        ctx["results"][vis_row["id"]] = res
    for r in cov_rows:  # last: needs every other result
        from qa.checks import coverage
        try:
            res = coverage.report(t, ctx)
        except Exception as ex:
            res = {"status": "WARN", "observed": f"coverage report error: {ex}", "expected": "", "evidence": [], "fix": ""}
        res.update(id=r["id"], req=r["req"], source=r["source"], blocking=False)
        results.append(res)
    blocking = [r for r in results if r["blocking"] and r["status"] in BLOCKING]
    overall = "PASS" if not blocking else ("FAIL" if any(r["status"] == "FAIL" for r in blocking) else "UNCERTAIN")
    t["runs"] += 1
    t["fail_streak"] = 0 if overall == "PASS" else t.get("fail_streak", 0) + 1
    t["last_status"] = overall
    t["last_run"] = time.time()
    save_task(t)
    rec = {"task": t["id"], "type": t["type"], "run": t["runs"], "at": datetime.now().isoformat(timespec="seconds"),
           "overall": overall, "images": ctx.get("images"), "results": results,
           "n_elements": len(ctx.get("elements") or []), "coverage": ctx.get("coverage")}
    with open(os.path.join(_tdir(t["id"]), "latest.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=1)
    with open(os.path.join(_tdir(t["id"]), f"record-{t['runs']}.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=1)
    print(f"QA {t['id']} run {t['runs']}: {overall}")
    for r in results:
        mark = {"PASS": "ok ", "FAIL": "XX ", "UNCERTAIN": "?? ", "WARN": "!  ", "NA": "-  "}.get(r["status"], "   ")
        print(f"  {mark}{r['id']} {r['status']}: {r['observed']}")
        if r["status"] in BLOCKING | {"WARN"}:
            for e in r["evidence"][:8]:
                print(f"       - {e}")
            if r.get("fix"):
                print(f"       FIX: {r['fix']}")
    if ctx.get("images"):
        print("  images: " + ", ".join(ctx["images"]))
    op = next((r for r in results if r["id"] == "OP-01"), None)
    cad_ok = all(r["status"] not in BLOCKING for r in results if r["blocking"] and r["id"] not in ("OP-01", "REQ-GEOM-01"))
    if op is not None:
        print(f"  CAD standards checks: {'PASS' if cad_ok else 'NOT PASSING'}   |   Requested edit verified: {op['status']} ({op['observed'][:90]})")
    if overall != "PASS":
        print("\nNOT DONE. Fix the XX/?? rows, then run `qa_gate run` again. "
              "If it cannot be fixed from evidence, `qa_gate escalate --reason ...` and tell the engineer.")
    return {"PASS": 0, "FAIL": 1, "UNCERTAIN": 2}[overall]


# ------------------------------------------------------------------ review / close
def cmd_review(a):
    t = load_task(a.task)
    p = os.path.join(_tdir(t["id"]), "agent-review.json")
    ar = json.load(open(p)) if os.path.exists(p) else {"disputes": {}}
    ar["t"] = time.time()
    renders = sorted(glob.glob(os.path.join(_tdir(t["id"]), "render-*.png")), key=os.path.getmtime)
    if renders:
        from qa.checks import visual
        ar["render_sha"] = visual._sha(renders[-1])
    if a.ok:
        ar.update(ok=True, issues=[], checked=a.checked or "")
    if a.issue:
        ar.update(ok=False, issues=ar.get("issues", []) + [a.issue])
    if a.dispute:
        n, reason = a.dispute
        ar.setdefault("disputes", {})[str(n)] = reason
    json.dump(ar, open(p, "w"), indent=1)
    print(f"{t['id']}: review recorded {json.dumps({k: v for k, v in ar.items() if k != 't'})[:300]}")
    print("Run `qa_gate run` to re-evaluate.")
    return 0


def cmd_close(a):
    t = load_task(a.task)
    if t.get("last_status") != "PASS":
        print(f"{t['id']}: cannot close - last run is {t.get('last_status') or 'none'}. Run `qa_gate run` until PASS.")
        return 1
    t["status"] = "closed"
    t["closed_iso"] = datetime.now().isoformat(timespec="seconds")
    save_task(t)
    print(f"{t['id']}: closed (PASS, {t['runs']} runs)")
    return 0


def cmd_escalate(a):
    t = load_task(a.task)
    t["status"] = "escalated"
    t["escalation"] = a.reason
    save_task(t)
    print(f"{t['id']}: ESCALATED - {a.reason}\nTell the engineer exactly this, with the latest record: "
          f"{os.path.join(_tdir(t['id']), 'latest.json')}")
    return 0


def cmd_dispose(a):
    t = load_task(a.task)
    if a.applies not in ("yes", "no"):
        raise SystemExit("--applies must be yes or no")
    if len((a.reason or "").strip()) < 10:
        raise SystemExit("--reason must say why (>= 10 characters)")
    hits = [c for c in t.get("candidates", []) if a.cite == "all" or a.cite.lower() in c["cite"].lower()]
    if not hits:
        raise SystemExit("no candidate matches --cite; see the list printed by `qa_gate start`")
    for c in hits:
        c["applies"], c["reason"] = a.applies, a.reason
    save_task(t)
    left = sum(1 for c in t["candidates"] if c.get("applies") not in ("yes", "no"))
    print(f"{t['id']}: {len(hits)} candidate(s) marked applies={a.applies}; {left} still open")
    return 0


def cmd_status(a):
    ts = [load_task(a.task)] if a.task else sorted(all_tasks(), key=lambda t: t["created"])[-15:]
    for t in ts:
        print(f"{t['id']}  {t['type']}  {t['status']}  last={t.get('last_status')}  runs={t.get('runs')}  {t['desc'][:60]}")
    return 0


def run_inline(ttype: str, desc: str, ids: str = "", expect: list | None = None, extent: list | None = None,
               citations: list | None = None, task_id: str | None = None, no_vision: bool = False,
               snapshot: dict | None = None, oplog_offset: int | None = None) -> dict:
    """In-process gate run (used by the in-MicroStation chat agent). Reuses an open task with the
    same id, else creates one. Returns the run record (overall, results, images)."""
    import contextlib
    import io
    t = None
    if task_id and os.path.exists(os.path.join(_tdir(task_id), "task.json")):
        t = json.load(open(os.path.join(_tdir(task_id), "task.json"), encoding="utf-8"))
    if t is None:
        ns = argparse.Namespace(type=ttype, desc=desc, files=[], extent=",".join(map(str, extent)) if extent else None,
                                ids=ids, id=task_id, text_heights=None, is3d=False, no_standards=True,
                                no_snapshot=True)
        with contextlib.redirect_stdout(io.StringIO()):
            cmd_start(ns)
        t = load_task(task_id) if task_id else sorted(
            [x for x in all_tasks() if x["status"] == "in_progress"], key=lambda x: x["created"])[-1]
    t["ids"] = ids or t.get("ids", "")
    if oplog_offset is not None and "oplog_offset" not in t:
        t["oplog_offset"] = oplog_offset
    if snapshot is not None and not os.path.exists(os.path.join(_tdir(t["id"]), "snapshot.json")):
        save_snapshot(t, snapshot)  # taken by the caller BEFORE the edits
    if expect is not None:
        t["expect"] = expect
    if citations is not None:
        t["citations"] = citations
    if extent:
        t["extent"] = extent
    save_task(t)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        cmd_run(argparse.Namespace(task=t["id"], image=[], ids=t["ids"], no_vision=no_vision))
    rec = json.load(open(os.path.join(_tdir(t["id"]), "latest.json"), encoding="utf-8"))
    rec["summary"] = buf.getvalue()
    return rec


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("start"); s.add_argument("--type", required=True); s.add_argument("--desc", required=True)
    s.add_argument("--files", nargs="*"); s.add_argument("--extent"); s.add_argument("--ids"); s.add_argument("--id")
    s.add_argument("--text-heights"); s.add_argument("--is3d", action="store_true"); s.add_argument("--no-standards", action="store_true")
    s.add_argument("--may-modify", action="store_true", help="engineer asked to edit/move/delete EXISTING elements")
    dp = sub.add_parser("dispose"); dp.add_argument("--task"); dp.add_argument("--cite", required=True)
    dp.add_argument("--applies", required=True); dp.add_argument("--reason", required=True)
    e = sub.add_parser("expect"); e.add_argument("--task"); e.add_argument("--id", required=True); e.add_argument("--desc", default="")
    e.add_argument("--where", nargs="*"); e.add_argument("--count"); e.add_argument("--value"); e.add_argument("--target")
    e.add_argument("--rel", help="offset|length|angle: geometric relation measured from saved geometry")
    e.add_argument("--a"); e.add_argument("--b"); e.add_argument("--tol")
    c = sub.add_parser("cite"); c.add_argument("--task"); c.add_argument("--doc", required=True); c.add_argument("--page", type=int, required=True)
    c.add_argument("--quote", required=True); c.add_argument("--req")
    k = sub.add_parser("calc"); k.add_argument("--task"); k.add_argument("--id", required=True); k.add_argument("--fn", required=True)
    k.add_argument("--inputs", required=True); k.add_argument("--result", type=float, required=True); k.add_argument("--units", required=True)
    k.add_argument("--criterion", required=True)
    r = sub.add_parser("run"); r.add_argument("--task"); r.add_argument("--image", nargs="*"); r.add_argument("--ids")
    r.add_argument("--no-vision", action="store_true")
    v = sub.add_parser("review"); v.add_argument("--task"); v.add_argument("--ok", action="store_true"); v.add_argument("--checked")
    v.add_argument("--issue"); v.add_argument("--dispute", nargs=2, metavar=("N", "REASON"))
    cl = sub.add_parser("close"); cl.add_argument("--task")
    es = sub.add_parser("escalate"); es.add_argument("--task"); es.add_argument("--reason", required=True)
    st = sub.add_parser("status"); st.add_argument("--task")
    a = ap.parse_args()
    return {"start": cmd_start, "expect": cmd_expect, "cite": cmd_cite, "calc": cmd_calc, "run": cmd_run,
            "review": cmd_review, "close": cmd_close, "escalate": cmd_escalate, "status": cmd_status,
            "dispose": cmd_dispose}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
