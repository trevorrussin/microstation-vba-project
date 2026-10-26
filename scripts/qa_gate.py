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


def read_state(wztc_ops):
    """Every element in the active model plus a vertex hash per linear element (see qa/changeset.py).
    None when the read was truncated."""
    from qa import changeset
    r = wztc_ops.get_element_cad_props(min_id=0, ids="", max_rows=SNAP_MAX)
    if r.get("truncated"):
        return None
    rows = r["rows"]
    lin = [str(x["elementId"]) for x in rows if changeset.is_linear(x)]
    vh = {}
    for e in lin[:changeset.VH_MAX]:
        # ONE op per request: multi-op batches return vertex rows attributed to the wrong elements (found live 2026-10-23)
        o = wztc_ops._bridge._call_batch_impl([("GET_ELEMENT_VERTICES", {"elementId": e})])[0]
        if str(o.get("status")) == "OK" and o.get("rows"):
            vh[e] = changeset.vhash([[p["x"], p["y"], p.get("z") or 0] for p in o["rows"]])
    nonlin = [x for x in rows if not changeset.is_linear(x)]
    scope = {"covers": "active model only", "elements": len(rows), "linear": len(lin), "vertex_hashed": len(vh),
             "vertex_hash_cap": changeset.VH_MAX, "vh_capped": len(lin) > changeset.VH_MAX,
             "geometry_by_properties_only": len(nonlin), "references_covered": False, "other_models_covered": False}
    return {"rows": rows, "vh": vh, "vh_capped": len(lin) > changeset.VH_MAX, "scope": scope}


def model_identity(wztc_ops) -> str:
    st = wztc_ops.describe_drawing_state()
    return f"{st.get('fileName', '')}|{st.get('modelName', '')}"


def state_fingerprint(wztc_ops):
    """-> (fingerprint, model identity, state) of the active model right now, or (None, model, None) if unreadable."""
    from qa import changeset
    model = model_identity(wztc_ops)
    state = read_state(wztc_ops)
    if state is None:
        return None, model, None
    return changeset.fingerprint(state["rows"], state["vh"], model), model, state


def take_snapshot(wztc_ops):
    """Every element in the active model (see qa/changeset.py). None when the read was truncated."""
    from qa import changeset
    st = read_state(wztc_ops)
    return None if st is None else changeset.snapshot(st["rows"], st["vh"])


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
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import ladders
    import modes
    mode = getattr(a, "mode", None) or "production"
    t["mode"] = mode
    if mode == "development":
        scratch = (getattr(a, "scratch", None) or "").strip()
        if not scratch or len((getattr(a, "mode_reason", None) or "").strip()) < 15:
            raise SystemExit("development mode needs --scratch <file.dgn|none> and --mode-reason \"why this is capability development\" (>= 15 characters)")
        if scratch.lower() != "none":
            try:
                active = str(_ops().describe_drawing_state().get("fileName", ""))
            except Exception as ex:
                raise SystemExit(f"cannot confirm the scratch file is the active file ({ex}). Open MicroStation on {scratch}, or use --scratch none for code-only work.")
            if active.lower() != scratch.lower():
                raise SystemExit(f"the active file is {active!r}, not the scratch file {scratch!r}. Development work never runs on another file: switch it (the engineer does that), or use --scratch none.")
        t["scratch"], t["mode_reason"] = scratch, a.mode_reason.strip()
    lad = ladders.match(a.desc)
    if lad:
        ladders.attach(t, lad)
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
            t["model"] = model_identity(_ops())  # the snapshot is of THIS model; run/close refuse a different one
        except Exception:
            t["model"] = ""
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
    print(f"QA task {tid} started ({a.type}, {mode} mode" + (f", scratch {t['scratch']}" if mode == "development" else "") + ")")
    if mode == "production":
        print("  PRODUCTION: established capabilities only. Editing the tooling (bridge, VBA, add-in, registry, contracts, QA code, tests) is refused "
              "until this task ends or the engineer starts a development task.")
    elif t["scratch"].lower() == "none":
        print("  DEVELOPMENT, code only: tooling may change; every drawing edit through the bridge is refused.")
    for n in notes:
        print("  " + n)
    print("\nCRITERIA this task must pass:")
    for r in criteria_for(a.type):
        print(f"  {r['id']}: {r['req']}  [{r['source'][:70]}]")
    if a.type == "cad.drafting":
        from qa.checks.operations_check import intent_lines
        print("\nINTENT the gate will hold the operations to (restart with a corrected --desc if this misreads the request):")
        for ln in intent_lines(a.desc):
            print("  " + ln)
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
            try:  # written procedures the matched playbooks cite (PDF / decks / sheets): read them, do not summarise
                import retrieval
                tops = [c["doc"].split(":", 1)[1] for c in t["candidates"] if c["doc"].startswith("playbook:")]
                t["sources"] = [{"cite": d["cite"], "doc": "tutorial:" + d["key"], "applies": None, "reason": ""}
                                for d in retrieval.doc_sources_for(tops)]
            except Exception:
                t["sources"] = []
            save_task(t)
            print("\nSTANDARDS likely governing this task (verify; cite with `qa_gate cite`):")
            for h in hits:
                print(f"  {h['cite']}  [{h['doc']}]")
            for sr in t.get("sources", []):
                print(f"  {sr['cite']}")
        except Exception as ex:
            print(f"\n(standards search unavailable: {ex})")
    try:
        import capabilities
        print("\n" + capabilities.format_result(capabilities.find(a.desc, 4, mode)))
    except Exception as ex:
        print(f"(capability search unavailable: {ex})")
    if lad:
        print("\n" + ladders.summary(lad))
    try:
        import experience
        past = experience.for_task(a.desc)
        if past:
            print("\n" + past)
    except Exception:
        pass
    try:
        import retrieval
        need = retrieval.gaps(t)
        if need:
            print("\nRESEARCH FIRST - drawing edits are refused until each is done:")
            for g in need:
                print("  " + g)
            print("  read with get_tutorial_playbook / get_standard_section / read_standard_pages / read_tutorial, then "
                  "qa_gate dispose --cite \"<cite>\" --applies yes|no --reason \"...\" (yes needs the read)")
    except Exception:
        pass
    if t["candidates"]:
        print("\nDISPOSITION each candidate above: qa_gate dispose --cite \"<cite text>\" --applies yes|no --reason \"...\"")
    print("\nNEXT: add measurable expectations (`qa_gate expect ...`) before drafting.")
    return 0


def cmd_expect(a):
    t = load_task(a.task)
    row = {"id": a.id, "desc": a.desc}
    if getattr(a, "rel", None):
        from qa.checks.operations_check import REL_KINDS
        if a.rel not in REL_KINDS or not a.a:
            raise SystemExit(f"--rel {'|'.join(REL_KINDS)} needs --a ID [--b ID] [--value N] [--tol T] [--params JSON]")
        row.update(rel=a.rel, a=a.a, b=a.b, tol=float(a.tol or 0.05))
        if a.value is not None:
            row["value"] = float(a.value)
        if getattr(a, "params", None):
            row.update(json.loads(a.params))  # e.g. '{"min":14.5,"ve":5,"x":[4800,4850]}' or '{"side":"left"}' / '{"target":"above"}'
        if a.rel in ("offset", "length", "angle", "offset_curve") and "value" not in row:
            raise SystemExit(f"--rel {a.rel} needs --value N")
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
            from qa.checks import process
            ctx["events"] = process.events_for_task(t)
        except Exception:
            pass
        try:
            wztc_ops = _ops()
            ids = a.ids or t.get("ids") or ""
            mn = t.get("baseline_max_id")
            if not ids and mn is None:
                raise RuntimeError("no baseline and no --ids")
            r = wztc_ops.get_element_cad_props(min_id=mn or 0, ids=ids, max_rows=3000)
            snap = _load_snapshot(t)
            from qa import changeset
            state = read_state(wztc_ops)
            cur_model = model_identity(wztc_ops)
            ctx["run_model"] = cur_model
            if state is not None:
                ctx["state_fp"] = changeset.fingerprint(state["rows"], state["vh"], cur_model)
                ctx["state_vh_capped"] = state["vh_capped"]
                ctx["state_scope"] = state["scope"]
                ctx["state_rows"] = state["rows"]
            if t.get("model") and t["model"] != cur_model:
                ctx["model_mismatch"] = f"task started in {t['model']} but the active model is {cur_model}"
            elif snap is not None:
                if state is None:
                    ctx["changes"] = {"error": f"post-edit read truncated at {SNAP_MAX} elements"}
                else:
                    ch = changeset.diff(snap, state["rows"], state["vh"])
                    ctx["changes"] = ch
                    if not ids:  # created + modified existing elements are what the checks cover
                        r = {"rows": ch["created"] + ch["modified"], "truncated": False}
            ctx["elements"] = r["rows"]
            ctx["declared_ids"] = [i for i in ids.split(",") if i.strip()]
            if "oplog_offset" in t:
                import op_recorder
                ctx["oplog"] = op_recorder.read_since(t["oplog_offset"])
                declared = {str(i).strip() for i in ctx["declared_ids"]}
                if declared:  # 2026-10-25: elements named with --ids may have been placed in an earlier task; their records are still in the persistent op log
                    seen = {json.dumps(r, sort_keys=True) for r in ctx["oplog"]}
                    older = [r for r in op_recorder.read_since(0)
                             if (set(map(str, r.get("new_ids") or [])) & declared) and json.dumps(r, sort_keys=True) not in seen]
                    ctx["oplog"] = sorted(older + ctx["oplog"], key=lambda r: r.get("ts", 0))
            rel_ids = {str(e[k]) for e in t.get("expect") or [] if e.get("rel") for k in ("a", "b") if e.get(k)}
            ctx["geom"] = {}
            if state is not None:  # range boxes of the elements named in relations (text and cells have no path)
                ctx["geom_boxes"] = {str(x["elementId"]): tuple(float(x[k]) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
                                     for x in state["rows"] if str(x["elementId"]) in rel_ids}
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
            try:  # all text in the model: style authority for new text (cad_rules.text_style) and the whole-model overlap audit
                ref = wztc_ops.get_element_cad_props(min_id=0, ids="TEXT", max_rows=3000)
                ctx["model_texts"] = ref["rows"]
                ctx["model_texts_truncated"] = bool(ref.get("truncated"))
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
                                   extent=t.get("extent"), reference_texts=ctx.get("reference_texts"),
                                   context_rows=state["rows"] if state else None)
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


# ------------------------------------------------------------------ engineer sign-off (independent of the agent)
KEY_PATH = os.path.join(os.path.expanduser("~"), ".qa_signoff_key")


def _signoff_key(create: bool = False):
    import secrets
    if not os.path.exists(KEY_PATH):
        if not create:
            return None
        with open(KEY_PATH, "w", encoding="utf-8") as f:
            f.write(secrets.token_hex(32))
    return open(KEY_PATH, encoding="utf-8").read().strip().encode()


def _sign(key: bytes, tid: str, items: list, fp: str) -> str:
    import hashlib
    import hmac
    return hmac.new(key, json.dumps([tid, sorted(items), fp or ""]).encode(), hashlib.sha256).hexdigest()


def load_signoff(t, fp):
    """Items the engineer signed off for THIS drawing state. A sign-off made for another state, with a bad
    signature, or without the engineer's key does not count (the agent cannot produce one)."""
    p = os.path.join(_tdir(t["id"]), "signoff.json")
    if not os.path.exists(p):
        return {"items": []}
    so = json.load(open(p, encoding="utf-8"))
    key = _signoff_key()
    if key is None or so.get("sig") != _sign(key, t["id"], so.get("items") or [], so.get("fp") or ""):
        return {"items": [], "invalid": "signature does not verify"}
    if t["type"] == "cad.drafting" and (so.get("fp") or "") != (fp or ""):
        return {"items": [], "invalid": "drawing changed since the sign-off"}
    return {"items": list(so.get("items") or []), "by": so.get("by"), "at": so.get("at"), "reason": so.get("reason")}


def _is_tty() -> bool:
    return sys.stdin.isatty() and sys.stdout.isatty()


def cmd_signoff(a):
    """Engineer-only: needs an interactive terminal; records items against the current drawing state."""
    t = load_task(a.task)
    if not _is_tty():
        print("signoff needs an interactive terminal. The engineer runs this in their own shell; an agent cannot sign off.")
        return 1
    latest = os.path.join(_tdir(t["id"]), "latest.json")
    need = json.load(open(latest, encoding="utf-8")).get("needs_signoff", []) if os.path.exists(latest) else []
    items = [x.strip() for x in (a.items or ",".join(n["id"] for n in need)).split(",") if x.strip()]
    if not items:
        print("nothing to sign off (run `qa_gate run` first, or pass --items)")
        return 1
    print(f"Task {t['id']}: {t['desc']}\nYou are signing off these items:")
    for i in items:
        print(f"  {i}: " + next((n["why"] for n in need if n["id"] == i), ""))
    fp = None
    if t["type"] == "cad.drafting":
        fp, _, _ = state_fingerprint(_ops())
        if fp is None:
            print("cannot read the drawing state; open MicroStation and retry")
            return 1
    if input("Type SIGN OFF to confirm you reviewed the drawing for these items: ").strip() != "SIGN OFF":
        print("not signed")
        return 1
    key = _signoff_key(create=True)
    reason = a.reason or input("Note (optional): ")
    so = {"items": items, "fp": fp or "", "by": os.environ.get("USERNAME", ""), "at": datetime.now().isoformat(timespec="seconds"),
          "reason": reason, "sig": _sign(key, t["id"], items, fp or "")}
    with open(os.path.join(_tdir(t["id"]), "signoff.json"), "w", encoding="utf-8") as f:
        json.dump(so, f, indent=1)
    print("signed. Re-run `qa_gate run`.")
    return 0


def _doc_revisions(docs):
    """sha256 + fetch date of each cited standards document (Standards/registry.json), so a PASS names the edition used."""
    try:
        reg = json.load(open(os.path.join(ROOT, "Standards", "registry.json"), encoding="utf-8"))["documents"]
    except Exception:
        return {}
    out = {}
    for d in docs:
        if not d or d.startswith("playbook:"):
            continue
        for x in reg:
            if x.get("path", "").endswith(d) or d.endswith(os.path.basename(x.get("path", "?"))):
                out[d] = f"{x.get('sha256', '')[:12]} fetched {x.get('fetched', '')}"
                break
    return out


def _file_sha(path):
    import hashlib
    try:
        return hashlib.sha256(open(path, "rb").read()).hexdigest()[:12]
    except OSError:
        return ""


def _file_stamp(path):
    try:
        st = os.stat(path)
        return f"{st.st_size} bytes, modified {datetime.fromtimestamp(st.st_mtime).isoformat(timespec='seconds')}"
    except OSError:
        return ""


def _provenance(t, ctx):
    import subprocess
    try:
        sha = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, capture_output=True, text=True, timeout=10).stdout.strip()
        dirty = bool(subprocess.run(["git", "status", "--porcelain", "qa", "scripts/qa_gate.py"], cwd=ROOT, capture_output=True,
                                    text=True, timeout=10).stdout.strip())
    except Exception:
        sha, dirty = "", None
    docs = sorted({c.get("doc", "") for c in t.get("candidates", [])} | {c.get("doc", "") for c in t.get("citations", [])})
    cov = ctx.get("coverage") or {}
    return {"gate_commit": sha, "gate_uncommitted_changes": dirty,
            "criteria_sha": _file_sha(os.path.join(ROOT, "qa", "criteria.yaml")),
            "applicability_sha": _file_sha(os.path.join(ROOT, "qa", "applicability.yaml")),
            "applicability_version": cov.get("registry_version"),
            "standards_registry_sha": _file_sha(os.path.join(ROOT, "Standards", "registry.json")),
            "standards_index": _file_stamp(os.path.join(ROOT, "Standards", "index", "standards.sqlite")),
            "standards_documents": _doc_revisions(docs), "model": ctx.get("run_model"), "state_fingerprint": ctx.get("state_fp"),
            "fingerprint_scope": ctx.get("state_scope"),
            "snapshot_sha": _file_sha(os.path.join(_tdir(t["id"]), "snapshot.json")),
            "oplog_records": len(ctx.get("oplog") or []), "vision_model": os.environ.get("QA_VISION_MODEL", "claude-opus-5-5"),
            "signoff": {k: v for k, v in (ctx.get("signoff") or {}).items() if k != "items"} | {"items": (ctx.get("signoff") or {}).get("items", [])},
            "requirement_rationale": [{"requirement": r["requirement"], "kinds": r["kinds"], "check": r["check"], "status": r["status"]}
                                      for r in cov.get("requirements", [])]}


def _state_row(t, ctx):
    """STATE-01: the drawing at the end of the run is the drawing the checks read at the start."""
    res = {"id": "STATE-01", "req": "The drawing did not change while the checks ran, so a PASS describes the artifact that exists now.",
           "source": "QA gate design (release integrity: the artifact released must be the artifact verified)", "blocking": True,
           "expected": "identical drawing fingerprint before and after the run", "evidence": [], "fix": ""}
    if not ctx.get("state_fp"):
        res.update(status="UNCERTAIN", observed="drawing state could not be fingerprinted",
                   fix="Open MicroStation with the bridge running and re-run.")
        return res
    try:
        fp, model, _ = state_fingerprint(_ops())
    except Exception as ex:
        res.update(status="UNCERTAIN", observed=f"drawing state could not be re-read: {ex}", fix="Re-run with MicroStation open.")
        return res
    if fp != ctx["state_fp"]:
        res.update(status="UNCERTAIN", observed="the drawing changed while the gate was running",
                   fix="Do not edit during a run. Re-run `qa_gate run`.")
        return res
    note = " (vertex hashes capped; very large model)" if ctx.get("state_vh_capped") else ""
    res.update(status="PASS", observed=f"drawing state {fp[:12]} unchanged during the run{note}")
    return res


STALL_S = 240  # a run still going after this long dumps every thread's stack to <task>/stall-trace.txt (diagnoses silent stalls)


def cmd_run(a):
    import faulthandler
    t = load_task(a.task)
    ctx: dict = {"stages": {}}
    trace = open(os.path.join(_tdir(t["id"]), "stall-trace.txt"), "w")
    faulthandler.dump_traceback_later(STALL_S, repeat=True, file=trace)
    try:
        return _run_impl(a, t, ctx)
    finally:
        faulthandler.cancel_dump_traceback_later()
        trace.close()


def _run_impl(a, t, ctx):
    t0 = time.time()
    _gather(t, a, ctx)
    ctx["stages"]["gather (read-back, render, capture)"] = time.time() - t0
    if ctx.get("model_mismatch"):
        # 2026-10-25: a run on the other model checked the wrong elements and still printed a mostly green report
        print(f"{t['id']}: BLOCKED - {ctx['model_mismatch']}. Nothing was checked. Activate the task's model "
              "(switching models inside the DGN is allowed) and run again.")
        return 1
    ctx["signoff"] = load_signoff(t, ctx.get("state_fp"))
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
        if r["check"] == "gate.state":
            continue  # appended after every other check (_state_row)
        mod, fn = r["check"].split(".")
        t0 = time.time()
        try:
            res = getattr(importlib.import_module(f"qa.checks.{mod}"), fn)(t, ctx)
        except Exception as ex:
            res = {"status": "UNCERTAIN", "observed": f"checker error: {ex}", "expected": "", "evidence": [], "fix": ""}
        ctx["stages"][r["id"]] = time.time() - t0
        res.update(id=r["id"], req=r["req"], source=r["source"], blocking=r.get("blocking", True))
        results.append(res)
        ctx["results"][r["id"]] = res
    if vis_row:
        from qa.checks import visual
        t0 = time.time()
        if ctx.get("images") and not a.no_vision:
            vr = visual.vision_review(ctx["images"], t, failure_modes_for(t["type"]),
                                      [r for r in rows if r["id"].startswith("CAD-")], det_results=ctx["results"])
            vr["t"] = time.time()
            json.dump(vr, open(os.path.join(_tdir(t["id"]), "vision-review.json"), "w"), indent=1)
        ctx["stages"]["vision call"] = time.time() - t0
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
        res.update(id=r["id"], req=r["req"], source=r["source"], blocking=r.get("blocking", True))
        results.append(res)
    if t["type"] == "cad.drafting":
        t0 = time.time()
        results.append(_state_row(t, ctx))  # last: compares the drawing now with the drawing the checks read
        ctx["stages"]["STATE-01 re-read"] = time.time() - t0
    blocking = [r for r in results if r["blocking"] and r["status"] in BLOCKING]
    overall = "PASS" if not blocking else ("FAIL" if any(r["status"] == "FAIL" for r in blocking) else "UNCERTAIN")
    if overall == "PASS" and ctx.get("state_fp"):
        t["verified_fp"], t["verified_model"] = ctx["state_fp"], ctx.get("run_model")
    else:
        t.pop("verified_fp", None)
        t.pop("verified_model", None)
    t["runs"] += 1
    t["fail_streak"] = 0 if overall == "PASS" else t.get("fail_streak", 0) + 1
    t["last_status"] = overall
    t["last_run"] = time.time()
    save_task(t)
    rec = {"task": t["id"], "type": t["type"], "run": t["runs"], "at": datetime.now().isoformat(timespec="seconds"),
           "stage_seconds": {k: round(v, 2) for k, v in ctx.get("stages", {}).items()},
           "overall": overall, "images": ctx.get("images"), "results": results,
           "n_elements": len(ctx.get("elements") or []), "coverage": ctx.get("coverage"),
           "needs_signoff": ctx.get("needs_signoff", []), "provenance": _provenance(t, ctx)}
    with open(os.path.join(_tdir(t["id"]), "latest.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=1)
    with open(os.path.join(_tdir(t["id"]), f"record-{t['runs']}.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=1)
    print(f"QA {t['id']} run {t['runs']}: {overall}")
    for r in results:
        mark = {"PASS": "ok ", "FAIL": "XX ", "UNCERTAIN": "?? ", "WARN": "!  ", "NA": "-  "}.get(r["status"], "   ")
        print(f"  {mark}{r['id']} {r['status']}: {r['observed']}")
        if r.get("signed_off_from"):
            print(f"       (was {r['signed_off_from']['status']} before the engineer's sign-off: {r['signed_off_from']['observed'][:100]})")
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
        ev = [x.strip() for x in (a.evidence or "").split(",") if x.strip()]
        ar.setdefault("disputes", {})[str(n)] = {"reason": reason, "evidence": ev}
    json.dump(ar, open(p, "w"), indent=1)
    print(f"{t['id']}: review recorded {json.dumps({k: v for k, v in ar.items() if k != 't'})[:300]}")
    print("Run `qa_gate run` to re-evaluate.")
    return 0


def cmd_close(a):
    t = load_task(a.task)
    if t.get("last_status") != "PASS":
        print(f"{t['id']}: cannot close - last run is {t.get('last_status') or 'none'}. Run `qa_gate run` until PASS.")
        return 1
    if t.get("verified_fp"):  # the drawing being released must be the drawing that was verified
        try:
            fp, model, _ = state_fingerprint(_ops())
        except Exception as ex:
            print(f"{t['id']}: BLOCKED - cannot re-read the drawing to confirm it is unchanged ({ex}). Open MicroStation and retry.")
            return 1
        if model != t.get("verified_model"):
            print(f"{t['id']}: BLOCKED - the active model is {model}; the PASS was for {t.get('verified_model')}. Activate that model, then close.")
            return 1
        if fp != t["verified_fp"]:
            print(f"{t['id']}: BLOCKED - the drawing changed after verification (state {str(fp)[:12]} vs verified {t['verified_fp'][:12]}). "
                  "Re-run `qa_gate run` so the changed elements are checked.")
            return 1
    from qa import claims as qclaims
    openc = qclaims.open_claims(t)
    if openc:
        print(f"{t['id']}: cannot close - {len(openc)} claim(s) are still hypothesis/unknown/contradicted. Test them and upgrade "
              "(`qa_gate claim --id N --level verified --evidence <passed check>`) or withdraw (`--withdraw N`):")
        for c in openc:
            print(f"  #{c['id']} {c['level'].upper()}: {c['text']}")
        return 1
    if t.get("mode") == "development" and str(t.get("scratch", "")).lower() not in ("", "none"):
        try:
            active = str(_ops().describe_drawing_state().get("fileName", ""))
        except Exception as ex:
            print(f"{t['id']}: BLOCKED - cannot confirm the active file is the scratch file ({ex}).")
            return 1
        if active.lower() != t["scratch"].lower():
            print(f"{t['id']}: BLOCKED - the active file is {active!r}; this development task was on scratch file {t['scratch']!r}.")
            return 1
    t["status"] = "closed"
    t["closed_iso"] = datetime.now().isoformat(timespec="seconds")
    save_task(t)
    _remember(t)
    print(f"{t['id']}: closed (PASS, {t['runs']} runs)")
    inferred = qclaims.inferred_claims(t)
    if inferred:
        print("Report these to the engineer as INFERRED or SUPPORTED (not verified):" + " | ".join(f"#{c['id']} {c['text']}" for c in inferred))
    return 0


def _remember(t):
    """Write the ended task into experience memory once (Data/experience/experience.jsonl). Best effort."""
    try:
        import experience
        row = experience.from_task(t)
        if not row or len(row["symptom"].strip()) < 15:
            return
        if any(r.get("task") == t["id"] and r.get("status") == row["status"] for r in experience.rows()):
            return
        experience.add(source="auto", **row)
        print(f"{t['id']}: recorded in experience memory as {row['status'].upper()}")
    except Exception as ex:
        print(f"(experience memory not written: {ex})")


def cmd_rung(a):
    import ladders
    t = load_task(a.task)
    if not t.get("ladder"):
        raise SystemExit("this task has no diagnostic ladder (none matched its description)")
    if a.id is None:
        print(ladders.summary(ladders.get(t["ladder"]["id"])))
        for rid, s in ladders.state(t).items():
            print(f"  {rid}: {s['status']} - {s['evidence'][:100]}")
        left = ladders.gaps(t)
        print("Still open: " + ("; ".join(left) if left else "none"))
        return 0
    try:
        ladders.record(t, a.id, a.status or "", a.evidence or "")
    except ValueError as ex:
        raise SystemExit(f"rung refused: {ex}")
    save_task(t)
    left = ladders.gaps(t)
    print(f"{t['id']}: {a.id} recorded as {a.status}. " + ("Still open: " + "; ".join(left) if left else "All cheaper rungs have a result."))
    return 0


def cmd_experience(a):
    import experience
    if a.action == "search":
        print(experience.for_task(a.text or "", k=5) or "no similar past experience")
        return 0
    tried = []
    for x in a.tried or []:
        what, _, result = x.partition("=>")
        tried.append({"what": what.strip(), "result": result.strip()})
    findings = []
    for f in a.finding or []:
        lvl, _, txt = f.partition(":")
        findings.append({"level": lvl.strip(), "text": txt.strip()})
    try:
        r = experience.add(a.symptom or "", a.status or "", tried, findings, a.context or "", a.task or "")
    except ValueError as ex:
        raise SystemExit(f"experience refused: {ex}")
    print(f"experience #{r['id']} recorded ({r['status']})")
    return 0


def cmd_escalate(a):
    t = load_task(a.task)
    t["status"] = "escalated"
    t["escalation"] = a.reason
    save_task(t)
    _remember(t)
    print(f"{t['id']}: ESCALATED - {a.reason}\nTell the engineer exactly this, with the latest record: "
          f"{os.path.join(_tdir(t['id']), 'latest.json')}")
    return 0


def cmd_dispose(a):
    t = load_task(a.task)
    if a.applies not in ("yes", "no"):
        raise SystemExit("--applies must be yes or no")
    if len((a.reason or "").strip()) < 10:
        raise SystemExit("--reason must say why (>= 10 characters)")
    pool = list(t.get("candidates", [])) + list(t.get("sources", []))
    hits = [c for c in pool if a.cite == "all" or a.cite.lower() in c["cite"].lower()]
    if not hits:
        raise SystemExit("no candidate matches --cite; see the list printed by `qa_gate start`")
    if a.applies == "yes":  # "applies" must rest on having read the source, not on its title
        import retrieval
        unread = [c["cite"] for c in hits if not retrieval.has_read(*retrieval.kind_key(c))]
        if unread:
            raise SystemExit("applies=yes needs the source read first (get_tutorial_playbook / get_standard_section / "
                             "read_standard_pages / read_tutorial). Not read: " + " | ".join(unread[:6]))
    for c in hits:
        c["applies"], c["reason"] = a.applies, a.reason
    save_task(t)
    left = sum(1 for c in pool if c.get("applies") not in ("yes", "no"))
    print(f"{t['id']}: {len(hits)} candidate(s) marked applies={a.applies}; {left} still open")
    return 0


def cmd_claim(a):
    from qa import claims as qclaims
    t = load_task(a.task)
    if a.withdraw is not None:
        try:
            qclaims.withdraw(t, a.withdraw)
        except ValueError as ex:
            raise SystemExit(str(ex))
        save_task(t)
        print(f"{t['id']}: claim {a.withdraw} withdrawn")
        return 0
    latest = os.path.join(_tdir(t["id"]), "latest.json")
    passed = qclaims.passed_checks(json.load(open(latest, encoding="utf-8")) if os.path.exists(latest) else None)
    try:
        c = qclaims.add(t, a.level, a.text or "", a.evidence or "", passed, a.id)
    except ValueError as ex:
        raise SystemExit(f"claim refused: {ex}")
    save_task(t)
    print(f"{t['id']}: claim #{c['id']} recorded as {c['level'].upper()}: {c['text']}")
    return 0


def cmd_replan(a):
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import task_guard
    t = load_task(a.task)
    if len((a.note or "").strip()) < 30:
        raise SystemExit("--note must say what failed, what the source says and the different approach (>= 30 characters)")
    fail_at = task_guard.last_failure_at(t["id"])
    rows = []
    rec = os.path.join(ROOT, "outputs", "qa", "receipts.jsonl")
    if os.path.exists(rec):
        for ln in open(rec, encoding="utf-8"):
            try:
                rows.append(json.loads(ln))
            except Exception:
                pass
    if not any(r.get("at", 0) > fail_at for r in rows):
        raise SystemExit("replan refused: nothing was read or searched since the last failure. Read the playbook step, the standard "
                         "or the tutorial source (get_tutorial_playbook / search_tutorials / read_tutorial / search_standards), "
                         "then replan. If the step is a GUI step (who: engineer), `qa_gate escalate` instead.")
    task_guard.record_replan(t["id"], a.note)
    print(f"{t['id']}: replan recorded; the failed-attempt count is reset. Use a DIFFERENT approach, not another variant of the failed one.")
    return 0


def cmd_reopen(a):
    t = load_task(a.task) if a.task else None
    if t is None:
        raise SystemExit("--task ID required")
    if t["status"] not in ("blocked", "expired"):
        print(f"{t['id']}: status is {t['status']}; only blocked or expired tasks are reopened")
        return 1
    t["status"], t["reopened"] = "in_progress", time.time()
    t.pop("_hook_blocks", None)
    t.pop("_hook_last_run", None)
    t.pop("state_reason", None)
    save_task(t)
    print(f"{t['id']}: reopened. Its last result was {t.get('last_status')}; edit, then `qa_gate run`.")
    return 0


def cmd_status(a):
    ts = [load_task(a.task)] if a.task else sorted(all_tasks(), key=lambda t: t["created"])[-15:]
    for t in ts:
        print(f"{t['id']}  {t['type']}  {t.get('mode') or 'production'}  {t['status']}  last={t.get('last_status')}  runs={t.get('runs')}  {t['desc'][:60]}")
        for rid, rs in (t.get("ladder") or {}).get("rungs", {}).items():
            print(f"    rung {rid}: {rs['status']} - {rs['evidence'][:80]}")
        for c in t.get("claims", []):
            print(f"    claim #{c['id']} {c['level'].upper()}: {c['text'][:90]}")
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
    for stream in (sys.stdout, sys.stderr):  # 2026-10-25: a "≈" in a check message crashed `run` on a cp1252 console after the checks ran
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("start"); s.add_argument("--type", required=True); s.add_argument("--desc", required=True)
    s.add_argument("--files", nargs="*"); s.add_argument("--extent"); s.add_argument("--ids"); s.add_argument("--id")
    s.add_argument("--text-heights"); s.add_argument("--is3d", action="store_true"); s.add_argument("--no-standards", action="store_true")
    s.add_argument("--may-modify", action="store_true", help="engineer asked to edit/move/delete EXISTING elements")
    s.add_argument("--mode", choices=["production", "development"], default="production",
                   help="production (default): established capabilities only. development: build/investigate tooling on a scratch file")
    s.add_argument("--scratch", help="development: the scratch DGN file name that is open, or 'none' for code-only work")
    s.add_argument("--mode-reason", help="development: why this is capability development")
    dp = sub.add_parser("dispose"); dp.add_argument("--task"); dp.add_argument("--cite", required=True)
    dp.add_argument("--applies", required=True); dp.add_argument("--reason", required=True)
    e = sub.add_parser("expect"); e.add_argument("--task"); e.add_argument("--id", required=True); e.add_argument("--desc", default="")
    e.add_argument("--where", nargs="*"); e.add_argument("--count"); e.add_argument("--value"); e.add_argument("--target")
    e.add_argument("--rel", help="offset|length|angle: geometric relation measured from saved geometry")
    e.add_argument("--a"); e.add_argument("--b"); e.add_argument("--tol")
    e.add_argument("--params", help="JSON merged into a --rel row: min/max/ve/x/side/target (distance, side, vclear, offset_curve)")
    c = sub.add_parser("cite"); c.add_argument("--task"); c.add_argument("--doc", required=True); c.add_argument("--page", type=int, required=True)
    c.add_argument("--quote", required=True); c.add_argument("--req")
    k = sub.add_parser("calc"); k.add_argument("--task"); k.add_argument("--id", required=True); k.add_argument("--fn", required=True)
    k.add_argument("--inputs", required=True); k.add_argument("--result", type=float, required=True); k.add_argument("--units", required=True)
    k.add_argument("--criterion", required=True)
    r = sub.add_parser("run"); r.add_argument("--task"); r.add_argument("--image", nargs="*"); r.add_argument("--ids")
    r.add_argument("--no-vision", action="store_true")
    v = sub.add_parser("review"); v.add_argument("--task"); v.add_argument("--ok", action="store_true"); v.add_argument("--checked")
    v.add_argument("--issue"); v.add_argument("--dispute", nargs=2, metavar=("N", "REASON"))
    v.add_argument("--evidence", help="with --dispute: deterministic check ids that PASSED and disprove the issue (CAD-LVL-01,...)")
    so = sub.add_parser("signoff", help="engineer only (interactive terminal)"); so.add_argument("--task"); so.add_argument("--items")
    so.add_argument("--reason")
    ro = sub.add_parser("reopen", help="resume a blocked or expired task"); ro.add_argument("--task", required=True)
    cm = sub.add_parser("claim", help="record an evidence-labelled claim: observed|inferred|supported|hypothesis|verified|unknown|contradicted")
    cm.add_argument("--task"); cm.add_argument("--level"); cm.add_argument("--text"); cm.add_argument("--evidence")
    cm.add_argument("--id", type=int, help="upgrade an existing claim"); cm.add_argument("--withdraw", type=int)
    rp = sub.add_parser("replan", help="after repeated failed edits: record the new approach (needs a read since the failure)")
    rp.add_argument("--task"); rp.add_argument("--note", required=True)
    cl = sub.add_parser("close"); cl.add_argument("--task")
    es = sub.add_parser("escalate"); es.add_argument("--task"); es.add_argument("--reason", required=True)
    st = sub.add_parser("status"); st.add_argument("--task")
    ru = sub.add_parser("rung", help="diagnostic ladder: record a rung result (no --id lists the ladder)")
    ru.add_argument("--task"); ru.add_argument("--id"); ru.add_argument("--status"); ru.add_argument("--evidence")
    ex = sub.add_parser("experience", help="experience memory: add|search")
    ex.add_argument("action", choices=["add", "search"]); ex.add_argument("--text"); ex.add_argument("--symptom")
    ex.add_argument("--status"); ex.add_argument("--context"); ex.add_argument("--task")
    ex.add_argument("--tried", action="append", help="'what => result' (repeat)")
    ex.add_argument("--finding", action="append", help="'observed|reported|inferred: text' (repeat)")
    a = ap.parse_args()
    return {"start": cmd_start, "expect": cmd_expect, "cite": cmd_cite, "calc": cmd_calc, "run": cmd_run,
            "review": cmd_review, "close": cmd_close, "escalate": cmd_escalate, "status": cmd_status,
            "dispose": cmd_dispose, "signoff": cmd_signoff, "reopen": cmd_reopen, "claim": cmd_claim, "replan": cmd_replan,
            "rung": cmd_rung, "experience": cmd_experience}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
