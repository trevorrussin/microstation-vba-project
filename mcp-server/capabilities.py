"""Capability catalog: one place to ask "can this be done, and which way?" before choosing a tool.

The catalog is built from what already exists, not a new registry to maintain:

  workflow   Data/workflow-contracts/*.yaml   multi-step workflows with inputs, guarded ops and completion requirements
  procedure  Data/tutorial-playbooks/*.md     the engineers' procedures (steps marked who: agent / who: engineer)
  tool       mcp-server/wztc_ops.py           public bridge functions (first docstring paragraph)
  keyin      Data/command-registry.tsv        key-in recipes with their safety status
  diagnosis  Data/diagnostic-ladders/*.yaml   ordered checks for a known symptom
  gui        Data/ord-ui-catalog.yaml         dialogs probed so far: what UI Automation can see, who operates them

`find(text, mode)` ranks entries by IDF-weighted word overlap (title and name count more than body text) and returns either
matches, each with its status, the execution method from Data/capabilities/policy.yaml, whether production mode may use it, how it
is verified and the next call, or UNSUPPORTED with what to do in that mode. A deterministic table picks the method; no model does.
There is no success-rate data in the catalog because none has been measured yet (qa/trajectory.py will supply it once tasks log edits).
"""
from __future__ import annotations

import ast
import csv
import glob
import math
import os
import re
from functools import lru_cache

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MIN_COVERAGE = 0.5   # an entry must contain at least this share of the request's distinct words
MIN_SCORE = 6.0      # and this much IDF-weighted overlap
_STOP = set("the and for with that this from have has not are was were but all any into onto over under when then than they them their "
            "will can could should would about after before while which what where there here make create place add please need want "
            "use using show get set run".split())

REGISTRY_STATUS = {"verified-headless-safe": "supported", "interactive-only-use-handoff": "engineer",
                   "external-app-blocked": "blocked", "unsafe-blocked": "blocked", "needs-testing": "unverified"}


def _tok(text: str) -> list[str]:
    out = []
    for w in re.findall(r"[a-z0-9]{3,}", (text or "").lower().replace("_", " ")):
        w = w[:-1] if w.endswith("s") and len(w) > 4 else w
        if w not in _STOP:
            out.append(w)
    return out


def _yaml(path: str):
    import yaml
    return yaml.safe_load(open(path, encoding="utf-8"))


def _policy() -> dict:
    return _yaml(os.path.join(ROOT, "Data", "capabilities", "policy.yaml"))


def _entry(cid, kind, title, name_text, body_text, status, verification="", next_call="", extra=None) -> dict:
    e = {"id": cid, "kind": kind, "title": title, "status": status, "verification": verification, "next": next_call,
         "_name": _tok(name_text), "_body": _tok(body_text)}
    e.update(extra or {})
    return e


def _workflows() -> list[dict]:
    out = []
    for p in sorted(glob.glob(os.path.join(ROOT, "Data", "workflow-contracts", "*.yaml"))):
        try:
            c = _yaml(p)
        except Exception:
            continue
        manual = [s for s in c.get("manual_steps", []) if s.get("who") == "engineer"]
        agent = c.get("agent_steps", [])
        status = "partial" if manual else "supported"
        body = " ".join([str(c.get("deliverable", ""))] + [str(s.get("what", "")) for s in manual] + [str(s.get("note", "")) for s in agent])
        out.append(_entry(f"workflow:{c['workflow_id']}", "workflow", c.get("title", c["workflow_id"]), c["workflow_id"] + " " + c.get("title", ""),
                          body, status, "; ".join(str(x) for x in c.get("completion_requirements", []))[:400],
                          f"python scripts/workflow_preflight.py {c['workflow_id']}",
                          {"engineer_steps": [s.get("what") for s in manual], "agent_steps": [s.get("op") for s in agent if s.get("enabled")]}))
    return out


def _procedures() -> list[dict]:
    out = []
    for p in sorted(glob.glob(os.path.join(ROOT, "Data", "tutorial-playbooks", "*.md"))):
        src = open(p, encoding="utf-8").read()
        topic = os.path.basename(p)[:-3]
        title = (re.search(r"^title:\s*(.+)$", src, re.M) or [None, topic])[1].strip()
        when = (re.search(r"^when:\s*(.+)$", src, re.M) or [None, ""])[1].strip()
        kws = (re.search(r"^keywords:\s*(.+)$", src, re.M) or [None, ""])[1]
        who = re.findall(r"^\s+who:\s*(agent|engineer)", src, re.M)
        status = "supported" if who and all(w == "agent" for w in who) else ("engineer" if who and all(w == "engineer" for w in who) else "partial")
        out.append(_entry(f"procedure:{topic}", "procedure", title, topic + " " + title, when + " " + kws, status,
                          "playbook steps carry tests (who: agent) and engineer-only reasons (who: engineer)", f"get_tutorial_playbook {topic}",
                          {"agent_steps_n": who.count("agent"), "engineer_steps_n": who.count("engineer")}))
    return out


def _tools() -> list[dict]:
    path = os.path.join(ROOT, "mcp-server", "wztc_ops.py")
    tree = ast.parse(open(path, encoding="utf-8").read())
    out = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and not n.name.startswith("_"):
            doc = ast.get_docstring(n) or ""
            if not doc:
                continue
            first = doc.strip().split("\n\n")[0].replace("\n", " ")
            out.append(_entry(f"tool:{n.name}", "tool", first[:110], n.name, first, "supported",
                              "recorded by op_recorder; OP-01 re-measures edits", f"call {n.name} (an open QA task is required for edits)"))
    return out


def _keyins() -> list[dict]:
    p = os.path.join(ROOT, "Data", "command-registry.tsv")
    out = []
    for r in csv.DictReader(open(p, encoding="utf-8"), delimiter="\t"):
        if r.get("category") == "direct_api":
            continue  # run_registry_command refuses these; the dedicated tool is listed under kind "tool"
        st = REGISTRY_STATUS.get(r["safetyStatus"], "unverified")
        out.append(_entry(f"keyin:{r['opName']}", "keyin", r["opName"], r["opName"], r.get("notes", "") + " " + r.get("category", ""), st,
                          "settings-only or recorded op; verify by read-back", f"describe_registry_command {r['opName']}",
                          {"safetyStatus": r["safetyStatus"]}))
    return out


def _gui() -> list[dict]:
    """Dialogs and windows probed so far (Data/ord-ui-catalog.yaml): what UI Automation can and cannot see, and who operates them."""
    p = os.path.join(ROOT, "Data", "ord-ui-catalog.yaml")
    if not os.path.exists(p):
        return []
    out = []
    for t in _yaml(p).get("targets", []):
        if not str(t.get("id", "")).startswith("dialog."):
            continue
        body = " ".join(str(t.get(k, "")) for k in ("notes", "evidence", "window_name", "access")) + " " + " ".join(map(str, t.get("try_without_dialog", [])))
        status = "engineer" if t.get("access") == "engineer_gui" or t.get("uia") in ("chrome_only", "empty") else "unverified"
        out.append(_entry(f"gui:{t['id']}", "gui", f"{t.get('window_name') or t['id']} (UIA: {t.get('uia', 'unknown')})", t["id"] + " " + str(t.get("window_name", "")),
                          body, status, "read back the state after any change; do not trust OK alone",
                          "hand to the engineer with the exact instruction" + (f"; untried without the dialog: {'; '.join(map(str, t['try_without_dialog']))}" if t.get("try_without_dialog") else ""),
                          {"uia": t.get("uia")}))
    return out


def _diagnoses() -> list[dict]:
    out = []
    for p in sorted(glob.glob(os.path.join(ROOT, "Data", "diagnostic-ladders", "*.yaml"))):
        try:
            l = _yaml(p)["ladder"]
        except Exception:
            continue
        out.append(_entry(f"diagnosis:{l['id']}", "diagnosis", l["title"], l["id"] + " " + l["title"], " ".join(map(str, l.get("keywords", []))) + " " + str(l.get("why", "")),
                          "partial", "each rung records a result (qa_gate rung); edits wait for the cheap rungs",
                          "python scripts/qa_gate.py start --desc \"<symptom>\" (the ladder attaches)"))
    return out


@lru_cache(maxsize=1)
def catalog() -> tuple:
    entries = _workflows() + _procedures() + _tools() + _keyins() + _diagnoses() + _gui()
    df: dict[str, int] = {}
    for e in entries:
        for t in set(e["_name"]) | set(e["_body"]):
            df[t] = df.get(t, 0) + 1
    return tuple(entries), df


def reset() -> None:
    catalog.cache_clear()


def _score(q: list[str], e: dict, df: dict, n: int) -> tuple[float, int]:
    name, body = set(e["_name"]), set(e["_body"])
    s, hits = 0.0, 0
    for t in set(q):
        if t in name or t in body:
            hits += 1
            s += math.log(1 + n / df.get(t, 1)) * (3.0 if t in name else 1.0)
    return s, hits


def executors_for(e: dict, pol: dict) -> tuple[list[str], str]:
    ov = (pol.get("overrides") or {}).get(e["id"])
    if ov:
        return ov["executors"], ov.get("why", "")
    k = (pol.get("kinds") or {}).get(e["kind"], {})
    return k.get("executors", []), k.get("why", "")


def production_ok(e: dict, pol: dict) -> bool:
    kind_ok = bool((pol.get("kinds") or {}).get(e["kind"], {}).get("production"))
    return kind_ok and e["status"] in ("supported", "partial", "engineer")


def find(text: str, k: int = 5, mode: str = "production") -> dict:
    entries, df = catalog()
    pol = _policy()
    q = _tok(text)
    scored = []
    for e in entries:
        s, hits = _score(q, e, df, len(entries))
        if hits:
            scored.append((s, hits, e))
    scored.sort(key=lambda x: -x[0])

    qset = sorted(set(q))

    def view(s, e):
        ex, why = executors_for(e, pol)
        out = {k2: v for k2, v in e.items() if not k2.startswith("_")}
        words = [t for t in qset if t in e["_name"] or t in e["_body"]]
        out.update(score=round(s, 1), executors=ex, why=why, production_ok=production_ok(e, pol),
                   matched_words=words, coverage=f"{len(words)} of {len(qset)} request words",
                   confidence="high" if len(words) >= 0.75 * len(qset) else "low")
        return out

    need = max(1, math.ceil(MIN_COVERAGE * len(set(q))))  # lexical match cannot tell an off-topic word from a topical one, so most of the request must match
    good = [(s, e) for s, h, e in scored if h >= need and s >= MIN_SCORE]
    res = {"query": text, "mode": mode}
    if good:
        res.update(status="FOUND", matches=[view(s, e) for s, e in good[:k]])
        if mode == "production" and not any(m["production_ok"] for m in res["matches"]):
            res["note"] = "Every match is blocked or unverified for production mode."
    else:
        res.update(status="UNSUPPORTED", matches=[], nearest=[view(s, e) for s, h, e in scored[:3]],
                   guidance=(pol.get("unsupported") or {}).get(mode, ""))
    return res


def format_result(r: dict) -> str:
    lines = []
    if r["status"] == "FOUND":
        lines.append(f"CAPABILITIES for this request ({r['mode']} mode):")
        for m in r["matches"][:4]:
            tag = "ok in production" if m["production_ok"] else "NOT for production"
            lines.append(f"  {m['id']} [{m['status']}, {tag}] {m['title'][:90]}")
            if m["confidence"] == "low":
                lines.append(f"      LOW CONFIDENCE: matched {m['coverage']} ({', '.join(m['matched_words'])}). Read it before relying on it.")
            lines.append(f"      method: {' -> '.join(m['executors'])}")
            lines.append(f"      next: {m['next']}")
        if r.get("note"):
            lines.append("  " + r["note"])
    else:
        lines.append(f"NO SUPPORTED CAPABILITY found for this request ({r['mode']} mode). {r['guidance']}")
        for m in r.get("nearest", []):
            lines.append(f"  nearest (weak): {m['id']} [{m['status']}] {m['title'][:80]}")
    return "\n".join(lines)
