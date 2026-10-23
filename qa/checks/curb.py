"""AutoCAD curb-classify checks (Screenshots SGD sheets).

CURB-ID-01   identity: the DWG's survey + plan xrefs name the same locations as the sheet file
             (catches Loc10_SGD holding Loc03-04's drawing).
CURB-GATE-01 runs scripts/_acad_verify_location_sheet.py --json (late-bound LF, Depressed>0
             where ramps exist) and maps its result.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def locs_from_name(name: str) -> set[int]:
    """'HWP24QC1_Location_03-04_53-55_SGD.dwg' -> {3,4,53,54,55}; 'T4318AG25712_3-4,53-55' -> same."""
    base = os.path.basename(name)
    base = re.sub(r"\.(dwg|bak)$", "", base, flags=re.I)
    base = re.sub(r"_?SGD$", "", base, flags=re.I)
    m = re.search(r"(?:Location_|T4318A_?G25712_)(.+)$", base, re.I)
    if not m:
        return set()
    out: set[int] = set()
    for part in re.split(r"[_,&]", m.group(1)):
        r = re.fullmatch(r"(\d{1,3})-(\d{1,3})", part)
        if r:
            a, b = int(r.group(1)), int(r.group(2))
            if 0 < b - a < 12:
                out.update(range(a, b + 1))
            continue
        if re.fullmatch(r"\d{1,3}", part):
            out.add(int(part))
    return out


def _verify_json(dwg: str) -> dict:
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "_acad_verify_location_sheet.py"),
                        "--file", dwg, "--json"], capture_output=True, text=True, timeout=1800, env=env)
    try:
        return json.loads(p.stdout[p.stdout.find("{"):])
    except Exception:
        return {"status": "ESCALATED", "fails": [f"verify output unreadable: {p.stdout[-300:]} {p.stderr[-300:]}"]}


def _dwg(task):
    files = [f for f in task.get("files") or [] if f.lower().endswith(".dwg")]
    return os.path.join(ROOT, files[0]) if files else None


def identity(task, ctx):
    dwg = _dwg(task)
    if not dwg:
        return {"status": "UNCERTAIN", "observed": "no .dwg in task files", "expected": "sheet DWG", "evidence": [], "fix": ""}
    v = ctx.setdefault("_curb_verify", _verify_json(dwg))
    sheet = locs_from_name(dwg)
    xr = v.get("xrefs") or []
    if not xr:
        return {"status": "UNCERTAIN", "observed": "no xrefs read", "expected": "survey + plan xrefs",
                "evidence": v.get("fails", []), "fix": "Check AutoCAD is running and the DWG opens."}
    survey = [(n, p) for n, p in ((x["name"], x["path"]) for x in xr) if "T4318" in (n + p).upper() or n.upper().startswith("SURVEY")]
    plans = [(x["name"], x["path"]) for x in v.get("plan_xrefs") or []]
    ev, bad = [f"sheet locations {sorted(sheet)}"], []
    for n, p in survey:
        s = locs_from_name(p) or locs_from_name(n)
        ev.append(f"survey {n} -> {os.path.basename(p)} locs {sorted(s)}")
        if s and not (s & sheet):
            bad.append(f"survey xref {os.path.basename(p)} is for locations {sorted(s)}, not {sorted(sheet)}")
    for n, p in plans:
        s = locs_from_name(n) or locs_from_name(p)
        ev.append(f"plan {n} locs {sorted(s)}")
        if s and not (s & sheet):
            bad.append(f"plan xref {n} is for locations {sorted(s)}")
    if bad:
        return {"status": "FAIL", "observed": "; ".join(bad), "expected": f"xrefs for locations {sorted(sheet)}",
                "evidence": ev, "fix": "This file holds another sheet's drawing - get the correct SGD from the engineer. Do not classify."}
    if not survey:
        return {"status": "UNCERTAIN", "observed": "no survey xref", "expected": "T4318*/Survey xref", "evidence": ev, "fix": ""}
    return {"status": "PASS", "observed": "xrefs match sheet locations", "expected": "", "evidence": ev, "fix": ""}


def verify_gate(task, ctx):
    dwg = _dwg(task)
    if not dwg:
        return {"status": "UNCERTAIN", "observed": "no .dwg", "expected": "", "evidence": [], "fix": ""}
    v = ctx.setdefault("_curb_verify", _verify_json(dwg))
    ctx.setdefault("results", {}).update({f"steel_{k}": val for k, val in (v.get("steel_lf") or {}).items()})
    if v.get("status") == "PASSED":
        return {"status": "PASS", "observed": f"steel {v.get('steel_lf')}", "expected": "", "evidence": [], "fix": ""}
    st = "UNCERTAIN" if v.get("status") == "ESCALATED" else "FAIL"
    return {"status": st, "observed": "; ".join(v.get("fails") or []), "expected": "verify PASSED",
            "evidence": [f"steel {v.get('steel_lf')}", f"rampish {v.get('rampish_count')}"],
            "fix": "Re-run the classifier for this sheet (batch --only) and fix the failing gate."}
