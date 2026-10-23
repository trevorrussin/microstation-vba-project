"""Tutorial playbooks must be traceable and complete.

- every `cite` / `sources` entry names a real tutorial file and a timestamp / slide / page inside it
- every agent step has an executable test; engineer steps say why the agent must not run them
- every preflight check type is one the runner implements
- every tutorial source under Standards/Tutorials is covered by at least one playbook
"""
from __future__ import annotations

import glob
import os
import re
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

yaml = pytest.importorskip("yaml")
SRC = os.path.join(ROOT, "Standards", "Tutorials")
DERIVED = os.path.join(ROOT, "Standards", "derived", "tutorials")
PB_DIR = os.path.join(ROOT, "Data", "tutorial-playbooks")

pytestmark = pytest.mark.skipif(not os.path.isdir(DERIVED), reason="tutorials not ingested")


def _playbooks():
    out = []
    for p in sorted(glob.glob(os.path.join(PB_DIR, "*.md"))):
        m = re.search(r"```yaml playbook\n(.*?)\n```", open(p, encoding="utf-8").read(), re.S)
        assert m, f"{p}: no yaml playbook block"
        out.append((os.path.basename(p), yaml.safe_load(m.group(1))["playbook"]))
    return out


def _norm(s):
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def _tutorial_files():
    return [p for p in glob.glob(os.path.join(SRC, "**", "*"), recursive=True) if p.lower().endswith((".mp4", ".pdf", ".pptx", ".xlsx"))]


def _resolve(cite: str):
    """'<file> @ mm:ss' | '<file> slide N' | '<file> page N' -> (path, kind, value)"""
    m = re.match(r"(.+?)\s+(?:@\s*([\d:]+)|slide\s+(\d+)|page\s+(\d+)|sheet\s+(\S+))\s*$", cite.strip())
    assert m, f"cannot parse cite {cite!r}"
    name = _norm(m.group(1))
    hits = [p for p in _tutorial_files() if _norm(os.path.basename(p)) == name]
    assert hits, f"cite names no tutorial file: {cite!r}"
    return hits[0], m


def _digest_dir(path):
    base = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.relpath(path, SRC))[0]).strip("_")[:120]
    return os.path.join(DERIVED, base)


def _seconds(ts):
    p = [int(x) for x in ts.split(":")]
    return p[0] * 3600 + p[1] * 60 + p[2] if len(p) == 3 else p[0] * 60 + p[1]


def _check_cite(cite):
    path, m = _resolve(cite)
    d = _digest_dir(path)
    if m.group(2):  # video timestamp
        tr = os.path.join(d, "transcript.md")
        assert os.path.exists(tr), f"{cite}: video not ingested yet"
        dur = re.search(r"duration ([\d:]+)", open(tr, encoding="utf-8").readline())
        assert dur and _seconds(m.group(2)) <= _seconds(dur.group(1)) + 2, f"{cite}: timestamp beyond video length {dur.group(1) if dur else '?'}"
    elif m.group(3) or m.group(4):
        txt = open(os.path.join(d, "text.md"), encoding="utf-8").read()
        kind, n = ("slide", m.group(3)) if m.group(3) else ("page", m.group(4))
        assert re.search(rf"^## {kind} {n}$", txt, re.M), f"{cite}: no {kind} {n} in the extracted text"


@pytest.mark.parametrize("name,pb", _playbooks())
def test_citations_resolve(name, pb):
    cites = list(pb.get("sources") or []) + [s["cite"] for s in pb.get("steps") or [] if s.get("cite")]
    assert cites, f"{name}: playbook cites nothing"
    for c in cites:
        _check_cite(c)


@pytest.mark.parametrize("name,pb", _playbooks())
def test_steps_are_executable_or_explained(name, pb):
    import playbook_runner as pr
    ids = set()
    for s in pb["steps"]:
        assert s["id"] not in ids, f"{name}: duplicate step id {s['id']}"
        ids.add(s["id"])
        assert s.get("who") in ("agent", "engineer"), f"{name} {s['id']}: who must be agent|engineer"
        assert s.get("cite"), f"{name} {s['id']}: step without a citation"
        if s["who"] == "agent":
            t = s.get("test")
            assert t, f"{name} {s['id']}: agent step without a test"
            assert isinstance(t.get("name", ""), str), f"{name} {s['id']}: test name must be a string (quote cell names like \"240_1\")"
            assert (t.get("action") or "op") in pr.LIVE, f"{name} {s['id']}: unknown test action {t.get('action')}"
        else:
            assert s.get("manual_reason"), f"{name} {s['id']}: engineer step must say why the agent may not run it"
    for c in pb.get("preflight") or []:
        assert isinstance(c.get("name", ""), str), f"{name} {c['id']}: name must be a string (quote it: YAML reads 240_1 as the number 2401)"
        assert c["check"] in pr.PREFLIGHT, f"{name} {c['id']}: unknown preflight check {c['check']}"


def test_every_tutorial_is_covered_by_a_playbook():
    covered = set()
    for _, pb in _playbooks():
        for c in list(pb.get("sources") or []) + [s["cite"] for s in pb.get("steps") or [] if s.get("cite")] + list(pb.get("also_covers") or []):
            try:
                covered.add(os.path.normcase(_resolve(c)[0]))
            except AssertionError:
                pass
    missing = [os.path.relpath(p, SRC) for p in _tutorial_files() if os.path.normcase(p) not in covered]
    assert not missing, "tutorials not covered by any playbook: " + "; ".join(missing)


def test_registry_playbook_references_point_at_real_steps():
    from qa.checks import coverage
    steps = {(n[:-3], s["id"]) for n, pb in _playbooks() for s in pb["steps"]}
    for kind in coverage.registry()["kinds"]:
        for r in kind["requirements"]:
            if r.get("playbook"):
                topic, _, step = r["playbook"].partition("#")
                assert (topic, step) in steps, f"{r['id']}: playbook reference {r['playbook']} does not exist"
