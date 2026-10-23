"""Standards retrieval + QA gate tools shared by server.py (MCP) and chat_driver.py.

Standards tools read the structure-aware index in Standards/index/ (see standards_index.py):
search_standards (hybrid BM25 + vectors + rerank + cross-refs), get_standard_section,
read_standard_pages, standards_toc, follow_standard_ref, list_standards.
"""
from __future__ import annotations

import json
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_ROOT, "scripts"))
sys.path.insert(0, _ROOT)

import standards_index as _si  # noqa: E402
import tutorial_kb as _tk  # noqa: E402


def search_standards(query: str, agency: str = "", collection: str = "", doc: str = "", k: int = 6) -> list[dict]:
    """Search the official standards library (NYSDOT HDM all chapters, Standard Specifications,
    Standard Sheet books, PDM/CPDM/survey/ROW manuals, NYC DDC General Requirements + spec
    bulletins, NYC DOT Standard Highway Specs + Standard Details, NYC DEP Sewer Design
    Standards / Stormwater Manual, MUTCD Part 6 + NYS Supplement) for a requirement.
    Hybrid retrieval: exact section refs (e.g. '8.7.5.1', '§20.5.4'), keyword BM25 with
    engineering synonyms, semantic vectors, cross-encoder rerank. Each hit has a citation
    (doc, section breadcrumb, page), the passage text, and in-text cross references (`refs`)
    you can resolve with follow_standard_ref. Filter with agency ('NYSDOT', 'NYC-DDC',
    'NYC-DEP', 'FHWA') or collection ('hdm', 'specs-and-std-sheets', 'spec-bulletins', ...).
    Use this BEFORE designing or drafting anything governed by a standard, and cite what you use."""
    return _si.search(query, agency or None, collection or None, doc or None, k=int(k or 6))


def get_standard_section(doc: str, section: str, max_chars: int = 9000) -> str:
    """Full text of one numbered section (and its subsections) of a standard, e.g.
    doc='hdm/chapt_08.pdf', section='8.7.5.1'. Use after search_standards to read the
    whole requirement instead of a fragment."""
    return _si.section_text(doc, section, int(max_chars))


def read_standard_pages(doc: str, first_page: int, last_page: int = 0) -> str:
    """Raw text of specific pages of a standard (1-based), for tables/exhibits or to verify a quote."""
    return _si.read_pages(doc, int(first_page), int(last_page) or None)


def standards_toc(doc: str, max_level: int = 2) -> list[dict]:
    """Table of contents (section ids, titles, page ranges) of one standard - navigate a
    large manual before reading it."""
    return _si.toc(doc, int(max_level))


def follow_standard_ref(doc: str, ref: str) -> str:
    """Resolve a cross reference returned in a search hit's `refs`, e.g. 'section:8.7.4',
    'exhibit:Exhibit-8-12', 'chapter:5', 'spec:603'."""
    return _si.follow(doc, ref)


def list_standards(agency: str = "", collection: str = "") -> list[dict]:
    """List indexed standards documents (doc id, title, agency, collection, pages)."""
    return _si.list_docs(agency or None, collection or None)


def take_snapshot():
    """Pre-edit snapshot of the active model for the QA change set (None if unavailable/truncated)."""
    import qa_gate
    import wztc_ops
    return qa_gate.take_snapshot(wztc_ops)


def oplog_offset() -> int:
    """Byte position of the edit-operation log (op_recorder) - taken before a turn's edits."""
    import op_recorder
    return op_recorder.log_size()


def qa_gate_run(task_type: str, description: str, element_ids: str = "", expectations_json: str = "[]",
                extent: str = "", citations_json: str = "[]", task_id: str = "", snapshot: dict | None = None,
                oplog_offset: int | None = None) -> dict:
    """Run the QA gate (scripts/qa_gate.py) and return PASS/FAIL/UNCERTAIN per requirement.
    task_type: cad.drafting | engineering.calc | generic. element_ids: comma list of the
    elements to check (cad.drafting). expectations_json: list of {id, desc, where:{level,type,
    cellName}, count:">=N"} measurable expectations from the sheet/standard/request.
    citations_json: list of {doc, page, quote}. Re-run after every fix until PASS."""
    import qa_gate
    ext = [float(x) for x in extent.split(",")] if extent else None
    rec = qa_gate.run_inline(task_type, description, ids=element_ids, expect=json.loads(expectations_json or "[]"),
                             extent=ext, citations=json.loads(citations_json or "[]"), task_id=task_id or None,
                             snapshot=snapshot, oplog_offset=oplog_offset)
    return {"overall": rec["overall"], "task": rec["task"], "run": rec["run"], "images": rec.get("images"),
            "results": [{k: r[k] for k in ("id", "status", "observed", "fix") if k in r} | {"evidence": r.get("evidence", [])[:6]}
                        for r in rec["results"]]}


# ---------------------------------------------------------------- engineers' tutorials (Standards/Tutorials)
def list_tutorial_playbooks() -> list[dict]:
    """Distilled how-to playbooks from the engineers' tutorials (topic, title, when to use). Check this BEFORE
    drafting or editing: a playbook gives the working MicroStation/OpenRoads procedure, settings and pitfalls
    for that kind of task, each step cited to a tutorial file and timestamp."""
    return _tk.list_playbooks()


def get_tutorial_playbook(topic: str) -> str:
    """Full text of one tutorial playbook (steps, exact tools/key-ins, settings, pitfalls, QA checks, citations)."""
    return _tk.get_playbook(topic)


def search_tutorials(query: str, k: int = 6, source: str = "") -> list[dict]:
    """Search the engineers' tutorial recordings, decks and the ORD training manual (speech transcripts + on-screen
    text). Each hit has cite (file + slide/page/mm:ss), text and a keyframe image path. Use when no playbook
    covers the task or to check a detail; cite what you use."""
    return _tk.search(query, int(k or 6), source)


def read_tutorial(source: str, start: str = "", end: str = "", max_chars: int = 9000) -> str:
    """Raw text of one tutorial source. Videos: start/end as mm:ss. Decks/PDF: slide/page numbers."""
    return _tk.read_source(source, start, end, int(max_chars or 9000))


def list_tutorials() -> list[dict]:
    """All ingested tutorial sources (videos, decks, PDF, spreadsheet)."""
    return _tk.list_tutorials()


def check_playbook(topic: str) -> dict:
    """Read-only preflight of a tutorial playbook against the OPEN MicroStation file: does this workspace have the
    levels, cells, text styles, line styles and models the playbook uses? Returns per-check PASS/FAIL and the list of
    agent vs engineer steps. Never edits anything."""
    import playbook_runner as pr
    pb = pr.load(topic)
    S = pr.Session()
    rep = pr.run(pb, False, S)
    rep["agent_steps"] = [s["id"] + " " + s["title"] for s in rep["steps"] if s["who"] == "agent"]
    rep["engineer_steps"] = [s["id"] + " " + s["title"] + " - " + s["detail"] for s in rep["steps"] if s["who"] == "engineer"]
    rep.pop("steps")
    return rep
