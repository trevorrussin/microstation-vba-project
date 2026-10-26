"""Retrieval receipts: proof that the agent read a playbook before it planned or edited from it.

Knowing a playbook's name is not reading it. `tutorial_kb.get_playbook` writes a receipt (topic, sha256 of the file,
time) to outputs/qa/receipts.jsonl each time it returns a real playbook; a failed lookup writes nothing. Consumers:

  - mcp-server/task_guard.py   refuses drawing edits while an open task has a required playbook with no fresh receipt
  - scripts/qa_gate.py start   prints which playbooks must be read
  - scripts/retrieval_prompt_hook.py (UserPromptSubmit) + scripts/qa_stop_hook.py (Stop)
                               make a chat answer about a matched workflow wait for the read, too
  - check_plan_against_playbook  diffs a proposed plan against the playbook's settings (lexical, not semantic)

A receipt is fresh when its hash equals the playbook file's hash now (an edited playbook must be re-read) and it is
younger than WINDOW_S. A task candidate dispositioned `applies=no` (with a reason) is waived.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAYBOOKS = os.path.join(ROOT, "Data", "tutorial-playbooks")
RECEIPTS = os.path.join(ROOT, "outputs", "qa", "receipts.jsonl")
PENDING = os.path.join(ROOT, "outputs", "qa", "pending_retrieval.json")
WINDOW_S = 12 * 3600
MAX_PENDING_BLOCKS = 2
HOOK_MIN_HITS = 2   # distinct playbook keywords a prompt must contain before the hook demands a read
HOOK_MAX_TOPICS = 2  # at most this many playbooks are demanded per prompt (best matches first)


def _path(topic: str) -> str:
    return os.path.join(PLAYBOOKS, re.sub(r"[^A-Za-z0-9_-]", "", topic) + ".md")


def playbook_hash(topic: str) -> str:
    p = _path(topic)
    return hashlib.sha256(open(p, "rb").read()).hexdigest() if os.path.exists(p) else ""


def _key(kind: str, key: str) -> str:
    """Normalised receipt key. Standards and tutorial sources compare by lower-case file name."""
    k = key.replace("\\", "/").strip()
    return k.lower() if kind == "playbook" else os.path.basename(k).lower()


def _append(row: dict) -> None:
    os.makedirs(os.path.dirname(RECEIPTS), exist_ok=True)
    with open(RECEIPTS, "a", encoding="utf-8") as f:
        f.write(json.dumps(row) + "\n")


def _rows():
    if not os.path.exists(RECEIPTS):
        return
    for ln in open(RECEIPTS, encoding="utf-8"):
        try:
            yield json.loads(ln)
        except Exception:
            continue


def record(topic: str) -> None:
    """Receipt for an existing playbook. Called by get_playbook after it read the file."""
    h = playbook_hash(topic)
    if h:
        _append({"kind": "playbook", "topic": topic, "key": _key("playbook", topic), "sha256": h, "at": time.time()})


def record_read(kind: str, key: str) -> None:
    """Receipt that a standard (kind="standard", key=doc id) or tutorial source (kind="tutorial", key=file name) was
    opened and returned text. A tool that returns an error or nothing must not call this."""
    if key:
        _append({"kind": kind, "key": _key(kind, key), "at": time.time()})


def record_search(what: str, query: str) -> None:
    _append({"kind": "search", "what": what, "key": query[:200], "at": time.time()})


def searched_since(t: float) -> bool:
    return any(r.get("kind") == "search" and r.get("at", 0) >= t for r in _rows())


def has_fresh_receipt(topic: str, now: float | None = None) -> bool:
    h = playbook_hash(topic)
    if not h:
        return False
    now = now or time.time()
    return any(r.get("topic") == topic and r.get("sha256") == h and now - r.get("at", 0) <= WINDOW_S for r in _rows())


def has_read(kind: str, key: str, now: float | None = None) -> bool:
    if kind == "playbook":
        return has_fresh_receipt(key, now)
    k, now = _key(kind, key), (now or time.time())
    return any(r.get("kind") == kind and r.get("key") == k and now - r.get("at", 0) <= WINDOW_S for r in _rows())


def required_topics(task: dict) -> list[str]:
    """Playbooks a task must have read: its playbook candidates not dispositioned `no`."""
    return [c["doc"].split(":", 1)[1] for c in task.get("candidates", [])
            if str(c.get("doc", "")).startswith("playbook:") and c.get("applies") != "no"]


def missing_for_task(task: dict) -> list[str]:
    return [t for t in required_topics(task) if not has_fresh_receipt(t)]


def doc_sources_for(topics: list[str]) -> list[dict]:
    """Written sources (PDF, deck, spreadsheet) cited by the playbooks' `sources:` lists. Recordings are not
    required reading: they are searched on demand. Returns [{key, cite}] keyed by file name."""
    import sys
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import tutorial_kb
    docs = [os.path.basename(t["source"].replace("\\", "/")) for t in tutorial_kb.list_tutorials() if t["kind"] == "doc"]
    out, seen = [], set()
    for topic in topics:
        p = _path(topic)
        if not os.path.exists(p):
            continue
        for ln in open(p, encoding="utf-8"):
            m = re.match(r'\s+- "(.+)"\s*$', ln)
            if not m:
                continue
            for d in docs:
                if d.lower() in m.group(1).lower() and d.lower() not in seen:
                    seen.add(d.lower())
                    out.append({"key": d, "cite": f"SOURCE {d} (cited by playbook {topic})"})
    return out


def kind_key(c: dict) -> tuple[str, str]:
    doc = str(c.get("doc", ""))
    if doc.startswith("playbook:"):
        return "playbook", doc.split(":", 1)[1]
    if doc.startswith("tutorial:"):
        return "tutorial", doc.split(":", 1)[1]
    return "standard", doc


def gaps(task: dict) -> list[str]:
    """Everything that must happen before the task may edit: each candidate (playbook, standard, written tutorial
    source) is dispositioned, and each one that applies has been read. `applies=no` needs only the reason."""
    out = []
    for c in list(task.get("candidates", [])) + list(task.get("sources", [])):
        kind, key = kind_key(c)
        ap = c.get("applies")
        if ap == "no":
            continue
        if kind == "playbook" and ap is None:
            if not has_read(kind, key):
                out.append(f"read playbook {key}")
            continue  # a playbook that was read needs no separate disposition
        if ap is None:
            out.append(f"disposition {kind} {key}")
        elif not has_read(kind, key):
            out.append(f"read {kind} {key}")
    return out


# ------------------------------------------------------------------ chat answers (UserPromptSubmit + Stop hooks)
def matched_for_text(text: str) -> list[dict]:
    import sys
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import tutorial_kb
    return tutorial_kb.match_playbooks(text)


def unread_for_text(text: str) -> list[dict]:
    import sys
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import tutorial_kb
    return [p for p in tutorial_kb.match_playbooks(text) if not has_fresh_receipt(p["topic"])]


_PASTED = re.compile(r"<pasted_content\b[^>]*>.*?</pasted_content[^>]*>", re.S | re.I)
_DISCUSSION = re.compile(r"\b(what do you think|your (?:opinion|thoughts|take)|thoughts on|do you agree|critique|assess(?:ment)? of|"
                         r"is this (?:a good|right)|review (?:this|the) (?:approach|architecture|design|process))\b", re.I)


def prompt_text_for_matching(prompt: str) -> str:
    """The engineer's own words. Pasted transcripts, logs and other agents' reports are material to discuss, not a task:
    their keywords must not arm the gate (found 2026-10-24: a pasted review mentioning 'levels' and 'alignments' demanded
    four unrelated playbooks)."""
    return _PASTED.sub(" ", prompt or "")


def hook_matches(prompt: str) -> list[dict]:
    """Playbooks the UserPromptSubmit hook may demand. Empty for discussion questions, and for prompts whose own words
    (pasted material removed) match no playbook on at least HOOK_MIN_HITS keywords."""
    own = prompt_text_for_matching(prompt)
    if _DISCUSSION.search(own) and len(own.split()) < 60:
        return []
    import sys
    sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
    import tutorial_kb
    return tutorial_kb.match_playbooks(own, k=HOOK_MAX_TOPICS, min_hits=HOOK_MIN_HITS)


def load_pending() -> dict:
    try:
        return json.load(open(PENDING, encoding="utf-8"))
    except Exception:
        return {}


def already_told(topics: list[str]) -> bool:
    """True when this exact demand was already issued inside the window and the stop hook has since given up on it, so
    the same text is not injected again (the loop seen 2026-10-24: the same RETRIEVE text on every send)."""
    p = load_pending()
    return bool(topics) and sorted(p.get("told", [])) == sorted(topics) and time.time() - p.get("told_at", 0) <= WINDOW_S         and p.get("blocks", 0) >= MAX_PENDING_BLOCKS


def set_pending(topics: list[str], need_search: bool = False) -> None:
    os.makedirs(os.path.dirname(PENDING), exist_ok=True)
    if topics or need_search:
        json.dump({"at": time.time(), "topics": topics, "need_search": need_search, "blocks": 0,
                   "told": topics, "told_at": time.time()}, open(PENDING, "w", encoding="utf-8"))
    elif os.path.exists(PENDING):
        json.dump({"at": time.time(), "topics": [], "blocks": 0}, open(PENDING, "w", encoding="utf-8"))


def stop_message() -> str:
    """Non-empty while the last prompt matched playbooks that still have no receipt. Gives up after MAX_PENDING_BLOCKS
    so a broken lookup cannot trap the session."""
    if not os.path.exists(PENDING):
        return ""
    try:
        p = json.load(open(PENDING, encoding="utf-8"))
    except Exception:
        return ""
    left = [t for t in p.get("topics", []) if not has_fresh_receipt(t)]
    no_search = bool(p.get("need_search")) and not searched_since(p.get("at", 0))
    if not (left or no_search) or time.time() - p.get("at", 0) > WINDOW_S:
        return ""
    todo = []
    if left:
        todo.append("get_tutorial_playbook for: " + ", ".join(left) + " (and any playbook they point to)")
    if no_search:
        todo.append("search_standards for the governing requirement and search_tutorials / read_tutorial for the "
                    "written procedure the playbooks cite")
    if p.get("blocks", 0) >= MAX_PENDING_BLOCKS:
        p["topics"], p["need_search"] = [], False
        json.dump(p, open(PENDING, "w", encoding="utf-8"))
        return "Retrieval was skipped (" + "; ".join(todo) + "). Tell the engineer the answer is NOT backed by those sources."
    p["blocks"] = p.get("blocks", 0) + 1
    json.dump(p, open(PENDING, "w", encoding="utf-8"))
    return ("This request is governed by sources you have not consulted. Call " + "; then ".join(todo) +
            ". Answer from the text and cite it. If retrieval fails, say so; do not answer from memory.")


# ------------------------------------------------------------------ plan-to-playbook diff
_NUM = re.compile(r"(?<![\w.])\d+(?:\.\d+)?(?:\+\d+(?:\.\d+)?)?(?:\s?(?:ft|%))?(?![\w])")
_QUOTED = re.compile(r"`([^`]+)`|\"([^\"]{2,60})\"")
_TITLE = re.compile(r"\b(?:[A-Z][a-z]+)(?:\s+[A-Z][a-z]+){1,4}\b")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.lower().replace("“", '"').replace("”", '"')).strip()


def _steps(src: str) -> list[dict]:
    out = []
    for m in re.finditer(r"^    - id: (\w+)\n(.*?)(?=^    - id: |\Z)", src, re.M | re.S):
        body = m.group(2)
        who = (re.search(r"who:\s*(\w+)", body) or [None, ""])[1]
        title = (re.search(r"title:\s*(.+)", body) or [None, ""])[1].strip()
        how = re.search(r"how:\s*\|\n(.*?)(?=^      (?:cite|pitfall|test|manual_reason):|\Z)", body, re.M | re.S)
        pit = re.search(r"pitfall:\s*(?:\|\n)?(.*?)(?=^      (?:cite|test):|\Z)", body, re.M | re.S)
        out.append({"id": m.group(1), "who": who, "title": title,
                    "text": (how.group(1) if how else "") + "\n" + (pit.group(1) if pit else "")})
    return out


def _tokens(text: str) -> list[str]:
    """Settings worth carrying into a plan. Sentences marked as examples (`example`, `e.g.`) hold the recording's
    project values, not settings, so they are skipped."""
    text = text.replace("e.g.", "eg")
    parts = re.split(r"(?<=[.!?])\s+|\n", text)
    text = " ".join(x for x in parts if "example" not in x.lower() and not re.search(r"\beg\b", x))
    toks = [n.group(0).strip() for n in _NUM.finditer(text)]
    toks += [a or b for a, b in _QUOTED.findall(text)]
    toks += _TITLE.findall(text)
    seen, out = set(), []
    for t in toks:
        k = _norm(t)
        if k and k not in seen:
            seen.add(k)
            out.append(t)
    return out


def check_plan_against_playbook(topic: str, plan: str) -> dict:
    """Lexical diff of a plan against one playbook. For each step the plan references (by id such as `N5`, or by a
    setting from that step), list the numbers, quoted values and capitalised setting names from the playbook that
    the plan leaves out. Also lists numbers in the plan that appear nowhere in the playbook. It finds omissions
    and invented values; it does not judge whether wording is right, so read the playbook for anything it flags."""
    p = _path(topic)
    if not os.path.exists(p):
        return {"status": "NO_PLAYBOOK", "topic": topic}
    src = open(p, encoding="utf-8").read()
    pl = _norm(plan)
    steps, report, unmentioned = _steps(src), [], []
    for s in steps:
        toks = _tokens(s["text"])
        present = [t for t in toks if _norm(t) in pl]
        named = re.search(r"(?<![\w])" + re.escape(s["id"].lower()) + r"(?![\w])", pl) is not None
        if not (named or present):
            unmentioned.append(f"{s['id']} {s['title']}")
            continue
        missing = [t for t in toks if _norm(t) not in pl]
        if missing:
            report.append({"step": s["id"], "title": s["title"], "who": s["who"], "missing": missing})
    book = _norm(src)
    invented = [n.group(0).strip() for n in _NUM.finditer(plan) if _norm(n.group(0)) not in book]
    status = "OK" if not report and not invented else "REVIEW"
    return {"status": status, "topic": topic, "sha256": playbook_hash(topic)[:12], "steps_with_omissions": report,
            "steps_not_mentioned": unmentioned, "values_not_in_playbook": sorted(set(invented))}
