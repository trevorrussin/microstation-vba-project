"""Searchable knowledge base over the engineers' tutorials (Standards/Tutorials -> Standards/derived/tutorials).

Sources: meeting-recording transcripts (speech + on-screen text, timestamped), PowerPoint decks, the ORD
tutorial PDF, the training spreadsheet. Built by scripts/tutorials_ingest.py; indexed here into SQLite FTS5
(Standards/index/tutorials.sqlite). Every hit carries its source file and where in it (slide / page / mm:ss) so
anything the agent does or the QA gate enforces from a tutorial can cite it, and a keyframe path when one exists.

  python mcp-server/tutorial_kb.py build
  python mcp-server/tutorial_kb.py search "cut a profile named boundary" [--k 6]

Also holds the distilled playbooks (Data/tutorial-playbooks/*.md): step-by-step procedures the engineers
demonstrated, with the exact tool/key-in, settings and pitfalls, each step cited to file@timestamp.
"""
from __future__ import annotations

import glob
import os
import re
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DERIVED = os.path.join(ROOT, "Standards", "derived", "tutorials")
DB = os.path.join(ROOT, "Standards", "index", "tutorials.sqlite")
PLAYBOOKS = os.path.join(ROOT, "Data", "tutorial-playbooks")
WINDOW_S = 75      # transcript chunk length
OVERLAP_S = 15

_STOP = set("the a an and or of to in on for with is are was were be this that it as at by from we you i so then now just can will".split())


def _conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB), exist_ok=True)
    c = sqlite3.connect(DB)
    c.execute("CREATE TABLE IF NOT EXISTS chunks (id INTEGER PRIMARY KEY, source TEXT, kind TEXT, loc TEXT, t0 REAL, "
              "text TEXT, frame TEXT)")
    c.execute("CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(text, source, tokenize='porter unicode61')")
    return c


def _ts(s: str) -> float:
    p = [int(x) for x in s.split(":")]
    return p[0] * 3600 + p[1] * 60 + p[2] if len(p) == 3 else p[0] * 60 + p[1]


def _stamp(sec: float) -> str:
    sec = int(sec)
    return f"{sec // 3600}:{sec % 3600 // 60:02d}:{sec % 60:02d}" if sec >= 3600 else f"{sec // 60:02d}:{sec % 60:02d}"


def _nearest_frame(folder: str, t: float) -> str:
    fr = glob.glob(os.path.join(folder, "frames", "*.jpg"))
    if not fr:
        return ""
    best = min(fr, key=lambda p: abs(int(os.path.basename(p)[:-4]) - t))
    return os.path.relpath(best, ROOT).replace("\\", "/")


def _source_of(path: str) -> str:
    first = open(path, encoding="utf-8").readline()
    m = re.match(r"# SOURCE: (.+?)(\s{2}\(|$)", first.strip())
    return (m.group(1) if m else os.path.dirname(path)).strip()


def _chunk_transcript(path: str):
    folder = os.path.dirname(path)
    src = _source_of(path)
    lines = []
    for ln in open(path, encoding="utf-8").read().splitlines()[1:]:
        m = re.match(r"\[([\d:]+)\] (SAY|SCREEN): (.*)", ln)
        if m:
            lines.append((_ts(m.group(1)), m.group(2), m.group(3)))
    if not lines:
        return
    end = lines[-1][0]
    t = 0.0
    while t <= end:
        win = [x for x in lines if t <= x[0] < t + WINDOW_S]
        if win:
            text = "\n".join(f"[{_stamp(x[0])}] {x[1]}: {x[2]}" for x in win)
            yield src, "video", f"@{_stamp(win[0][0])}", win[0][0], text, _nearest_frame(folder, win[0][0])
        t += WINDOW_S - OVERLAP_S


def _chunk_doc(path: str):
    src = _source_of(path)
    body = open(path, encoding="utf-8").read().split("\n", 1)[1]
    for part in re.split(r"\n(?=## (?:page|slide|sheet) )", body):
        m = re.match(r"## (page|slide|sheet) (.+)", part.strip())
        if not m:
            continue
        loc, text = f"{m.group(1)} {m.group(2).splitlines()[0].strip()}", part.strip()
        for i in range(0, len(text), 2400):  # long pages: several chunks
            yield src, "doc", loc, 0.0, text[i:i + 2800], ""


def build(log=print) -> dict:
    c = _conn()
    c.execute("DELETE FROM chunks"); c.execute("DELETE FROM fts")
    n = 0
    for tr in sorted(glob.glob(os.path.join(DERIVED, "*", "transcript.md"))):
        for src, kind, loc, t0, text, frame in _chunk_transcript(tr):
            cur = c.execute("INSERT INTO chunks(source,kind,loc,t0,text,frame) VALUES (?,?,?,?,?,?)", (src, kind, loc, t0, text, frame))
            c.execute("INSERT INTO fts(rowid,text,source) VALUES (?,?,?)", (cur.lastrowid, text, src)); n += 1
    for tx in sorted(glob.glob(os.path.join(DERIVED, "*", "text.md"))):
        for src, kind, loc, t0, text, frame in _chunk_doc(tx):
            cur = c.execute("INSERT INTO chunks(source,kind,loc,t0,text,frame) VALUES (?,?,?,?,?,?)", (src, kind, loc, t0, text, frame))
            c.execute("INSERT INTO fts(rowid,text,source) VALUES (?,?,?)", (cur.lastrowid, text, src)); n += 1
    c.commit()
    log(f"indexed {n} chunks")
    return {"chunks": n}


def _fts_query(q: str) -> str:
    toks = [t for t in re.findall(r"[A-Za-z0-9_]+", q.lower()) if t not in _STOP and len(t) > 1]
    return " OR ".join(f'"{t}"' for t in toks)


def search(query: str, k: int = 6, source: str = "") -> list[dict]:
    if not os.path.exists(DB):
        return [{"error": "tutorial index not built: python mcp-server/tutorial_kb.py build"}]
    c = _conn()
    q = _fts_query(query)
    if not q:
        return []
    sql = ("SELECT c.id, c.source, c.kind, c.loc, c.text, c.frame, bm25(fts) FROM fts JOIN chunks c ON c.id = fts.rowid "
           "WHERE fts MATCH ?" + (" AND c.source LIKE ?" if source else "") + " ORDER BY bm25(fts) LIMIT ?")
    args = [q] + ([f"%{source}%"] if source else []) + [int(k)]
    out = []
    for cid, src, kind, loc, text, frame, score in c.execute(sql, args):
        out.append({"cite": f"{os.path.basename(src)} {loc}", "source": src, "where": loc, "kind": kind,
                    "text": text[:1800], "keyframe": frame, "id": cid})
    return out


def read_source(source: str, start: str = "", end: str = "", max_chars: int = 9000) -> str:
    """Raw text of a tutorial source. For videos start/end are mm:ss (e.g. '03:00'-'06:30'); for decks/PDF they
    are slide/page numbers."""
    hits = glob.glob(os.path.join(DERIVED, "*"))
    key = re.sub(r"[^a-z0-9]+", "", source.lower())
    match = [h for h in hits if key in re.sub(r"[^a-z0-9]+", "", os.path.basename(h).lower())] or \
            [h for h in hits if key in re.sub(r"[^a-z0-9]+", "", _source_of(next(iter(glob.glob(os.path.join(h, "*.md"))), h)).lower())]
    if not match:
        return f"no tutorial source matches {source!r}; see list_tutorials()"
    folder = match[0]
    tr = os.path.join(folder, "transcript.md")
    if os.path.exists(tr):
        lines = open(tr, encoding="utf-8").read().splitlines()
        a, b = (_ts(start) if start else 0), (_ts(end) if end else 10 ** 9)
        sel = [ln for ln in lines[1:] if (m := re.match(r"\[([\d:]+)\]", ln)) and a <= _ts(m.group(1)) <= b]
        return (lines[0] + "\n" + "\n".join(sel))[:max_chars]
    text = open(os.path.join(folder, "text.md"), encoding="utf-8").read()
    if start:
        parts = re.split(r"\n(?=## (?:page|slide|sheet) )", text)
        want = {str(i) for i in range(int(start), int(end or start) + 1)}
        text = "\n".join(p for p in parts if (m := re.match(r"## \w+ (\d+)", p.strip())) and m.group(1) in want)
    return text[:max_chars]


def list_tutorials() -> list[dict]:
    out = []
    for d in sorted(glob.glob(os.path.join(DERIVED, "*"))):
        f = next((p for p in (os.path.join(d, "transcript.md"), os.path.join(d, "text.md")) if os.path.exists(p)), None)
        if f:
            out.append({"id": os.path.basename(d), "source": _source_of(f), "kind": "video" if f.endswith("transcript.md") else "doc"})
    return out


# ------------------------------------------------------------------ playbooks (distilled procedures)
def list_playbooks() -> list[dict]:
    out = []
    for p in sorted(glob.glob(os.path.join(PLAYBOOKS, "*.md"))):
        txt = open(p, encoding="utf-8").read()
        m = re.search(r"^title:\s*(.+)$", txt, re.M)
        w = re.search(r"^when:\s*(.+)$", txt, re.M)
        out.append({"topic": os.path.basename(p)[:-3], "title": m.group(1).strip() if m else os.path.basename(p),
                    "when": w.group(1).strip() if w else ""})
    return out


def match_playbooks(text: str, k: int = 4) -> list[dict]:
    """Playbooks whose `keywords:` front-matter terms appear in the task text (most matches first)."""
    t = (text or "").lower()
    out = []
    for p in sorted(glob.glob(os.path.join(PLAYBOOKS, "*.md"))):
        src = open(p, encoding="utf-8").read()
        m = re.search(r"^keywords:\s*(.+)$", src, re.M)
        kws = [x.strip().lower() for x in m.group(1).split(",")] if m else []
        hits = sum(1 for kw in kws if kw and kw in t)
        if hits:
            tt = re.search(r"^title:\s*(.+)$", src, re.M)
            out.append((hits, {"topic": os.path.basename(p)[:-3], "title": tt.group(1).strip() if tt else ""}))
    return [x[1] for x in sorted(out, key=lambda r: -r[0])[:k]]


def get_playbook(topic: str) -> str:
    p = os.path.join(PLAYBOOKS, re.sub(r"[^A-Za-z0-9_-]", "", topic) + ".md")
    if not os.path.exists(p):
        known = ", ".join(x["topic"] for x in list_playbooks())
        return f"no playbook {topic!r}. Available: {known}"
    return open(p, encoding="utf-8").read()


def main() -> int:
    import argparse
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("build")
    s = sub.add_parser("search"); s.add_argument("query"); s.add_argument("--k", type=int, default=6); s.add_argument("--source", default="")
    sub.add_parser("list")
    a = ap.parse_args()
    if a.cmd == "build":
        build()
    elif a.cmd == "list":
        for t in list_tutorials():
            print(t["kind"], t["source"])
    else:
        for h in search(a.query, a.k, a.source):
            print(f"{h.get('cite')}  [{h.get('keyframe', '')}]\n   {h.get('text', '')[:300].replace(chr(10), ' ')}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
