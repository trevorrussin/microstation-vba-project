"""Standards library index + hybrid retrieval (structure-aware RAG).

Indexes every PDF in Standards/registry.json into Standards/index/standards.sqlite
(+ dense embeddings in Standards/index/emb.npy). Built so the agent can pull exact
requirements out of very large PDFs (HDM, specs volumes, standard-sheet books,
DEP/DOT standards) without reading them end to end.

Indexing (per PDF, incremental by sha256):
  * Structure first: PDF outline (bookmarks) -> section tree. If the outline is
    missing/sparse, headings are detected from numbered lines (8.7.4.4, 6A.01,
    SECTION 603, 20.5.4) with larger/bold fonts.
  * Repeated page headers/footers removed before chunking.
  * Section-bounded chunks (~1500 chars, 250 overlap). Every chunk carries
    doc / agency / collection / section id / breadcrumb path / page range, and the
    breadcrumb is prepended for embedding ("contextual chunk headers").
  * Tables extracted separately (PyMuPDF find_tables -> markdown) so numeric
    criteria are retrievable as whole tables, not shredded text.
  * Drawing pages (standard sheets, details: short scattered labels) become one
    "sheet" chunk per page keyed by sheet numbers found on the page.

Retrieval (search()):
  1. exact section-reference lookup (query mentions "8.7.4" / "§20.5.4" / "603-01")
  2. BM25 (SQLite FTS5, porter stemming, breadcrumb/title weighted) with an
     engineering synonym expansion (CB -> catch basin, DWS -> detectable warning, ...)
  3. dense vector search (local sentence-transformers model, cosine)
  4. reciprocal-rank fusion + boosts (heading-term match, tables for numeric
     questions, filters by agency/collection/doc)
  5. small-to-big: hits are returned with breadcrumb + page cite and, on
     request, the full parent section text.

CLI:
  python mcp-server/standards_index.py build            # index new/changed PDFs
  python mcp-server/standards_index.py build --no-embed # BM25 only (fast)
  python mcp-server/standards_index.py search "minimum pipe cover storm drain" --agency NYSDOT
  python mcp-server/standards_index.py section hdm/chapt_08.pdf 8.7.4
  python mcp-server/standards_index.py stats
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sqlite3
import statistics
import sys
import time
from collections import Counter, defaultdict
from typing import Any

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STD = os.path.join(ROOT, "Standards")
REGISTRY = os.path.join(STD, "registry.json")
IDX_DIR = os.path.join(STD, "index")
DB_PATH = os.path.join(IDX_DIR, "standards.sqlite")
EMB_PATH = os.path.join(IDX_DIR, "emb.npy")
EMB_IDS_PATH = os.path.join(IDX_DIR, "emb_ids.npy")
SYN_PATH = os.path.join(STD, "synonyms.tsv")
EMB_MODEL = "BAAI/bge-small-en-v1.5"

CHUNK_CHARS = 1500
CHUNK_OVERLAP = 250

# numbered headings seen in NYSDOT / MUTCD / DEP / DDC documents
SEC_RE = re.compile(
    r"^\s*(?:§\s*)?("
    r"\d{1,2}[A-Z]\.\d{2}[A-Za-z]?"            # 6A.01 (MUTCD)
    r"|\d{1,3}(?:\.\d{1,3}){1,5}[A-Z]?"        # 8.7.4.4, 20.5.4, 4.2B
    r"|SECTION\s+\d{3}(?:-\d+)?"               # SECTION 603
    r"|\d{3}-\d{1,2}(?:\.\d{1,2})?"            # 603-1.01 / 619-10
    r"|[A-Z]-\d{1,3}(?:\.\d{1,3})*"            # GR-4 style / Appendix A-1
    r")[\s\.\-:]+(\S.{1,110})$"
)
REF_RE = re.compile(r"(?:§\s*)?\b(\d{1,2}[A-Z]\.\d{2}|\d{1,3}(?:\.\d{1,3}){1,5}[A-Z]?|\d{3}-\d{1,2}(?:\.\d{1,2})?)\b")
SHEET_NUM_RE = re.compile(r"\b\d{3}-\d{2,3}\b")
NUMERIC_INTENT = re.compile(r"\b(min(imum)?|max(imum)?|table|exhibit|spacing|width|slope|cover|diameter|"
                            r"radius|length|depth|height|rate|value|offset|clearance|size|ft|in\.?|%)\b", re.I)

_model = None
_reranker = None
RERANK_MODEL = "BAAI/bge-reranker-base"
# in-text cross references -> link graph (GraphRAG-lite): "Section 8.7.5", "§20.5.4", "Exhibit 8-12", "Chapter 5"
XREF_RES = [
    ("section", re.compile(r"(?:Section|Sections|§)\s*(\d{1,3}(?:\.\d{1,3}){1,5}[A-Z]?|\d[A-Z]\.\d{2})", re.I)),
    ("exhibit", re.compile(r"\b(Exhibit|Table|Figure)\s+(\d{1,3}[A-Z]?-\d{1,3}[a-z]?)", re.I)),
    ("chapter", re.compile(r"\bChapter\s+(\d{1,2})\b", re.I)),
    ("spec", re.compile(r"\b(?:Section|Item)\s+(\d{3})(?:-\d)?\b")),
]


# ---------------------------------------------------------------- storage
def connect() -> sqlite3.Connection:
    os.makedirs(IDX_DIR, exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.executescript(
        """
        CREATE TABLE IF NOT EXISTS docs(doc_id TEXT PRIMARY KEY, path TEXT, title TEXT, agency TEXT,
            collection TEXT, sha256 TEXT, pages INT, n_chunks INT, outline INT, indexed_at TEXT);
        CREATE TABLE IF NOT EXISTS sections(doc_id TEXT, sec_id TEXT, title TEXT, level INT, path TEXT,
            page_start INT, page_end INT);
        CREATE INDEX IF NOT EXISTS sec_doc ON sections(doc_id, sec_id);
        CREATE TABLE IF NOT EXISTS chunks(id INTEGER PRIMARY KEY, doc_id TEXT, kind TEXT, sec_id TEXT,
            sec_path TEXT, page_start INT, page_end INT, text TEXT);
        CREATE INDEX IF NOT EXISTS chunk_doc ON chunks(doc_id);
        CREATE TABLE IF NOT EXISTS refs(chunk_id INT, doc_id TEXT, ref_type TEXT, target TEXT);
        CREATE INDEX IF NOT EXISTS refs_chunk ON refs(chunk_id);
        CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(text, sec_path, title,
            content='', tokenize='porter unicode61');
        """
    )
    return c


def doc_id_for(path: str) -> str:
    rel = path.replace("\\", "/")
    for pre in ("Standards/", "Project Documentation/"):
        if rel.startswith(pre):
            rel = rel[len(pre):]
    return rel


# ---------------------------------------------------------------- PDF parsing
def _page_lines(page) -> list[dict]:
    """Lines with font size / bold, reading order."""
    out = []
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        for ln in b.get("lines", []):
            spans = [s for s in ln.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            text = "".join(s["text"] for s in spans).strip()
            size = max(s["size"] for s in spans)
            bold = any("bold" in s.get("font", "").lower() or (s.get("flags", 0) & 16) for s in spans)
            out.append({"text": text, "size": size, "bold": bold, "y": ln["bbox"][1]})
    return out


def _strip_repeats(pages_lines: list[list[dict]]) -> list[list[dict]]:
    """Remove header/footer lines that repeat on many pages (normalize digits)."""
    norm = lambda t: re.sub(r"\d+", "#", t.strip().lower())
    cnt = Counter()
    for lines in pages_lines:
        edge = lines[:3] + lines[-3:]
        cnt.update({norm(l["text"]) for l in edge})
    n = max(1, len(pages_lines))
    rep = {k for k, v in cnt.items() if v >= max(3, 0.4 * n)}
    return [[l for i, l in enumerate(lines) if not ((i < 3 or i >= len(lines) - 3) and norm(l["text"]) in rep)]
            for lines in pages_lines]


def _is_drawing_page(lines: list[dict]) -> bool:
    if len(lines) < 25:
        return False
    avg = statistics.mean(len(l["text"]) for l in lines)
    return avg < 22


def _sections_from_outline(toc: list, n_pages: int) -> list[dict]:
    secs = []
    for i, (lvl, title, page) in enumerate(toc):
        title = re.sub(r"\s+", " ", str(title)).strip()
        m = SEC_RE.match(title)
        sec_id = m.group(1) if m else title[:40]
        secs.append({"level": lvl, "title": title, "sec_id": sec_id, "page_start": max(1, page)})
    for i, s in enumerate(secs):
        nxt = next((t["page_start"] for t in secs[i + 1:] if t["level"] <= s["level"]), n_pages)
        s["page_end"] = max(s["page_start"], nxt)
    return secs


def _sections_from_headings(pages_lines: list[list[dict]]) -> list[dict]:
    sizes = [l["size"] for lines in pages_lines for l in lines]
    if not sizes:
        return []
    body = statistics.median(sizes)
    secs = []
    for pno, lines in enumerate(pages_lines, start=1):
        for l in lines:
            t = l["text"]
            if len(t) > 130:
                continue
            m = SEC_RE.match(t)
            if not m:
                continue
            if not (l["size"] >= body + 0.4 or l["bold"] or t.isupper()):
                continue
            sid = m.group(1)
            if re.fullmatch(r"\d{1,3}\.\d{1,3}", sid) and l["size"] < body + 0.4 and not l["bold"]:
                continue  # plain "3.5 ft" style lines
            depth = sid.count(".") + 1 if re.match(r"^\d", sid) else 1
            secs.append({"level": depth, "title": t, "sec_id": sid, "page_start": pno})
    n = len(pages_lines)
    for i, s in enumerate(secs):
        nxt = next((t["page_start"] for t in secs[i + 1:] if t["level"] <= s["level"]), n)
        s["page_end"] = max(s["page_start"], nxt)
    return secs


def _breadcrumbs(secs: list[dict]) -> None:
    stack: list[dict] = []
    for s in secs:
        while stack and stack[-1]["level"] >= s["level"]:
            stack.pop()
        stack.append(s)
        s["path"] = " > ".join(x["title"][:70] for x in stack)


def _split(text: str) -> list[str]:
    text = re.sub(r"[ \t]+", " ", text)
    if len(text) <= CHUNK_CHARS:
        return [text] if text.strip() else []
    out, i = [], 0
    while i < len(text):
        j = min(len(text), i + CHUNK_CHARS)
        if j < len(text):
            k = text.rfind("\n", i + CHUNK_CHARS // 2, j)
            if k > 0:
                j = k
        out.append(text[i:j].strip())
        if j >= len(text):
            break
        i = max(j - CHUNK_OVERLAP, i + 1)
    return [o for o in out if o]


def _tables(page) -> list[str]:
    try:
        tabs = page.find_tables()
    except Exception:
        return []
    md = []
    for t in getattr(tabs, "tables", [])[:6]:
        try:
            rows = t.extract()
        except Exception:
            continue
        rows = [[(c or "").replace("\n", " ").strip() for c in r] for r in rows if r]
        if len(rows) < 2 or max(len(r) for r in rows) < 2:
            continue
        lines = ["| " + " | ".join(r) + " |" for r in rows[:60]]
        lines.insert(1, "|" + "---|" * len(rows[0]))
        md.append("\n".join(lines))
    return md


_ocr_engine = None


def _ocr_page(page, sha: str, pno: int) -> list[dict]:
    """OCR an image-only (scanned) page; cached by (sha256, page) in Standards/index/ocr_cache.sqlite."""
    global _ocr_engine
    os.makedirs(IDX_DIR, exist_ok=True)
    cc = sqlite3.connect(os.path.join(IDX_DIR, "ocr_cache.sqlite"))
    cc.execute("CREATE TABLE IF NOT EXISTS ocr(sha TEXT, page INT, lines TEXT, PRIMARY KEY(sha, page))")
    row = cc.execute("SELECT lines FROM ocr WHERE sha=? AND page=?", (sha, pno)).fetchone()
    if row:
        return json.loads(row[0])
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    import numpy as np
    pix = page.get_pixmap(dpi=170)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = img[:, :, :3]
    res, _ = _ocr_engine(img)
    lines = []
    for box, text, score in sorted(res or [], key=lambda r: (round(r[0][0][1] / 12), r[0][0][0])):
        if float(score) < 0.5 or not text.strip():
            continue
        h = abs(box[3][1] - box[0][1]) * 72 / 170
        lines.append({"text": text.strip(), "size": round(h, 1), "bold": False, "y": box[0][1]})
    cc.execute("INSERT OR REPLACE INTO ocr VALUES(?,?,?)", (sha, pno, json.dumps(lines)))
    cc.commit()
    return lines


def parse_pdf(path: str, title: str, sha: str = "") -> dict:
    import fitz

    doc = fitz.open(path)
    n = doc.page_count
    pages_lines = []
    n_ocr = 0
    for pno, p in enumerate(doc, start=1):
        try:
            lines = _page_lines(p)
            if not lines and p.get_images() and sha:
                lines = _ocr_page(p, sha, pno)
                n_ocr += 1
            pages_lines.append(lines)
        except Exception:
            pages_lines.append([])
    pages_lines = _strip_repeats(pages_lines)
    toc = doc.get_toc(simple=True)
    use_outline = len(toc) >= max(3, n // 40)
    secs = _sections_from_outline(toc, n) if use_outline else _sections_from_headings(pages_lines)
    if not secs:
        secs = [{"level": 1, "title": title, "sec_id": "", "page_start": 1, "page_end": n}]
    _breadcrumbs(secs)

    # assign each page to the deepest section starting at/before it
    page_sec: dict[int, dict] = {}
    for s in sorted(secs, key=lambda s: (s["page_start"], s["level"])):
        for pg in range(s["page_start"], s["page_end"] + 1):
            cur = page_sec.get(pg)
            if cur is None or s["level"] >= cur["level"] or s["page_start"] >= cur["page_start"]:
                page_sec[pg] = s
    chunks = []
    # Walk lines in reading order; switch section exactly at its heading line
    # (falls back to the page start when the heading text is not found on the page).
    norm = lambda t: re.sub(r"[^a-z0-9]", "", t.lower())[:28]
    starts: dict[int, list[dict]] = defaultdict(list)
    for s in secs:
        starts[s["page_start"]].append(s)
    cur = secs[0]
    groups: list[tuple[dict, list[tuple[str, int]]]] = []
    for pno in range(1, n + 1):
        lines = pages_lines[pno - 1]
        pending = list(starts.get(pno, []))
        if _is_drawing_page(lines):
            if pending:
                cur = pending[-1]
            txt = " ".join(l["text"] for l in lines)
            sheets = sorted(set(SHEET_NUM_RE.findall(txt)))[:12]
            chunks.append({"kind": "sheet", "sec_id": ",".join(sheets) or cur["sec_id"],
                           "sec_path": (cur["path"] + (" > sheet " + ", ".join(sheets) if sheets else "")),
                           "page_start": pno, "page_end": pno, "text": txt[:6000]})
            continue
        keys = {norm(s["title"]): s for s in pending}
        for li, l in enumerate(lines):
            nl = norm(l["text"])
            # heading number on its own line, title on the next ("20.6" / "REJECTION OF ...")
            if re.fullmatch(r"\s*(?:§\s*)?\d{1,3}(?:\.\d{1,3})*[A-Z]?\.?\s*", l["text"]) and li + 1 < len(lines):
                nl = norm(l["text"] + " " + lines[li + 1]["text"])
            hit = keys.get(nl) if nl else None
            if hit is None and pending and nl:
                hit = next((s for k, s in keys.items() if k and (nl.startswith(k[:20]) or k.startswith(nl[:20]))
                            and len(nl) >= 6), None)
            if hit is not None:
                cur = hit
                keys = {k: s for k, s in keys.items() if s is not hit}
                pending = [s for s in pending if s is not hit]
            if groups and groups[-1][0] is cur:
                groups[-1][1].append((l["text"], pno))
            else:
                groups.append((cur, [(l["text"], pno)]))
        if pending:  # headings not found in text: section starts at next page boundary
            cur = pending[-1]
    for s, lines_pg in groups:
        buf = [t for t, _ in lines_pg]
        page_of = [p for _, p in lines_pg]
        pnos = [page_of[0], page_of[-1]]
        text = "\n".join(buf)
        pos = 0
        for piece in _split(text):
            k = text.find(piece[:60], pos)
            line_idx = text[:max(k, 0)].count("\n")
            p0 = page_of[min(line_idx, len(page_of) - 1)] if page_of else pnos[0]
            p1 = page_of[min(line_idx + piece.count("\n"), len(page_of) - 1)] if page_of else pnos[-1]
            chunks.append({"kind": "text", "sec_id": s["sec_id"], "sec_path": s["path"],
                           "page_start": p0, "page_end": p1, "text": piece})
            pos = max(k, pos)
    # tables on pages that look tabular
    for pno, p in enumerate(doc, start=1):
        lines = pages_lines[pno - 1]
        if not lines or _is_drawing_page(lines):
            continue
        joined = " ".join(l["text"] for l in lines)
        digits = sum(ch.isdigit() for ch in joined) / max(1, len(joined))
        if digits < 0.08 and not re.search(r"\b(Exhibit|Table)\b", joined):
            continue
        s = page_sec.get(pno, secs[0])
        for md in _tables(p):
            chunks.append({"kind": "table", "sec_id": s["sec_id"], "sec_path": s["path"] + " > table",
                           "page_start": pno, "page_end": pno, "text": md[:6000]})
    return {"pages": n, "outline": int(use_outline), "sections": secs, "chunks": chunks, "ocr_pages": n_ocr}


# ---------------------------------------------------------------- build
def _model_get():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        _model = SentenceTransformer(EMB_MODEL, device="cpu")
    return _model


def _reranker_get():
    global _reranker
    if _reranker is None:
        from sentence_transformers import CrossEncoder
        _reranker = CrossEncoder(RERANK_MODEL, device="cpu", max_length=512)
    return _reranker


def follow(doc: str, ref: str, max_chars: int = 6000) -> str:
    """Resolve a cross reference ('section:8.7.5', 'exhibit:Exhibit-8-12', 'chapter:5', 'spec:603')."""
    rtype, _, tgt = ref.partition(":")
    c = connect()
    if rtype == "section":
        return section_text(doc, tgt, max_chars)
    if rtype == "exhibit":
        label = tgt.replace("-", " ", 1)
        rows = c.execute("SELECT c.doc_id, c.page_start, c.text FROM chunks_fts f JOIN chunks c ON c.id=f.rowid "
                         "WHERE chunks_fts MATCH ? AND c.doc_id LIKE ? ORDER BY c.kind='table' DESC LIMIT 3",
                         (f'"{label}"', f"%{doc}%")).fetchall()
        return "\n\n".join(f"[{r[0]} p.{r[1]}]\n{r[2]}" for r in rows)[:max_chars]
    if rtype in ("chapter", "spec"):
        return "\n".join(f"{x['section']} {x['title']} p.{x['pages']}" for x in toc(f"{tgt}", 2))[:max_chars]
    return ""


def build(embed: bool = True, only: str | None = None, log=print, force: bool = False) -> dict:
    with open(REGISTRY, encoding="utf-8") as f:
        reg = json.load(f)
    c = connect()
    have = {r[0]: r[1] for r in c.execute("SELECT doc_id, sha256 FROM docs")}
    done = 0
    for d in reg["documents"]:
        if not d["path"].lower().endswith(".pdf"):
            continue
        p = os.path.join(ROOT, d["path"])
        if not os.path.exists(p):
            continue
        did = doc_id_for(d["path"])
        if only and only not in did:
            continue
        if have.get(did) == d["sha256"] and not force:
            continue
        t0 = time.time()
        try:
            parsed = parse_pdf(p, d.get("title") or did, d["sha256"])
        except Exception as ex:
            log(f"  PARSE FAIL {did}: {ex}")
            continue
        old = [r[0] for r in c.execute("SELECT id FROM chunks WHERE doc_id=?", (did,))]
        for cid, txt, sp in c.execute("SELECT id, text, sec_path FROM chunks WHERE doc_id=?", (did,)).fetchall():
            c.execute("INSERT INTO chunks_fts(chunks_fts, rowid, text, sec_path, title) VALUES('delete',?,?,?,?)",
                      (cid, txt, sp, d.get("title") or did))
        c.execute("DELETE FROM refs WHERE doc_id=?", (did,))
        c.execute("DELETE FROM chunks WHERE doc_id=?", (did,))
        c.execute("DELETE FROM sections WHERE doc_id=?", (did,))
        for s in parsed["sections"]:
            c.execute("INSERT INTO sections VALUES(?,?,?,?,?,?,?)",
                      (did, s["sec_id"], s["title"], s["level"], s["path"], s["page_start"], s["page_end"]))
        for ch in parsed["chunks"]:
            cur = c.execute("INSERT INTO chunks(doc_id,kind,sec_id,sec_path,page_start,page_end,text) VALUES(?,?,?,?,?,?,?)",
                            (did, ch["kind"], ch["sec_id"], ch["sec_path"], ch["page_start"], ch["page_end"], ch["text"]))
            c.execute("INSERT INTO chunks_fts(rowid, text, sec_path, title) VALUES(?,?,?,?)",
                      (cur.lastrowid, ch["text"], ch["sec_path"], d.get("title") or did))
            for rtype, rre in XREF_RES:
                for mm in set(rre.findall(ch["text"])):
                    tgt = "-".join(mm) if isinstance(mm, tuple) else mm
                    c.execute("INSERT INTO refs VALUES(?,?,?,?)", (cur.lastrowid, did, rtype, tgt))
        c.execute("INSERT OR REPLACE INTO docs VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (did, d["path"], d.get("title") or did, d["agency"], d["collection"], d["sha256"],
                   parsed["pages"], len(parsed["chunks"]), parsed["outline"], time.strftime("%Y-%m-%d %H:%M")))
        c.commit()
        done += 1
        log(f"  indexed {did}: {parsed['pages']}p {len(parsed['sections'])} sec {len(parsed['chunks'])} chunks "
            f"outline={parsed['outline']} ocr={parsed.get('ocr_pages', 0)} {time.time()-t0:.0f}s")
    if embed:
        _embed_missing(c, log)
    return stats()


def _embed_missing(c: sqlite3.Connection, log=print) -> None:
    import numpy as np

    ids_all = [r[0] for r in c.execute("SELECT id FROM chunks ORDER BY id")]
    if os.path.exists(EMB_PATH) and os.path.exists(EMB_IDS_PATH):
        emb = np.load(EMB_PATH)
        eids = np.load(EMB_IDS_PATH)
        keep = np.isin(eids, ids_all)
        emb, eids = emb[keep], eids[keep]
    else:
        emb, eids = np.zeros((0, 384), dtype=np.float16), np.zeros((0,), dtype=np.int64)
    have = set(eids.tolist())
    todo = [i for i in ids_all if i not in have]
    if not todo:
        return
    log(f"  embedding {len(todo)} chunks with {EMB_MODEL} ...")
    m = _model_get()
    rows = {}
    for i in range(0, len(todo), 2000):
        part = todo[i:i + 2000]
        q = ",".join("?" * len(part))
        for cid, sp, txt in c.execute(f"SELECT id, sec_path, text FROM chunks WHERE id IN ({q})", part):
            rows[cid] = f"{sp}\n{txt}"[:2000]
    texts = [rows[i] for i in todo]
    vecs = m.encode(texts, batch_size=64, normalize_embeddings=True, show_progress_bar=False).astype(np.float16)
    emb = np.vstack([emb, vecs])
    eids = np.concatenate([eids, np.array(todo, dtype=np.int64)])
    np.save(EMB_PATH, emb)
    np.save(EMB_IDS_PATH, eids)
    log(f"  embeddings: {len(eids)} total")


# ---------------------------------------------------------------- retrieval
def _synonyms() -> dict[str, list[str]]:
    syn: dict[str, list[str]] = {}
    if os.path.exists(SYN_PATH):
        for line in open(SYN_PATH, encoding="utf-8"):
            if line.startswith("#") or "\t" not in line:
                continue
            k, v = line.rstrip("\n").split("\t", 1)
            syn[k.lower()] = [x.strip() for x in v.split("|") if x.strip()]
    return syn


def _fts_query(q: str) -> str:
    syn = _synonyms()
    words = re.findall(r"[A-Za-z0-9][A-Za-z0-9\-\.]*", q)
    parts = []
    # reverse expansion: a spelled-out phrase in the query also searches its abbreviation
    ql = " " + re.sub(r"\s+", " ", q.lower()) + " "
    for k, alts in syn.items():
        if any(" " in a and f" {a.lower()} " in ql for a in alts):
            parts.append(f'"{k}"')
    for w in words:
        wl = w.lower().strip(".")
        if len(wl) < 2 or wl in {"the", "and", "for", "what", "is", "of", "to", "in", "a", "an", "how", "per", "on"}:
            continue
        alts = [wl] + [a.lower() for a in syn.get(wl, [])]
        parts.append("(" + " OR ".join('"' + a.replace('"', "") + '"' for a in alts) + ")")
    return " OR ".join(parts) if parts else '""'


def _filters(agency, collection, doc) -> tuple[str, list]:
    cl, args = [], []
    if agency:
        cl.append("d.agency = ?"); args.append(agency)
    if collection:
        cl.append("d.collection = ?"); args.append(collection)
    if doc:
        cl.append("c.doc_id LIKE ?"); args.append(f"%{doc}%")
    return (" AND " + " AND ".join(cl)) if cl else "", args


def search(query: str, agency: str | None = None, collection: str | None = None, doc: str | None = None,
           k: int = 8, dense: bool = True, expand_section: bool = False, rerank: bool = True,
           follow_refs: bool = True) -> list[dict]:
    c = connect()
    where, args = _filters(agency, collection, doc)
    ranked: dict[int, float] = defaultdict(float)
    RRF = 60

    # 1. exact section refs
    refs = REF_RE.findall(query)
    for r in refs:
        for (cid,) in c.execute(
            f"SELECT c.id FROM chunks c JOIN docs d ON d.doc_id=c.doc_id WHERE (c.sec_id=? OR c.sec_id LIKE ?){where} "
            f"ORDER BY c.page_start LIMIT 5", [r, r + ".%", *args]):
            ranked[cid] += 2.0 / (RRF + 1)

    # 2. BM25
    fq = _fts_query(query)
    try:
        rows = c.execute(
            f"SELECT c.id FROM chunks_fts f JOIN chunks c ON c.id=f.rowid JOIN docs d ON d.doc_id=c.doc_id "
            f"WHERE chunks_fts MATCH ?{where} ORDER BY bm25(chunks_fts, 1.0, 3.0, 2.0) LIMIT 60", [fq, *args]).fetchall()
    except sqlite3.OperationalError:
        rows = []
    for r, (cid,) in enumerate(rows):
        ranked[cid] += 1.0 / (RRF + r + 1)

    # 3. dense
    if dense and os.path.exists(EMB_PATH):
        import numpy as np
        emb = np.load(EMB_PATH)  # not mmap: a Windows mmap would lock emb.npy against rebuilds
        eids = np.load(EMB_IDS_PATH)
        qv = _model_get().encode([query], normalize_embeddings=True)[0].astype(np.float32)
        sims = emb.astype(np.float32) @ qv
        if where:
            allowed = {r[0] for r in c.execute(
                f"SELECT c.id FROM chunks c JOIN docs d ON d.doc_id=c.doc_id WHERE 1=1{where}", args)}
            mask = np.isin(eids, list(allowed)) if allowed else np.zeros(len(eids), bool)
            sims = np.where(mask, sims, -1)
        top = np.argsort(-sims)[:60]
        for r, i in enumerate(top):
            if sims[i] > 0:
                ranked[int(eids[i])] += 1.0 / (RRF + r + 1)

    if not ranked:
        return []
    ids = list(ranked)
    q = ",".join("?" * len(ids))
    meta = {r[0]: r for r in c.execute(
        f"SELECT c.id, c.doc_id, c.kind, c.sec_id, c.sec_path, c.page_start, c.page_end, c.text, d.title, d.agency, "
        f"d.collection FROM chunks c JOIN docs d ON d.doc_id=c.doc_id WHERE c.id IN ({q})", ids)}
    qterms = {w.lower() for w in re.findall(r"[A-Za-z]{4,}", query)}
    numeric = bool(NUMERIC_INTENT.search(query))
    for cid in ids:
        m = meta.get(cid)
        if not m:
            continue
        head = (m[4] or "").lower()
        ranked[cid] += 0.004 * sum(1 for t in qterms if t in head)
        if numeric and m[2] == "table":
            ranked[cid] += 0.006
    order = sorted(ids, key=lambda i: -ranked[i])
    # 4b. cross-encoder rerank of the fused top 30 (query vs breadcrumb+text)
    if rerank and len(order) > 1:
        try:
            ce = _reranker_get()
            cand = [i for i in order[:30] if i in meta]
            scores = ce.predict([(query, f"{meta[i][4]}\n{meta[i][7][:1200]}") for i in cand], batch_size=16)
            pos = {i: r for r, i in enumerate(sorted(cand, key=lambda i: -float(scores[cand.index(i)])))}
            for i in cand:
                ranked[i] = 1.0 / (10 + pos[i]) + ranked[i] * 0.1
            order = sorted(ids, key=lambda i: -ranked[i])
        except Exception:
            pass
    out, seen_sec = [], Counter()
    for cid in order:
        m = meta.get(cid)
        if not m:
            continue
        key = (m[1], m[3])
        if seen_sec[key] >= 2:
            continue
        seen_sec[key] += 1
        hit = {"doc": m[1], "title": m[8], "agency": m[9], "collection": m[10], "kind": m[2],
               "section": m[3], "path": m[4], "pages": f"{m[5]}-{m[6]}" if m[6] != m[5] else str(m[5]),
               "score": round(ranked[cid], 5), "text": m[7][:1800],
               "cite": f"{m[8]} — {m[4].split(' > ')[-1][:80] if m[4] else m[3]} (p.{m[5]})"}
        if expand_section:
            hit["section_text"] = section_text(m[1], m[3], max_chars=8000)
        if follow_refs:
            hit["refs"] = [f"{t}:{v}" for t, v in c.execute(
                "SELECT DISTINCT ref_type, target FROM refs WHERE chunk_id=? LIMIT 12", (cid,))]
        out.append(hit)
        if len(out) >= k:
            break
    return out


def section_text(doc: str, sec_id: str, max_chars: int = 12000) -> str:
    c = connect()
    did = next((r[0] for r in c.execute("SELECT doc_id FROM docs WHERE doc_id LIKE ? LIMIT 1", (f"%{doc}%",))), None)
    if not did:
        return ""
    rows = c.execute("SELECT text FROM chunks WHERE doc_id=? AND (sec_id=? OR sec_id LIKE ?) AND kind!='table' "
                     "ORDER BY page_start, id", (did, sec_id, sec_id + ".%")).fetchall()
    txt, last = [], ""
    for (t,) in rows:
        if last and t[:CHUNK_OVERLAP // 2] in last:
            t = t[CHUNK_OVERLAP:]
        txt.append(t)
        last = t
    return "\n".join(txt)[:max_chars]


def read_pages(doc: str, p1: int, p2: int | None = None, max_chars: int = 15000) -> str:
    import fitz
    c = connect()
    row = c.execute("SELECT path FROM docs WHERE doc_id LIKE ? LIMIT 1", (f"%{doc}%",)).fetchone()
    if not row:
        return ""
    d = fitz.open(os.path.join(ROOT, row[0]))
    p2 = p2 or p1
    return "\n".join(f"--- p.{i}\n" + d[i - 1].get_text() for i in range(max(1, p1), min(d.page_count, p2) + 1))[:max_chars]


def toc(doc: str, max_level: int = 3) -> list[dict]:
    c = connect()
    did = next((r[0] for r in c.execute("SELECT doc_id FROM docs WHERE doc_id LIKE ? LIMIT 1", (f"%{doc}%",))), None)
    if not did:
        return []
    return [{"section": r[0], "title": r[1], "level": r[2], "pages": f"{r[3]}-{r[4]}"}
            for r in c.execute("SELECT sec_id, title, level, page_start, page_end FROM sections WHERE doc_id=? "
                               "AND level<=? ORDER BY page_start, level", (did, max_level))]


def list_docs(agency: str | None = None, collection: str | None = None) -> list[dict]:
    c = connect()
    where, args = [], []
    if agency:
        where.append("agency=?"); args.append(agency)
    if collection:
        where.append("collection=?"); args.append(collection)
    q = "SELECT doc_id, title, agency, collection, pages, n_chunks FROM docs" + (" WHERE " + " AND ".join(where) if where else "")
    return [dict(zip(("doc", "title", "agency", "collection", "pages", "chunks"), r)) for r in c.execute(q + " ORDER BY agency, collection, doc_id", args)]


def stats() -> dict:
    c = connect()
    return {"docs": c.execute("SELECT count(*) FROM docs").fetchone()[0],
            "pages": c.execute("SELECT coalesce(sum(pages),0) FROM docs").fetchone()[0],
            "chunks": c.execute("SELECT count(*) FROM chunks").fetchone()[0],
            "tables": c.execute("SELECT count(*) FROM chunks WHERE kind='table'").fetchone()[0],
            "sheet_pages": c.execute("SELECT count(*) FROM chunks WHERE kind='sheet'").fetchone()[0],
            "embedded": int(__import__("numpy").load(EMB_IDS_PATH).shape[0]) if os.path.exists(EMB_IDS_PATH) else 0}


def main() -> int:
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build"); b.add_argument("--no-embed", action="store_true"); b.add_argument("--only")
    b.add_argument("--force", action="store_true", help="re-index even if unchanged (e.g. after adding OCR)")
    s = sub.add_parser("search"); s.add_argument("query"); s.add_argument("--agency"); s.add_argument("--collection")
    s.add_argument("--doc"); s.add_argument("-k", type=int, default=6); s.add_argument("--no-dense", action="store_true")
    sc = sub.add_parser("section"); sc.add_argument("doc"); sc.add_argument("sec")
    t = sub.add_parser("toc"); t.add_argument("doc"); t.add_argument("--level", type=int, default=2)
    sub.add_parser("stats")
    sub.add_parser("list")
    a = ap.parse_args()
    if a.cmd == "build":
        print(json.dumps(build(embed=not a.no_embed, only=a.only, force=a.force), indent=1))
    elif a.cmd == "search":
        for h in search(a.query, a.agency, a.collection, a.doc, a.k, dense=not a.no_dense):
            print(f"[{h['score']}] {h['cite']}  <{h['kind']}> {h['doc']}\n    {h['path'][:150]}\n    "
                  + h["text"][:400].replace("\n", " ") + "\n")
    elif a.cmd == "section":
        print(section_text(a.doc, a.sec))
    elif a.cmd == "toc":
        for r in toc(a.doc, a.level):
            print("  " * (r["level"] - 1) + f"{r['section']}  {r['title'][:80]}  p.{r['pages']}")
    elif a.cmd == "stats":
        print(json.dumps(stats(), indent=1))
    elif a.cmd == "list":
        for d in list_docs():
            print(d)
    return 0


if __name__ == "__main__":
    sys.exit(main())
