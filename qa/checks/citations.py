"""STD-CITE-01: every cited standard requirement is real and current.

Task card `citations`: [{doc: "hdm/chapt_08.pdf", page: 94, quote: "minimum of 8 in. below the top of grate"}]
PASS when the quote (whitespace/case-insensitive, >= 12 chars) is found on the cited page
(+/- 1 page) of an indexed document. A doc missing from the index -> FAIL.
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "mcp-server"))


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def verify(task, ctx):
    cites = task.get("citations") or []
    if not cites:
        if task.get("type") == "engineering.calc":
            return {"status": "FAIL", "observed": "no citations", "expected": ">=1 governing criterion citation",
                    "evidence": [], "fix": "Use search_standards, then qa_gate cite --doc ... --page N --quote \"...\""}
        return {"status": "NA", "observed": "no citations in task", "expected": "", "evidence": [], "fix": ""}
    import standards_index as si
    bad, ok = [], []
    for c in cites:
        q = _norm(c.get("quote", ""))
        if len(q) < 12:
            bad.append(f"{c.get('doc')} p{c.get('page')}: quote too short to verify")
            continue
        p = int(c.get("page") or 0)
        txt = _norm(si.read_pages(c.get("doc", ""), max(1, p - 1), p + 1, max_chars=200000))
        if not txt:
            bad.append(f"{c.get('doc')}: not in standards index")
        elif q in txt:
            ok.append(f"{c.get('doc')} p{p}: verified")
        else:
            bad.append(f"{c.get('doc')} p{p}: quote not found on page")
    if bad:
        return {"status": "FAIL", "observed": f"{len(bad)}/{len(cites)} citations unverified", "expected": "all verified",
                "evidence": bad + ok, "fix": "Re-read the cited page (read_standard_pages) and correct doc/page/quote."}
    return {"status": "PASS", "observed": f"{len(ok)} citations verified on page", "expected": "", "evidence": ok, "fix": ""}
