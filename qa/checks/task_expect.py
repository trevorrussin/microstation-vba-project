"""Task-specific expectations from the task card (`expect` rows).

Each expect row: {id, desc, where: {level?, type?, cellName?, textContains?}, count: ">=3" | "==2" | "<=0"}
Evaluated against ctx["elements"] (new elements). For non-CAD tasks a row may instead give
`value` + `target` comparison: {id, desc, value: <number from task results>, target: ">=0.5"}.
"""
from __future__ import annotations

import operator
import re

OPS = {">=": operator.ge, "<=": operator.le, "==": operator.eq, ">": operator.gt, "<": operator.lt, "!=": operator.ne}


def _cmp(expr: str, n: float) -> bool:
    m = re.match(r"\s*(>=|<=|==|!=|>|<)\s*(-?[\d.]+)\s*$", str(expr))
    if not m:
        raise ValueError(f"bad comparison {expr!r}")
    return OPS[m.group(1)](n, float(m.group(2)))


def _match(el: dict, where: dict) -> bool:
    for k, v in (where or {}).items():
        if k == "textContains":
            if v.lower() not in (el.get("text") or "").lower():
                return False
        elif not re.search(str(v), str(el.get(k) or ""), re.I):
            return False
    return True


def evaluate(task, ctx):
    rows = task.get("expect") or []
    if not rows:
        return {"status": "UNCERTAIN", "observed": "task card has no `expect` rows",
                "expected": "at least one measurable expectation from the standard/sheet/request",
                "evidence": [], "fix": "Add expectations: qa_gate expect --task ID --id E1 --desc ... --where level=TWZCD --count '>=1'"}
    fails, ev = [], []
    for r in rows:
        if r.get("rel"):  # geometric relations are measured by operations_check.relations (REQ-GEOM-01)
            continue
        try:
            if "count" in r:
                n = sum(1 for el in ctx.get("elements") or [] if _match(el, r.get("where")))
                ok = _cmp(r["count"], n)
                ev.append(f"{r['id']}: {n} {'ok' if ok else 'FAIL'} ({r['count']}) {r.get('desc','')[:60]}")
            else:
                val = (ctx.get("results") or {}).get(r["value"]) if isinstance(r.get("value"), str) else r.get("value")
                ok = val is not None and _cmp(r["target"], float(val))
                ev.append(f"{r['id']}: {val} {'ok' if ok else 'FAIL'} ({r['target']}) {r.get('desc','')[:60]}")
        except Exception as ex:
            ok = False
            ev.append(f"{r.get('id')}: ERROR {ex}")
        if not ok:
            fails.append(r.get("id"))
    if fails:
        return {"status": "FAIL", "observed": f"{len(fails)}/{len(rows)} expectations not met",
                "expected": "all expectations met", "evidence": ev, "fix": "Fix the drawing/result for: " + ", ".join(fails)}
    return {"status": "PASS", "observed": f"{len(rows)} expectations met", "expected": "", "evidence": ev, "fix": ""}
