"""Deterministic CAD-standards checks over element records (pure logic, no COM).

Element record = one row from wztc_ops.get_element_cad_props (strings):
  elementId, type, level, color, weight, lineStyle, lvlColor, lvlWeight, lvlStyle,
  lvlFromLibrary, cellName, textHeight, elClass, zLow, zHigh, rangeLowX/Y, rangeHighX/Y

ctx keys used:
  elements      list[dict]   new elements for this task (required)
  level_defs    dict name -> {color, weight, style, fromLibrary}  (optional; else row lvl* fields)
  rules         qa/project_rules.yaml content
  is2d          bool (default True)
  extent        [lowX, lowY, highX, highY] task work area (optional)
  extent_margin ft (default 50)
  text_heights  list[float] allowed text heights (optional)
Return per check: {status, observed, expected, evidence, fix}
"""
from __future__ import annotations

import re
from collections import defaultdict

MAX_IDS = 25


def _els(ctx):
    """New elements minus internal helper geometry (project_rules.ignore_levels, e.g. OpenRoads _CIVIL_*)."""
    pats = [re.compile(p, re.I) for p in (ctx.get("rules") or {}).get("ignore_levels", [])]
    return [e for e in ctx.get("elements") or [] if not any(p.search(e.get("level") or "") for p in pats)]


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _res(status, observed="", expected="", evidence=None, fix=""):
    return {"status": status, "observed": observed, "expected": expected,
            "evidence": (evidence or [])[:MAX_IDS], "fix": fix}


def _rule_for(el, rules):
    for r in (rules or {}).get("rules", []):
        if not re.search(r["level"], el.get("level") or "", re.I):
            continue
        if r.get("type") and r["type"].lower() not in (el.get("type") or "").lower():
            continue
        return r
    return None


def _lvl(el, ctx):
    d = (ctx.get("level_defs") or {}).get(el.get("level") or "")
    if d:
        return d
    return {"color": el.get("lvlColor", ""), "weight": el.get("lvlWeight", ""),
            "style": el.get("lvlStyle", ""), "fromLibrary": el.get("lvlFromLibrary", "")}


def new_elements(task, ctx):
    n = len(ctx.get("elements") or [])
    if ctx.get("elements_error"):
        return _res("UNCERTAIN", ctx["elements_error"], ">=1 new element read back",
                    fix="Open MicroStation with WZTCBridge polling, then re-run qa_gate run.")
    if n == 0:
        return _res("UNCERTAIN", "0 new elements since task baseline", ">=1",
                    fix="Nothing was drawn after `qa_gate start`, or the baseline was taken after drafting. "
                        "Re-run start before drafting, or pass --ids.")
    return _res("PASS", f"{n} new elements", ">=1")


def levels(task, ctx):
    els = _els(ctx)
    approved = set((ctx.get("rules") or {}).get("approved_levels", []))
    bad, unknown = [], []
    for el in els:
        lv = el.get("level") or ""
        if not lv and "cell" in (el.get("type") or "").lower() and el.get("childLevels"):
            # a cell header has no level of its own: its content levels (from the cell definition) are what count
            lib = ctx.get("level_library")
            for child in str(el["childLevels"]).split("|"):
                if child and child not in approved and lib is not None and child not in lib:
                    bad.append(f"{el['elementId']}:{child}(in cell {el.get('cellName', '')})")
            continue
        if not lv:
            bad.append(f"{el['elementId']}:no-level")
            continue
        if lv in approved or _rule_for(el, ctx.get("rules")):
            continue
        frm = str(_lvl(el, ctx).get("fromLibrary", "")).strip().lower()
        if frm in ("true", "-1", "1"):
            continue
        if frm in ("false", "0"):
            bad.append(f"{el['elementId']}:{lv}")
            continue
        # Level.IsFromLibrary is not exposed by MicroStation VBA (live 2026-10-23: blank) -
        # fall back to the NYSDOT level-library snapshot (Standards/derived/nysdot-levels.tsv).
        lib = ctx.get("level_library")
        if lib is not None:
            if lv not in lib:
                bad.append(f"{el['elementId']}:{lv}")
        else:
            unknown.append(f"{el['elementId']}:{lv}")
    if bad:
        return _res("FAIL", f"{len(bad)} elements on non-library levels", "NYSDOT library or approved level",
                    bad, "Move these elements to the correct NYSDOT level (list_levels / HDM Ch 20 Exhibit 20-5).")
    if unknown:
        return _res("UNCERTAIN", f"library flag unreadable for {len(unknown)} elements",
                    "level library membership known", unknown,
                    "Run qa_gate run with MicroStation open so LIST_LEVEL_DEFS can resolve these levels.")
    return _res("PASS", f"{len(els)} elements on library/approved levels")


def symbology(task, ctx):
    els = _els(ctx)
    tokens = set((ctx.get("rules") or {}).get("bylevel_tokens", []))
    bad, unknown = [], []
    for el in els:
        if not (el.get("level") or "") and "cell" in (el.get("type") or "").lower() and el.get("childLevels"):
            continue  # cell header: symbology comes from the cell's children (ByLevel in the NYSDOT libraries)
        col, wt = (el.get("color") or "").strip(), (el.get("weight") or "").strip()
        rule = _rule_for(el, ctx.get("rules"))
        if rule:
            if col == "" or wt == "":
                unknown.append(el["elementId"])
                continue
            lv = _lvl(el, ctx)
            # ByLevel resolves to the level's own symbology (live 2026-10-23: Default ByLevel = 0/0)
            eff_c = _f(lv.get("color")) if col in tokens else _f(col)
            eff_w = _f(lv.get("weight")) if wt in tokens else _f(wt)
            if eff_c != rule["color"] or eff_w != rule["weight"]:
                bad.append(f"{el['elementId']}:{el.get('level')} c{col}/w{wt} want c{rule['color']}/w{rule['weight']}")
            continue
        if col in tokens and wt in tokens:
            continue
        lv = _lvl(el, ctx)
        lc, lw = str(lv.get("color", "")).strip(), str(lv.get("weight", "")).strip()
        if "" in (col, wt, lc, lw):
            unknown.append(el["elementId"])
            continue
        ok_c = col in tokens or _f(col) == _f(lc)
        ok_w = wt in tokens or _f(wt) == _f(lw)
        if not (ok_c and ok_w):
            bad.append(f"{el['elementId']}:{el.get('level')} c{col}/w{wt} level c{lc}/w{lw}")
    if bad:
        return _res("FAIL", f"{len(bad)} elements off-symbology", "project table or ByLevel", bad,
                    "Set color/weight to ByLevel (or the CLAUDE.md table value) on these elements.")
    if unknown:
        return _res("UNCERTAIN", f"symbology unreadable for {len(unknown)} elements", "readable color/weight",
                    unknown, "Bridge returned blank color/weight; check GET_ELEMENT_CAD_PROPS after hot-reload.")
    return _res("PASS", f"{len(els)} elements on correct symbology")


def proposed_existing_weight(task, ctx):
    bad = []
    for el in _els(ctx):
        lv, wt = el.get("level") or "", _f(el.get("weight"))
        if wt is None or _rule_for(el, ctx.get("rules")):
            continue
        if re.search(r"_P$", lv, re.I) and wt == 0:
            bad.append(f"{el['elementId']}:{lv} w0")
        if re.search(r"_E$", lv, re.I) and wt not in (0, None):
            bad.append(f"{el['elementId']}:{lv} w{wt:g}")
    if bad:
        return _res("WARN", f"{len(bad)} proposed/existing weight mismatches", "_P weight 2, _E weight 0", bad,
                    "HDM §20.7.1: proposed typically weight 2, existing weight 0 - confirm with level definition.")
    return _res("PASS", "proposed/existing weights consistent")


def elevation(task, ctx):
    if not ctx.get("is2d", True):
        return _res("NA", "3D model - elevation rule not applied")
    bad, unknown = [], []
    for el in _els(ctx):
        zl, zh = _f(el.get("zLow")), _f(el.get("zHigh"))
        if zl is None or zh is None:
            unknown.append(el["elementId"])
        elif abs(zl) > 0.001 or abs(zh) > 0.001:
            bad.append(f"{el['elementId']}:z={zl:g}..{zh:g}")
    if bad:
        return _res("FAIL", f"{len(bad)} elements off Z=0", "Z = 0 in 2D", bad,
                    "Flatten these elements to Z=0 (HDM §20.5.4.2.b).")
    if unknown:
        return _res("UNCERTAIN", f"Z unreadable for {len(unknown)}", "Z known", unknown)
    return _res("PASS", "all new elements at Z=0")


def location(task, ctx):
    ext = ctx.get("extent")
    if not ext:
        return _res("UNCERTAIN", "no task extent", "task card `extent` or baseline extent",
                    fix="Pass --extent lowX,lowY,highX,highY at qa_gate start (work area), or let start record it.")
    m = float(ctx.get("extent_margin", 50))
    lx, ly, hx, hy = ext[0] - m, ext[1] - m, ext[2] + m, ext[3] + m
    bad = []
    for el in _els(ctx):
        a, b, c, d = (_f(el.get(k)) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
        if None in (a, b, c, d):
            continue
        if c < lx or a > hx or d < ly or b > hy:
            bad.append(f"{el['elementId']}:({a:.0f},{b:.0f})")
    if bad:
        return _res("FAIL", f"{len(bad)} elements outside work extent (+{m:g} ft)", f"inside {ext}", bad,
                    "Stray elements - delete or move them; check coordinate input / ACS.")
    return _res("PASS", "all new elements inside work extent")


def duplicates(task, ctx):
    seen = defaultdict(list)
    for el in _els(ctx):
        key = (el.get("type"), el.get("level"), el.get("cellName"),
               *[round(_f(el.get(k)) or 0, 2) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY")])
        seen[key].append(el["elementId"])
    dups = [",".join(v) for v in seen.values() if len(v) > 1]
    if dups:
        return _res("FAIL", f"{len(dups)} duplicate groups", "no stacked duplicates", dups,
                    "Delete the extra copies (keep one per group).")
    return _res("PASS", "no duplicates")


def degenerate(task, ctx):
    bad = []
    for el in _els(ctx):
        t = (el.get("type") or "").lower()
        if any(k in t for k in ("text", "cell", "point")):
            continue
        a, b, c, d = (_f(el.get(k)) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
        if None in (a, b, c, d):
            continue
        if abs(c - a) < 1e-4 and abs(d - b) < 1e-4:
            bad.append(el["elementId"])
    if bad:
        return _res("FAIL", f"{len(bad)} zero-size elements", "no degenerate geometry", bad, "Delete them.")
    return _res("PASS", "no degenerate geometry")


def cells(task, ctx):
    bad = [el["elementId"] for el in _els(ctx)
           if "cell" in (el.get("type") or "").lower() and not (el.get("cellName") or "").strip()]
    if bad:
        return _res("FAIL", f"{len(bad)} unnamed cells", "named library cells", bad,
                    "Re-place from the NYSDOT cell library (attach_cell_library / place_cell).")
    return _res("PASS", "cells named")


def text_heights(task, ctx):
    allowed = _allowed_heights(ctx)
    if not allowed:
        return _res("NA", "no text-height table in task card")
    bad = []
    for el in _els(ctx):
        if "text" not in (el.get("type") or "").lower():
            continue
        h = _f(el.get("textHeight"))
        if h is None:
            bad.append(f"{el['elementId']}:unreadable")
        elif not any(abs(h - a) <= 0.01 * a for a in allowed):
            bad.append(f"{el['elementId']}:{h:g}")
    if bad:
        return _res("FAIL", f"{len(bad)} text heights off table", f"one of {allowed}", bad,
                    "Use the NYSDOT text style for this scale (HDM §20.9.6 / §20.9.8).")
    return _res("PASS", "text heights match scale table")


def _is_text(e):
    return "text" in (e.get("type") or "").lower()


def _style(e):
    return (_f(e.get("textHeight")), _f(e.get("textWidth")), (e.get("textFont") or "").strip().lower())


def _bold(e):
    """True / False from the textBold column, None when the model did not report it."""
    v = str(e.get("textBold") if e.get("textBold") is not None else "").strip()
    return True if v == "1" else False if v == "0" else None


def _allowed_heights(ctx):
    """Task text-height table, in the units the read-back reports. Annotative text reads back in paper units
    (0.00521 at annotation scale 1200 = 6.252 ft), so a table given in design feet is also accepted divided by the scale."""
    table = [float(a) for a in ctx.get("text_heights") or []]
    sc = _f(((ctx.get("model_state") or {}).get("annotationScaleFactor")))
    if sc and sc > 1:
        table += [a / sc for a in table if a > 0.1]
    return table


def _box(e):
    return tuple(_f(e.get(k)) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))


def text_style(task, ctx):
    """New text must use a height / width / font that already exists on the sheet's text.

    Incident 2026-10-23: new labels at 0.003 x 0.010 sat beside sheet notes at 0.007 x 0.007 and the
    gate passed because text size was never read back."""
    new = [e for e in _els(ctx) if _is_text(e)]
    if not new:
        return _res("NA", "no new text")
    if ctx.get("reference_texts_error"):
        return _res("UNCERTAIN", f"sheet text unreadable: {ctx['reference_texts_error']}", "reference text read back",
                    fix="Open MicroStation with the bridge running and re-run.")
    unread = [e["elementId"] for e in new if None in _style(e)[:2] or not _style(e)[2]]
    if unread:
        return _res("UNCERTAIN", f"text height/width/font unreadable on {len(unread)} elements", "readable text style",
                    unread, "Reload WZTCQuery.bas (GET_ELEMENT_CAD_PROPS text columns) and re-run.")
    refs = {(round(h, 4), round(w, 4), f, _bold(e), (e.get("level") or "")) for e in ctx.get("reference_texts") or []
            for h, w, f in [_style(e)] if h is not None and w is not None}
    allowed = _allowed_heights(ctx)
    if not refs and not allowed:
        return _res("UNCERTAIN", "no existing text on the sheet and no text-height table", "a style authority",
                    fix="Give the task a text style: set text_heights on the task card or match an engineer-named sheet note.")
    bad = []
    for e in new:
        h, w, f = _style(e)
        b = _bold(e)
        # the sheet's own convention for THIS level is the authority when it has one (2026-10-25: regular labels matched
        # some other level's text and passed); otherwise any text on the sheet
        lvl = e.get("level") or ""
        pool = [r for r in refs if r[4] == lvl] or list(refs)
        ok = any(abs(h - rh) <= 0.02 * rh and abs(w - rw) <= 0.02 * rw and f == rf and (b is None or rb is None or b == rb)
                 for rh, rw, rf, rb, _lv in pool)
        if not ok and allowed:
            ok = any(abs(h - a) <= 0.02 * a for a in allowed) and abs(w - h) <= 0.02 * h
        if not ok:
            bad.append(f"{e['elementId']}: h{h:g} w{w:g} {f} {'bold' if b else 'regular' if b is not None else 'bold?'} "
                       f"'{(e.get('text') or '')[:24]}'")
    if bad:
        have = sorted(f"h{h:g} w{w:g} {f} {'bold' if b else 'regular' if b is not None else 'bold?'}" for h, w, f, b, _lv in refs)[:6]
        return _res("FAIL", f"{len(bad)} new text elements do not match any text style on the sheet",
                    f"one of {have}", bad,
                    "Copy height/width/font from the nearest existing note of the same kind (read_standard / GET_ELEMENT_CAD_PROPS ids=TEXT), "
                    "then re-place. Do not shrink text to fit clutter; move it.")
    return _res("PASS", f"{len(new)} new text elements match existing sheet text styles")


def text_overlap(task, ctx):
    """New text boxes must not overlap other new text or the sheet's existing text."""
    new = [e for e in _els(ctx) if _is_text(e)]
    if not new:
        return _res("NA", "no new text")
    others = [(e, False) for e in ctx.get("reference_texts") or []]
    hits = []
    for i, a in enumerate(new):
        ba = _box(a)
        pool = others + [(b, True) for b in new[i + 1:]]
        for b, _ in pool:
            bb = _box(b)
            if None in ba or None in bb:
                continue
            if _boxes_overlap(ba, bb):
                hits.append(f"{a['elementId']} overlaps {b['elementId']} '{(b.get('text') or '')[:20]}'")
    if hits:
        return _res("FAIL", f"{len(hits)} text overlaps", "no text-on-text overlap", hits, "Move the new text clear of the other text.")
    return _res("PASS", f"{len(new)} new text boxes clear of other text")


def _boxes_overlap(a, b):
    """True when two text boxes overlap by more than 5% of the smaller box (edge touches of rotated text ranges do not count)."""
    ix = min(a[2], b[2]) - max(a[0], b[0])
    iy = min(a[3], b[3]) - max(a[1], b[1])
    if ix <= 0 or iy <= 0:
        return False
    small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    return ix * iy > 0.05 * small


def text_model_clear(task, ctx):
    """No two text boxes in the active model overlap, new or existing (including ORD annotation nodes and text in cells).

    Incident 2026-10-25: the CURVE: CREST node overlapped the VHP node; both were ORD-made, so CAD-TXT-03
    (new text against other text) never looked at the pair. Boxes that merely touch at the edge (rotated text
    ranges) are ignored: the overlap must cover more than 5% of the smaller box."""
    if ctx.get("reference_texts_error"):
        return _res("UNCERTAIN", f"model text unreadable: {ctx['reference_texts_error']}", "all text read back",
                    fix="Open MicroStation with the bridge running and re-run.")
    texts = ctx.get("model_texts")
    if texts is None:
        return _res("NA", "model text not read")
    boxes = []
    for e in texts:
        b = _box(e)
        if None in b or b[2] <= b[0] or b[3] <= b[1]:
            continue
        boxes.append((str(e.get("elementId")), (e.get("text") or "").strip()[:24], b))
    hits = []
    for i in range(len(boxes)):
        ida, ta, a = boxes[i]
        for j in range(i + 1, len(boxes)):
            idb, tb, b = boxes[j]
            if ida.split(".")[0] == idb.split(".")[0] and "." in ida and "." in idb:
                continue                                    # two lines of the same cell
            ix = min(a[2], b[2]) - max(a[0], b[0])
            iy = min(a[3], b[3]) - max(a[1], b[1])
            if ix <= 0 or iy <= 0:
                continue
            small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
            if ix * iy > 0.05 * small:
                hits.append(f"{ida} '{ta}' overlaps {idb} '{tb}' ({ix:.1f} x {iy:.1f})")
    if hits:
        return _res("FAIL", f"{len(hits)} overlapping text pairs in the model", "no text box overlaps another", hits,
                    "Move one of the texts clear (HDM 20.5.1.1 h: text must not overlap so it becomes unreadable); "
                    "for ORD annotation nodes move the node, not the leader.")
    if ctx.get("model_texts_truncated"):
        return _res("UNCERTAIN", "text read-back truncated at the row cap: overlaps beyond it were not checked", "all text read",
                    fix="Check a smaller area or raise the read cap.")
    return _res("PASS", f"{len(boxes)} text boxes in the model, none overlap")


def _is_linear(e):
    t = (e.get("type") or "").lower()
    if "text" in t or "cell" in t or "dimension" in t:
        return False
    return any(k in t for k in ("line", "curve", "arc", "shape", "string", "complex"))


def _pt_in_box(x, y, box, pad=0.0):
    return box[0] - pad <= x <= box[2] + pad and box[1] - pad <= y <= box[3] + pad


def _seg_crosses(ax, ay, bx, by, cx, cy, dx, dy):
    """True if segment AB properly intersects CD (including touching)."""
    def cross(ox, oy, px, py, qx, qy):
        return (px - ox) * (qy - oy) - (py - oy) * (qx - ox)

    d1 = cross(cx, cy, dx, dy, ax, ay)
    d2 = cross(cx, cy, dx, dy, bx, by)
    d3 = cross(ax, ay, bx, by, cx, cy)
    d4 = cross(ax, ay, bx, by, dx, dy)
    if ((d1 > 0 and d2 < 0) or (d1 < 0 and d2 > 0)) and ((d3 > 0 and d4 < 0) or (d3 < 0 and d4 > 0)):
        return True
    # endpoint on segment counts as overlap
    for x, y, p, q, r, s in (
        (ax, ay, cx, cy, dx, dy), (bx, by, cx, cy, dx, dy),
        (cx, cy, ax, ay, bx, by), (dx, dy, ax, ay, bx, by),
    ):
        if abs(cross(p, q, r, s, x, y)) > 1e-9:
            continue
        if min(p, r) - 1e-9 <= x <= max(p, r) + 1e-9 and min(q, s) - 1e-9 <= y <= max(q, s) + 1e-9:
            return True
    return False


def _seg_hits_box(x0, y0, x1, y1, box, pad=0.0):
    """True if open segment intersects axis-aligned box expanded by pad.

    Uses the line's actual endpoints — not the element's range AABB (a diagonal
    line's range box is huge and would false-positive every nearby label).
    """
    b = (box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad)
    if _pt_in_box(x0, y0, b) or _pt_in_box(x1, y1, b):
        return True
    corners = ((b[0], b[1]), (b[2], b[1]), (b[2], b[3]), (b[0], b[3]))
    for i in range(4):
        cx, cy = corners[i]
        dx, dy = corners[(i + 1) % 4]
        if _seg_crosses(x0, y0, x1, y1, cx, cy, dx, dy):
            return True
    return False


def text_line_clearance(task, ctx):
    """New text must not sit on new linear geometry (segment vs text box).

    Incident 2026-10-23: CAD-TXT-03 only checked text-on-text; RW labels sat on
    TOW lines and the gate still PASSed. Engineer correction: check text vs lines.
    """
    new_txt = [e for e in _els(ctx) if _is_text(e)]
    if not new_txt:
        return _res("NA", "no new text")
    lines = [e for e in _els(ctx) if _is_linear(e)]
    if not lines:
        return _res("NA", "no new linear elements")
    paths = ctx.get("line_paths") or {}
    missing = [e["elementId"] for e in lines if str(e["elementId"]) not in paths
               or len(paths[str(e["elementId"])]) < 2]
    if missing and not paths:
        return _res("UNCERTAIN", "line vertices not available for text/line clearance",
                    "path vertices for new lines",
                    missing[:MAX_IDS],
                    "Re-run with MicroStation open so get_element_vertices can load line paths.")
    pad = float((ctx.get("rules") or {}).get("text_line_pad", 1.0))
    hits = []
    for te in new_txt:
        box = _box(te)
        if None in box:
            continue
        # HDM 20.5.1.1(i): put text beside the feature and point to it with a leader line. A line running along the edge of
        # the text block is that leader, so only a line that cuts at least `pad` into the block counts as text on a line.
        if box[2] - box[0] > 2 * pad and box[3] - box[1] > 2 * pad:
            box = (box[0] + pad, box[1] + pad, box[2] - pad, box[3] - pad)
        for ln in lines:
            pts = paths.get(str(ln["elementId"])) or []
            if len(pts) < 2:
                continue
            for i in range(len(pts) - 1):
                x0, y0 = float(pts[i][0]), float(pts[i][1])
                x1, y1 = float(pts[i + 1][0]), float(pts[i + 1][1])
                if _seg_hits_box(x0, y0, x1, y1, box, pad=0.0):
                    hits.append(f"{te['elementId']} '{(te.get('text') or '')[:24]}' on line {ln['elementId']}")
                    break
    if hits:
        return _res("FAIL", f"{len(hits)} text/line overlaps", "text clear of new linework", hits,
                    "Move the text off the line (offset perpendicular to the segment). "
                    "Do not rely on the line's range box — use the actual segment.")
    checked = sum(1 for e in lines if len(paths.get(str(e["elementId"]), []) or []) >= 2)
    note = f" ({len(missing)} lines without vertices skipped)" if missing else ""
    return _res("PASS", f"{len(new_txt)} text boxes clear of {checked} lines{note}")


# ---------------------------------------------------------------- tutorial-derived rules (Data/tutorial-playbooks)
_BAR = re.compile(r"^(\d+)_(?:1|Vert)(?:\s|$)")


def scale_bar_scale(task, ctx):
    """A graphic scale bar cell `<N>_1` / `<N>_Vert` (ny_barscale.cel) must match the model's annotation scale 1:N.
    Playbooks: plan-sheets-named-boundaries S8, profile-sheet-annotation S6."""
    bars = [(e, _BAR.match(e.get("cellName") or "")) for e in _els(ctx) if "cell" in (e.get("type") or "").lower()]
    bars = [(e, m) for e, m in bars if m]
    if not bars:
        return _res("NA", "no scale bar cells")
    st = ctx.get("model_state") or {}
    ann = _f(st.get("annotationScaleFactor"))
    if ann is None:
        return _res("UNCERTAIN", "model annotation scale unreadable", "annotationScaleFactor",
                    fix="Run with the target Drawing model active so describe_drawing_state can read its annotation scale.")
    bad = [f"{e['elementId']}: cell {e['cellName']} is 1:{m.group(1)} but the model annotation scale is 1:{ann:g}"
           for e, m in bars if abs(float(m.group(1)) - ann) > 0.5]
    if bad:
        return _res("FAIL", f"{len(bad)} scale bars do not match the annotation scale", f"cell for 1:{ann:g}", bad,
                    "Use the bar cell named for this scale, e.g. 240_1 (1in=20ft), 480_1 / 480_Vert (1in=40ft), 600_1 (1in=50ft).")
    return _res("PASS", f"{len(bars)} scale bars match annotation scale 1:{ann:g}")


_ANNOT = ("text", "cell", "dimension")


def annotation_model(task, ctx):
    """Annotation belongs in the Drawing model; the Sheet model only references it (ORD manual p.468; BRP Profiles)."""
    st = ctx.get("model_state") or {}
    if str(st.get("modelType", "")) != "1":  # 1 = sheet model
        return _res("NA", "active model is not a Sheet model")
    hit = [e for e in _els(ctx) if any(k in (e.get("type") or "").lower() for k in _ANNOT)]
    if hit:
        return _res("WARN", f"{len(hit)} annotation elements were placed in a SHEET model ({st.get('modelName')})",
                    "annotation placed in the Drawing model", [f"{e['elementId']} {e.get('type')}" for e in hit],
                    "Place text, cells and dimensions in the Drawing model so they follow the sheet; sheet-model annotation is only for sheet-specific content.")
    return _res("PASS", "no annotation placed in the sheet model")
