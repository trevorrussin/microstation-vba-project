"""Visual QC: geometry render + independent vision review + agent review record.

Why a geometry render: screen captures of MicroStation have disagreed with what the
engineer saw (2026-08-10, GPU/PrintWindow). The render is drawn from element data read
back from the DGN/DWG, so it shows what is actually in the file. A screen capture is
added as a second image when available.

VIS-01 PASS requires, after the latest run, one of:
  * an independent vision review (Claude, fresh context, checklist, deterministic results, surrounding linework)
    with no open issues; or
  * the engineer's sign-off of VIS-01 (`qa_gate signoff`, interactive terminal, bound to the drawing state).
The agent's own `qa_gate review --ok` is recorded and its listed issues still block, but it can never PASS VIS-01:
a review the agent writes about its own work is not independent.
A vision issue is dismissed only by a dispute that cites deterministic checks that PASSED on this run:
`qa_gate review --dispute N "why" --evidence CAD-LVL-01,CAD-SYM-01`. A dispute without evidence, or citing a
check that did not pass, leaves the issue open. Reviewer issues that contradict a deterministic PASS but name no
element ids are rejected automatically as unsupported.
"""
from __future__ import annotations

import base64
import json
import os
import re
import time

VISION_MODEL = os.environ.get("QA_VISION_MODEL", "claude-opus-5-5")
VISION_TIMEOUT_S = float(os.environ.get("QA_VISION_TIMEOUT_S", "180"))
LEVEL_COLORS = ["#d62728", "#1f77b4", "#2ca02c", "#ff7f0e", "#9467bd", "#8c564b", "#e377c2", "#17becf", "#bcbd22"]


def _draw_text(ax, el, color, lw=0.0):
    """Draw the text string itself, scaled into its real range box and rotation, so size/aspect/position are visible."""
    from matplotlib.textpath import TextPath
    from matplotlib.patches import PathPatch
    from matplotlib.transforms import Affine2D
    s = (el.get("text") or "").strip()
    a, b, c, d = (float(el[k]) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
    if not s:
        return False
    tp = TextPath((0, 0), s, size=1)
    bb = tp.get_extents()
    if bb.width <= 0 or bb.height <= 0:
        return False
    rot = float(el.get("textRotDeg") or 0)
    vertical = abs(abs(rot) % 180 - 90) < 45
    length, thick = ((d - b), (c - a)) if vertical else ((c - a), (d - b))
    tf = (Affine2D().translate(-bb.x0, -bb.y0).scale(length / bb.width, thick / bb.height)
          .rotate_deg(90 if vertical else 0).translate(c if vertical else a, b if vertical else b))
    ax.add_patch(PathPatch(tp, transform=tf + ax.transData, fc=color, ec="none", lw=lw))
    return True


def _context_rows(context_rows, elements, extent, margin=150.0):
    """Existing elements near the work area (not the task's own), so the reviewer sees what the new work sits among."""
    new_ids = {str(e.get("elementId")) for e in elements}
    try:
        if extent:
            box = (extent[0] - margin, extent[1] - margin, extent[2] + margin, extent[3] + margin)
        else:
            xs = [float(e["rangeLowX"]) for e in elements] + [float(e["rangeHighX"]) for e in elements]
            ys = [float(e["rangeLowY"]) for e in elements] + [float(e["rangeHighY"]) for e in elements]
            box = (min(xs) - margin, min(ys) - margin, max(xs) + margin, max(ys) + margin)
    except Exception:
        return []
    out = []
    for e in context_rows or []:
        if str(e.get("elementId")) in new_ids:
            continue
        try:
            a, b, c, d = (float(e[k]) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
        except Exception:
            continue
        if c >= box[0] and a <= box[2] and d >= box[1] and b <= box[3]:
            out.append(e)
    return out


def render_elements(elements: list[dict], vertices_fn, out_png: str, title: str, extent=None, max_vertices=400,
                    reference_texts=None, context_rows=None, max_context=300) -> str:
    """Plot new elements: polylines from vertices (when vertices_fn works), else range boxes.
    context_rows (every element of the model) are drawn light gray behind them: existing linework near the work."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(11, 8.5))
    lv_color = {}
    n_ctx = 0
    for ce in _context_rows(context_rows, elements, extent):
        t = (ce.get("type") or "").lower()
        if "text" in t:
            continue  # existing text is drawn from reference_texts
        drawn = False
        if vertices_fn and n_ctx < max_context and not any(k in t for k in ("cell", "dimension")):
            try:
                v = vertices_fn(ce["elementId"])
                pts = v.get("path_vertices") or v.get("vertices") or []
                if len(pts) >= 2:
                    ax.plot([p[0] for p in pts], [p[1] for p in pts], color="#b8b8b8", lw=0.7, zorder=1)
                    drawn = True
                    n_ctx += 1
            except Exception:
                pass
        if not drawn:
            a, b, c, d = (float(ce[k]) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
            ax.add_patch(plt.Rectangle((a, b), max(c - a, 0.05), max(d - b, 0.05), fill=False, ec="#cccccc", lw=0.5, zorder=1))
    for rt in reference_texts or []:  # existing sheet text, gray, for style/position comparison
        try:
            if extent and not (extent[0] - 100 <= float(rt["rangeLowX"]) <= extent[2] + 100
                               and extent[1] - 100 <= float(rt["rangeLowY"]) <= extent[3] + 100):
                continue
            _draw_text(ax, rt, "#888888")
        except Exception:
            pass
    n_vert = 0
    for el in elements:
        lv = el.get("level") or (el.get("childLevels") or "").replace("|", " + ") or "?"  # a cell header has no level of its own
        col = lv_color.setdefault(lv, LEVEL_COLORS[len(lv_color) % len(LEVEL_COLORS)])
        try:
            a, b, c, d = (float(el[k]) for k in ("rangeLowX", "rangeLowY", "rangeHighX", "rangeHighY"))
        except Exception:
            continue
        drawn = False
        t = (el.get("type") or "").lower()
        if vertices_fn and n_vert < max_vertices and not any(k in t for k in ("text", "cell")):
            try:
                v = vertices_fn(el["elementId"])
                pts = v.get("path_vertices") or v.get("vertices") or []
                if len(pts) >= 2:
                    ax.plot([p[0] for p in pts], [p[1] for p in pts], color=col, lw=1.6)
                    drawn = True
                    n_vert += 1
            except Exception:
                pass
        if not drawn and "text" in t:
            try:
                drawn = _draw_text(ax, el, col)
                ax.add_patch(plt.Rectangle((a, b), c - a, d - b, fill=False, ec=col, lw=0.5, ls=":"))
            except Exception:
                drawn = False
        if not drawn:
            ax.add_patch(plt.Rectangle((a, b), max(c - a, 0.05), max(d - b, 0.05), fill="cell" in t,
                                       alpha=0.35 if "cell" in t else 1, ec=col, fc=col, lw=1))
        if "cell" in t:
            label = (el.get("cellName") or "")[:18]
            if label:
                ax.text(a, d, label, fontsize=6, color=col)
    if extent:
        ax.add_patch(plt.Rectangle((extent[0], extent[1]), extent[2] - extent[0], extent[3] - extent[1],
                                   fill=False, ls="--", ec="gray", lw=0.8))
    for lv, col in lv_color.items():
        ax.plot([], [], color=col, label=lv)
    if lv_color:
        ax.legend(fontsize=7, loc="upper right")
    ax.set_aspect("equal", adjustable="datalim")
    ax.autoscale_view()
    ax.set_title(title[:120], fontsize=9)
    os.makedirs(os.path.dirname(out_png), exist_ok=True)
    fig.savefig(out_png, dpi=120, bbox_inches="tight")
    plt.close(fig)
    return out_png


def _checklist(task: dict, failure_modes: list[dict], cad_rows: list[dict], det_results: dict | None = None) -> str:
    lines = [f"TASK: {task.get('desc','')}", f"TYPE: {task.get('type')}"]
    for rid, r in (det_results or {}).items():
        lines.append(f"- DETERMINISTIC {rid} = {r.get('status')}: {str(r.get('observed', ''))[:140]}")
    for e in task.get("expect") or []:
        lines.append(f"- EXPECT {e.get('id')}: {e.get('desc','')} {e.get('count') or e.get('target') or ''}")
    for r in cad_rows:
        lines.append(f"- RULE {r['id']}: {r['req']}")
    for f in failure_modes:
        lines.append(f"- KNOWN FAILURE {f['id']}: {f['mode']}")
    return "\n".join(lines)


def vision_review(images: list[str], task: dict, failure_modes: list[dict], cad_rows: list[dict], det_results: dict | None = None) -> dict:
    """Independent reviewer: fresh-context Claude call over the images + checklist. Returns
    {"ran": bool, "issues": [...], "looks_correct": bool, "note": str}."""
    try:
        import anthropic
    except ImportError:
        return {"ran": False, "note": "anthropic SDK not installed"}
    try:  # corporate TLS interception: use the OS cert store (same fix as chat_driver.py)
        import truststore
        truststore.inject_into_ssl()
    except Exception:
        pass
    try:  # same credential source as chat_driver.py
        from dotenv import load_dotenv
        load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "mcp-server", ".env"))
    except ImportError:
        pass
    if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        return {"ran": False, "note": "no ANTHROPIC_API_KEY / ANTHROPIC_AUTH_TOKEN"}
    content = []
    for p in images[:3]:
        with open(p, "rb") as f:
            content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                        "data": base64.standard_b64encode(f.read()).decode()}})
    content.append({"type": "text", "text": (
        "You are an independent CAD QC reviewer for NYSDOT / NYC DDC plan drafting. You did not make this "
        "drawing. The first image is a render of the elements the task created, drawn from data read back "
        "from the design file (colors = levels; a cell's legend entry is the levels of its contents). Light gray "
        "linework is EXISTING geometry near the work. A screen capture may follow.\n\n"
        + _checklist(task, failure_modes, cad_rows, det_results) +
        "\n\nLines marked DETERMINISTIC were measured from the CAD data itself (levels, symbology, extent, text "
        "style, overlaps). If you disagree with one, set contradicts_deterministic=true and name the element ids and "
        "the evidence you see; an issue that contradicts a deterministic PASS without element ids is discarded." +
        "\n\nFind real problems only: missing or extra elements vs the expectations, elements in the wrong "
        "place, overlaps/collisions, gaps, stray elements, wrong symbology/levels visible in the legend, "
        "labels or dimensions that would be unreadable or misplaced. Gray text in the render is existing sheet text: "
        "new (colored) text must look the same size and proportion as comparable gray text and must not collide with it. "
        "Do not report other style preferences.\n"
        "Reply with JSON only: {\"looks_correct\": true|false, \"issues\": [{\"severity\": \"major|minor\", "
        "\"description\": \"...\", \"where\": \"...\", \"elements\": [\"element ids\"], \"requirement\": \"rule id or standard\", "
        "\"evidence\": \"what in the image shows it\", \"contradicts_deterministic\": false}]}")})
    client = anthropic.Anthropic(timeout=anthropic.Timeout(VISION_TIMEOUT_S, connect=15.0), max_retries=1)
    try:
        msg = client.beta.messages.create(
            model=VISION_MODEL, max_tokens=4000,
            output_config={"effort": "medium"},
            betas=["server-side-fallback-2026-07-01"],
            extra_body={"fallbacks": "default"},
            messages=[{"role": "user", "content": content}])
    except anthropic.APIStatusError as ex:
        return {"ran": False, "note": f"vision API error {ex.status_code}: {str(ex)[:200]}"}
    except anthropic.APITimeoutError:
        return {"ran": False, "note": f"vision API timed out after {VISION_TIMEOUT_S:g} s (a timeout is not a clean review)"}
    except anthropic.APIConnectionError as ex:
        return {"ran": False, "note": f"vision API connection error: {ex}"}
    if msg.stop_reason == "refusal":
        return {"ran": False, "note": "vision reviewer refused"}
    text = "".join(getattr(b, "text", "") for b in msg.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", text, re.S)
    try:
        data = json.loads(m.group(0)) if m else {}
    except json.JSONDecodeError:
        data = {}
    if not data:
        return {"ran": False, "note": "vision reply not JSON: " + text[:200]}
    return {"ran": True, "issues": data.get("issues") or [], "looks_correct": bool(data.get("looks_correct")),
            "model": VISION_MODEL}


def _sha(path):
    import hashlib
    try:
        with open(path, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()
    except OSError:
        return ""


def _dispute_ok(d, results):
    """A dispute counts only if it cites deterministic checks that PASSED on this run. -> (accepted, why)"""
    if not isinstance(d, dict) or not d.get("evidence"):
        return False, "no evidence: a dispute must cite the deterministic checks that disprove the issue (--evidence CHECK_ID,...)"
    bad = [c for c in d["evidence"] if (results.get(c) or {}).get("status") != "PASS"]
    if bad:
        return False, f"cited check(s) did not PASS in this run: {', '.join(bad)}"
    return True, ""


def review(task, ctx):
    """VIS-01 evaluation from the artifacts produced in this run (ctx['images']) and review files."""
    tdir = ctx["task_dir"]
    images = [p for p in ctx.get("images") or [] if os.path.exists(p)]
    if not images:
        return {"status": "UNCERTAIN", "observed": "no visual artifact produced", "expected": "render/capture image",
                "evidence": [ctx.get("render_error", "")], "fix": "Fix the render error, or attach an image with qa_gate run --image PATH."}
    newest_img = max(os.path.getmtime(p) for p in images)
    vr_path = os.path.join(tdir, "vision-review.json")
    ar_path = os.path.join(tdir, "agent-review.json")
    vr = json.load(open(vr_path)) if os.path.exists(vr_path) else None
    ar = json.load(open(ar_path)) if os.path.exists(ar_path) else None
    # An agent review stays valid when the render it reviewed is byte-identical to this run's render
    # (the gate re-renders on every run, so a timestamp alone can never be newer).
    render_sha = _sha(images[0]) if images else ""
    ar_fresh = bool(ar) and (ar.get("t", 0) >= newest_img or (render_sha and ar.get("render_sha") == render_sha))
    disputes = (ar or {}).get("disputes", {}) if ar_fresh else {}
    results = ctx.get("results") or {}
    so = ctx.get("signoff") or {}
    signed = "VIS-01" in (so.get("items") or [])
    need = {"id": "VIS-01", "why": "no independent vision review of the latest render; the agent's own review does not count", "kind": "signoff"}

    if ar_fresh and ar.get("issues"):  # the agent's own findings can only add problems, never clear them
        return {"status": "FAIL", "observed": f"agent review lists {len(ar['issues'])} issues", "expected": "none",
                "evidence": ar["issues"], "fix": "Fix the listed issues and re-run qa_gate run."}
    if vr and vr.get("ran") and vr.get("t", 0) >= newest_img:
        open_issues, notes = [], []
        for i, x in enumerate(vr.get("issues") or []):
            tag = f"#{i}: [{x.get('severity')}] {x.get('description')} @ {x.get('where','')}"
            if x.get("contradicts_deterministic") and not x.get("elements"):
                notes.append(f"auto-rejected as unsupported (contradicts a deterministic PASS, names no elements): {tag}")
                continue
            if str(i) in disputes:
                ok, why = _dispute_ok(disputes[str(i)], results)
                if ok:
                    notes.append(f"disputed {tag}: {disputes[str(i)].get('reason', '')} [evidence {', '.join(disputes[str(i)]['evidence'])} PASS]")
                    continue
                tag += f"  (dispute rejected: {why})"
            open_issues.append(tag)
        if open_issues and signed:
            return {"status": "PASS", "observed": f"engineer signed off VIS-01 with {len(open_issues)} reviewer issues open",
                    "expected": "", "evidence": [f"signed by {so.get('by')} {so.get('at')}: {so.get('reason', '')}"] + open_issues + images, "fix": "",
                    "signed_off_from": {"status": "FAIL", "observed": f"vision reviewer found {len(open_issues)} open issues"}}
        if open_issues:
            ctx.setdefault("needs_signoff", []).append(need)
            return {"status": "FAIL", "observed": f"vision reviewer found {len(open_issues)} open issues",
                    "expected": "no open visual issues", "evidence": open_issues + notes + [f"images: {images}"],
                    "fix": "Fix each issue and re-run, or dispute a false positive with evidence: "
                           "qa_gate review --dispute N \"reason\" --evidence CAD-LVL-01,..."}
        return {"status": "PASS", "observed": "independent visual review clean", "expected": "",
                "evidence": [f"vision reviewer ({vr.get('model')}): no open issues"] + notes + images, "fix": ""}
    if signed:
        return {"status": "PASS", "observed": "engineer sign-off of the visual check (no independent vision review this run)", "expected": "",
                "evidence": [f"signed by {so.get('by')} {so.get('at')}: {so.get('reason', '')}", f"vision: {(vr or {}).get('note', 'not run')}"] + images,
                "fix": "", "signed_off_from": {"status": "UNCERTAIN", "observed": f"no independent visual review ({(vr or {}).get('note', 'not run')})"}}
    ctx.setdefault("needs_signoff", []).append(need)
    note = (vr or {}).get("note", "vision review not run for the latest images")
    return {"status": "UNCERTAIN", "observed": f"no independent visual review ({note})",
            "expected": "independent vision review of the latest render, or the engineer's sign-off",
            "evidence": ([f"agent review recorded (not independent): {ar.get('checked')}"] if ar_fresh and ar.get("ok") else []) + images,
            "fix": "Restore the vision service (ANTHROPIC_API_KEY / network) and re-run, or have the engineer review the images and run "
                   "`python scripts/qa_gate.py signoff --task ID --items VIS-01` in their own terminal."}
