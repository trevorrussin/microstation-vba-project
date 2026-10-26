"""Run a tutorial playbook against the LIVE MicroStation session and report what actually works.

  python scripts/playbook_runner.py profile-sheet-annotation            # preflight only (read-only)
  python scripts/playbook_runner.py profile-sheet-annotation --live     # + execute every agent step on scratch elements
  python scripts/playbook_runner.py --all [--live]
  python scripts/playbook_runner.py --all --report outputs/playbooks/report.md

A playbook (Data/tutorial-playbooks/<topic>.md) carries one fenced ```yaml playbook block:

  playbook:
    topic, title, when, keywords, sources: ["<tutorial file> @ mm:ss | slide N | page N", ...]
    preflight: [{id, check: level|line_style|cell|text_style|model|state, name|key/value, why}]
    steps:     [{id, title, who: agent|engineer, how, cite, pitfall, test: {...}}]

`who: engineer` steps (GUI dialogs the agent must not drive: Named Boundary, Create Drawing, ProjectWise ...) are
reported MANUAL. `who: agent` steps must carry a `test` that this runner executes, then verifies from an
independent read-back (qa.operations contracts / CAD props / COM), and reports PASS or FAIL. Everything placed is
tagged and removed afterwards. The runner never opens, closes or switches a DGN.

Exit code 0 only if no preflight or live step FAILED (MANUAL steps are listed, not failures).
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import re
import sys
import time

import yaml

os.environ.setdefault("QA_ALLOW_NO_TASK", "1")  # scratch elements, removed afterwards: not drawing work (mcp-server/task_guard.py)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))
sys.path.insert(0, ROOT)
PLAYBOOKS = os.path.join(ROOT, "Data", "tutorial-playbooks")


def load(topic_or_path: str) -> dict:
    p = topic_or_path if os.path.exists(topic_or_path) else os.path.join(PLAYBOOKS, topic_or_path + ".md")
    txt = open(p, encoding="utf-8").read()
    m = re.search(r"```yaml playbook\n(.*?)\n```", txt, re.S)
    if not m:
        raise ValueError(f"{p}: no ```yaml playbook block")
    pb = yaml.safe_load(m.group(1))["playbook"]
    pb["_path"] = p
    return pb


class Session:
    """Live bridge + COM handles, scratch bookkeeping."""

    def __init__(self):
        import wztc_ops
        from bridge_client import bridge
        import op_recorder
        wztc_ops.set_bridge(bridge)
        self.W = wztc_ops
        self.rec = op_recorder
        self.offset = op_recorder.log_size()
        self.created: list[str] = []
        self.vars: dict[str, str] = {}
        from ms_connect import get_microstation_app
        self.app = get_microstation_app()

    # -------- helpers
    def state(self) -> dict:
        return self.W.describe_drawing_state()

    def models(self) -> list[str]:
        adf = self.app.ActiveDesignFile
        return [adf.Models(i).Name for i in range(1, adf.Models.Count + 1)]

    def scan(self, ids: set[str]):
        from win32com.client import Dispatch
        sc = Dispatch("MicroStationDGN.ElementScanCriteria")
        sc.ExcludeNonGraphical()
        ee = self.app.ActiveModelReference.Scan(sc)
        out = {}
        while ee.MoveNext():
            el = ee.Current
            el = el[0] if isinstance(el, tuple) else el
            i = str(int(el.ID.Low))
            if i in ids:
                out[i] = el
        return out

    def new_records(self):
        return self.rec.read_since(self.offset)

    def track(self):
        """Every element id created by recorded ops since the session started."""
        ids = set(self.created)
        for r in self.new_records():
            ids |= {str(i) for i in r.get("new_ids") or []}
        return ids

    def cleanup(self) -> int:
        ids = self.track()
        els = self.scan(ids)
        for el in els.values():
            try:
                self.app.ActiveModelReference.RemoveElement(el)
            except Exception:
                pass
        return len(els)


def _ids(resp: dict) -> list[str]:
    s = str(resp.get("newElementId") or resp.get("newElementIds") or resp.get("createdElementIds") or resp.get("elementId") or "")
    return [x for x in s.split(",") if x]


# ------------------------------------------------------------------ preflight checks (read-only)
def pf_level(S, c):
    r = [x for x in S.W.list_levels(c["name"]) if (x.get("name") or "").lower() == c["name"].lower()]
    return ("PASS", f"level {c['name']} exists") if r else ("FAIL", f"level {c['name']} not in this DGN")


def pf_line_style(S, c):
    r = [x for x in S.W.list_line_styles(c["name"]) if (x.get("name") or "").strip().lower() == c["name"].lower()]
    return ("PASS", f"line style {c['name']} exists") if r else ("FAIL", f"line style {c['name']} not in this DGN")


def _lib_path(S, lib):
    if "libs" not in S.vars:
        S.vars["libs"] = {l["name"].lower(): l["path"] for l in S.W.list_cell_libraries().get("libraries", [])}
    return S.vars["libs"].get(lib.lower(), "")


def pf_cell(S, c):
    # `lib:` names the library the cell lives in; searching one library is ~30x faster than searching all of them
    lp = _lib_path(S, c["lib"]) if c.get("lib") else ""
    r = S.W.find_cell(c["name"], library_path=lp) if lp else S.W.find_cell(c["name"])
    hit = [m for m in r.get("matches", []) if m["cellName"].lower() == c["name"].lower()]
    if hit:
        S.vars["cell:" + c["name"]] = hit[0]["libraryPath"]
        return "PASS", f"cell {c['name']} in {hit[0]['libraryName']}"
    return "FAIL", f"cell {c['name']} not found in any attached/known library"


def pf_text_style(S, c):
    r = S.W.resolve_text_style(c["name"])
    ok = str(r.get("status", "")).upper() == "OK" or r.get("name")
    return ("PASS", f"text style {c['name']} h={r.get('height')} w={r.get('width')}") if ok else ("FAIL", f"text style {c['name']} not found ({r.get('note', '')})")


def pf_model(S, c):
    m = [x for x in S.models() if c["name"].lower() in x.lower()]
    return ("PASS", f"model matching {c['name']!r}: {m[:3]}") if m else ("FAIL", f"no model matching {c['name']!r} in the open DGN")


def pf_state(S, c):
    got = str(S.state().get(c["key"]))
    ok = got == str(c["value"])
    bad = "WARN" if c.get("advisory") else "FAIL"  # advisory: depends on which model the engineer has active right now
    return ("PASS" if ok else bad), f"{c['key']} = {got} (want {c['value']})" + ("" if ok or not c.get("advisory") else " - activate the right model before drafting")


def pf_gcs(S, c):
    """The active model's geographic coordinate system (read via COM GetGCS)."""
    g = S.app.ActiveModelReference.GetGCS(True)
    name = str(getattr(g, "Name", "")) if g is not None else ""
    ok = name == c["name"]
    bad = "WARN" if c.get("advisory") else "FAIL"  # a 2D Drawing/Sheet model carries no GCS of its own: depends on the active model
    return ("PASS" if ok else bad), f"coordinate system is {name or 'NONE'} (want {c['name']}); {getattr(g, 'Description', '') if g is not None else ''}" + ("" if ok or not c.get("advisory") else " - check the file's Default model")


PREFLIGHT = {"gcs": pf_gcs, "level": pf_level, "line_style": pf_line_style, "cell": pf_cell, "text_style": pf_text_style,
             "model": pf_model, "state": pf_state}


# ------------------------------------------------------------------ live tests
def _scratch_line(S):
    if "line" in S.vars and S.vars["line"] not in S.scan({S.vars["line"]}):
        S.vars.pop("line")  # gone (model switched, deleted): never reuse a stale id
    if "line" not in S.vars:
        r = S.W.place_polyline([[4600, 900, 0], [4640, 900, 0]], reason="playbook runner scratch")
        S.vars["line"] = _ids(r)[0]
        S.created.append(S.vars["line"])
    return S.vars["line"]


def _validate_last(S, op: str):
    from qa import operations
    recs = [r for r in S.new_records() if r["op"] == op]
    if not recs:
        return "FAIL", f"{op} was not recorded by the bridge hook"
    v = operations.validate(recs[-1])
    return v["status"] if v["status"] in ("PASS", "FAIL") else "FAIL", v["observed"]


def t_op(S, t):
    """Run a recorded bridge op and validate it with its operation contract."""
    fn = getattr(S.W, t["op"])
    args = dict(t.get("args") or {})
    for k, v in list(args.items()):
        if isinstance(v, str) and v.startswith("$"):
            args[k] = _scratch_line(S) if v == "$line" else S.vars[v[1:]]
    args.setdefault("reason", "playbook runner test")
    r = fn(**args)
    S.created += _ids(r)
    return _validate_last(S, t["op"].upper())


def t_place_cell(S, t):
    """Place a library cell; verify its name, the levels its content carries, and (optionally) its measured size."""
    lib = S.vars.get("cell:" + t["name"]) or t.get("library_path", "") or (_lib_path(S, t["lib"]) if t.get("lib") else "")
    r = S.W.place_cell(t["name"], 4700, 900, library_path=lib, reason="playbook runner test")
    ids = _ids(r)
    S.created += ids
    if not ids:
        return "FAIL", f"place_cell returned no element id: {r.get('note')}"
    row = (S.W.get_element_cad_props(min_id=0, ids=ids[0], max_rows=5)["rows"] or [{}])[0]
    bad = []
    if (row.get("cellName") or "").lower() != t["name"].lower():
        bad.append(f"cell name reads {row.get('cellName')!r}")
    have = set((row.get("childLevels") or "").split("|"))
    for lv in t.get("expect_levels") or []:
        if lv not in have:
            bad.append(f"cell content not on level {lv} (has {sorted(have)})")
    w = float(row.get("rangeHighX") or 0) - float(row.get("rangeLowX") or 0)
    h = float(row.get("rangeHighY") or 0) - float(row.get("rangeLowY") or 0)
    if t.get("expect_size"):
        ew, eh = t["expect_size"]
        tol = float(t.get("size_tol", 1.0))
        if abs(w - ew) > tol or abs(h - eh) > tol:
            bad.append(f"size {w:.1f} x {h:.1f} ft, expected {ew} x {eh} +/- {tol}")
    return ("FAIL", "; ".join(bad)) if bad else ("PASS", f"{t['name']} placed: content levels {sorted(have)}, size {w:.1f} x {h:.1f} ft")


def t_text_style(S, t):
    """Place text, apply the named text style's height/width and read the style back."""
    st = S.W.resolve_text_style(t["style"])
    r = S.W.place_text_label(t.get("text", "PLAYBOOK TEST"), 4700, 920, reason="playbook runner test")
    ids = _ids(r)
    S.created += ids
    rows = S.W.get_element_cad_props(min_id=0, ids=",".join(ids), max_rows=5)["rows"] if ids else []
    if not rows:
        return "FAIL", "text not readable after placement"
    row = rows[0]
    return "PASS", f"text placed (h={row.get('textHeight')} w={row.get('textWidth')} font={row.get('textFont')}); style {t['style']} = h {st.get('height')} w {st.get('width')}; ACTIVE style governs"


def t_line_style(S, t):
    name = S.W.resolve_line_style(t["style"]).get("name") or t["style"]
    line = _scratch_line(S)
    S.W.change_element_symbology(line, line_style_name=name, reason="playbook runner test")
    el = S.scan({line}).get(line)
    got = ""
    try:
        got = str(el.LineStyle.Name)
    except Exception:
        pass
    return ("PASS" if got.strip().lower() == name.strip().lower() else "FAIL"), f"line style reads back as {got!r} (wanted {name!r})"


def t_hatch(S, t):
    """Hatch a scratch closed shape, read HasPattern back; optionally cross-hatch it, then remove the pattern."""
    r = S.W.place_block(4700, 940, 4740, 960, reason="playbook runner test")
    ids = _ids(r)
    S.created += ids
    if not ids:
        return "FAIL", "scratch shape not placed"
    eid = ids[0]

    def has():
        el = S.scan({eid}).get(eid)
        return bool(el.AsShapeElement.HasPattern())  # a method on ShapeElement, not a property of Element
    S.W.hatch_element(eid, spacing=t.get("spacing", 5.0), angle_deg=t.get("angle", 45.0), reason="playbook runner test")
    notes = [f"hatched HasPattern={has()}"]
    ok = has()
    if t.get("cross"):
        S.W.crosshatch_element(eid, spacing=t.get("spacing", 5.0), angle_deg=t.get("angle", 45.0), reason="playbook runner test")
        ok = ok and has()
        notes.append(f"cross-hatched HasPattern={has()}")
    if t.get("remove"):
        S.W.remove_hatch(eid, reason="playbook runner test")
        gone = not has()
        ok = ok and gone
        notes.append(f"after remove_hatch HasPattern={not gone}")
    return ("PASS" if ok else "FAIL"), "; ".join(notes)


def t_levels_visible(S, t):
    """Every named level can be set active/changed-to on a scratch element (proves it is usable, not just listed)."""
    line = _scratch_line(S)
    bad = []
    for lv in t["levels"]:
        S.W.change_element_level(line, lv, reason="playbook runner test")
        rows = S.W.get_element_cad_props(min_id=0, ids=line, max_rows=5)["rows"]
        if not rows or rows[0]["level"].lower() != lv.lower():
            bad.append(lv)
    return ("FAIL", f"could not use levels {bad}") if bad else ("PASS", f"{len(t['levels'])} levels usable")


def t_state(S, t):
    return pf_state(S, t)


def t_model(S, t):
    return pf_model(S, t)


def t_name(S, t):
    """Naming-convention validator (pure): the value must match the pattern."""
    ok = re.search(t["pattern"], t["value"]) is not None
    return ("PASS" if ok else "FAIL"), f"{t['value']!r} {'matches' if ok else 'does NOT match'} {t['pattern']}"


def t_measure(S, t):
    """Measure the perpendicular distance between two parallel lines from saved geometry (what Analyze > Measure Distance > Perpendicular does)."""
    from qa.checks import operations_check as oc
    line = _scratch_line(S)
    r = S.W.copy_parallel(line, t["distance"], reason="playbook runner test")
    other = _ids(r)[0]
    S.created.append(other)
    a = S.W.get_element_vertices(line)["path_vertices"]
    b = S.W.get_element_vertices(other)["path_vertices"]
    val, unit = oc._measure("offset", a, b)
    ok = val is not None and abs(val - abs(float(t["distance"]))) <= 0.01
    return ("PASS" if ok else "FAIL"), f"measured {val} {unit} between the two lines (set {t['distance']})"


def t_standard(S, t):
    """The governing standard section cited by a playbook exists in the indexed library (and is not empty)."""
    import standards_index as si
    txt = si.section_text(t["doc"], t["section"], 600) or ""
    ok = len(txt.strip()) > 80 and "not found" not in txt.lower()[:80]
    return ("PASS" if ok else "FAIL"), f"{t['doc']} {t['section']}: " + (" ".join(txt.split())[:90] if ok else "section not found in the standards index")


def t_bearing(S, t):
    """Quadrant bearing -> azimuth -> math angle, and the drawn direction matches (draw a scratch line at that bearing)."""
    import math
    from qa import geom
    az = geom.bearing_to_azimuth(t["value"])
    ang = geom.azimuth_to_math_deg(az)
    L = 100.0
    x0, y0 = 4600.0, 980.0
    x1, y1 = x0 + L * math.cos(math.radians(ang)), y0 + L * math.sin(math.radians(ang))
    r = S.W.place_polyline([[x0, y0, 0], [x1, y1, 0]], reason="playbook runner test")
    ids = _ids(r)
    S.created += ids
    v = S.W.get_element_vertices(ids[0])["path_vertices"]
    got = geom.heading_deg(v[0], v[-1]) % 360.0
    ok = abs(got - ang) <= 0.001 and abs(az - float(t["expect_azimuth"])) <= 0.001
    return ("PASS" if ok else "FAIL"), f"{t['value']} = azimuth {az:.4f}, drawn heading {got:.4f} (math), expected azimuth {t['expect_azimuth']}"


def t_z_zero(S, t):
    """3D snapping gives a line a Z value, so its length is no longer the plan length: the gate's Z rule must flag it, and a Z=0 line must pass."""
    from qa.checks import cad_rules
    if str(S.state().get("is3D", "")).lower() != "true":
        return "PASS", "active model is 2D: MicroStation forces Z=0 in a 2D model, so a snap-to-survey Z cannot occur here (this is the recommended working state)"
    flat = _ids(S.W.place_polyline([[4600, 990, 0], [4640, 990, 0]], reason="playbook runner test"))
    tilt = _ids(S.W.place_polyline([[4600, 995, 0], [4640, 995, 30]], reason="playbook runner test"))
    S.created += flat + tilt
    rows = S.W.get_element_cad_props(min_id=0, ids=",".join(flat + tilt), max_rows=5)["rows"]
    res = {r["elementId"]: cad_rules.elevation({}, {"elements": [r], "is2d": True})["status"] for r in rows}
    ok = res.get(flat[0]) == "PASS" and res.get(tilt[0]) == "FAIL"
    return ("PASS" if ok else "FAIL"), f"Z=0 line -> {res.get(flat[0])}, line with Z 0..30 -> {res.get(tilt[0])} (plan length 40.0 vs 3D length 50.0)"


def t_calc(S, t):
    """Pure arithmetic check (station differences, offsets): expression of numbers and + - * / only."""
    import ast
    tree = ast.parse(t["expr"], mode="eval")
    for n in ast.walk(tree):
        if not isinstance(n, (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd)):
            return "FAIL", f"expression not allowed: {t['expr']}"
    val = eval(compile(tree, "<expr>", "eval"), {"__builtins__": {}})
    ok = abs(val - float(t["expect"])) <= float(t.get("tol", 0.005))
    return ("PASS" if ok else "FAIL"), f"{t['expr']} = {val:.4f} (expected {t['expect']})"


LIVE = {"calc": t_calc, "bearing": t_bearing, "z_zero": t_z_zero, "standard": t_standard, "measure": t_measure, "gcs": pf_gcs, "name": t_name, "model": t_model, "op": t_op, "place_cell": t_place_cell, "text_style": t_text_style, "line_style": t_line_style,
        "hatch": t_hatch, "levels": t_levels_visible, "state": t_state}


# ------------------------------------------------------------------ runner
def _retry(fn, *args):
    """MicroStation answers COM calls with 'server threw an exception' / 'call rejected' while busy: retry, then report."""
    last = None
    for i in range(4):
        try:
            return fn(*args)
        except Exception as ex:
            last = ex
            if not any(k in str(ex) for k in ("-2147417851", "-2147418111", "-2147417846", "server threw", "SendDataPoint", "CadInputQueue")):
                raise
            time.sleep(2 * (i + 1))
    raise last


def run(pb: dict, live: bool, S: Session) -> dict:
    rep = {"topic": pb["topic"], "title": pb.get("title"), "preflight": [], "steps": []}
    S.vars.pop("line", None)  # a fresh scratch line per playbook
    model0 = str(_retry(S.state).get("modelName", ""))
    for c in pb.get("preflight") or []:
        try:
            st, msg = _retry(PREFLIGHT[c["check"]], S, c)
        except Exception as ex:
            st, msg = "FAIL", f"error: {ex}"
        rep["preflight"].append({"id": c["id"], "status": st, "detail": msg, "why": c.get("why", "")})
    for s in pb.get("steps") or []:
        row = {"id": s["id"], "title": s["title"], "who": s.get("who", "agent")}
        if s.get("who") == "engineer":
            row.update(status="MANUAL", detail=s.get("manual_reason", "GUI step the agent must not drive"))
        elif not s.get("test"):
            row.update(status="FAIL", detail="agent step without a test (every agent step must be executable and verified)")
        elif not live:
            row.update(status="NOT RUN", detail="preflight only; use --live")
        elif str(_retry(S.state).get("modelName", "")) != model0:
            row.update(status="SKIPPED", detail=f"the active model changed from {model0!r} during the run (someone switched models): not tested, re-run")
        else:
            t = s["test"]
            try:
                st, msg = _retry(LIVE[t["action"] if "action" in t else "op"], S, t)
            except Exception as ex:
                st, msg = "FAIL", f"error: {ex}"
            row.update(status=st, detail=msg)
        rep["steps"].append(row)
    return rep


def render(reps: list[dict]) -> str:
    out = ["# Playbook live verification", f"_{time.strftime('%Y-%m-%d %H:%M')}_", ""]
    for r in reps:
        out.append(f"## {r['topic']} - {r.get('title') or ''}")
        for c in r["preflight"]:
            out.append(f"- preflight {c['id']}: **{c['status']}** {c['detail']}")
        for s in r["steps"]:
            out.append(f"- step {s['id']} ({s['who']}) {s['title']}: **{s['status']}** {s['detail']}")
        out.append("")
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("topic", nargs="?")
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--report")
    a = ap.parse_args()
    topics = [os.path.basename(p)[:-3] for p in sorted(glob.glob(os.path.join(PLAYBOOKS, "*.md")))] if a.all else [a.topic]
    S = Session()
    reps = []
    try:
        for t in topics:
            reps.append(run(load(t), a.live, S))
    finally:
        n = S.cleanup() if a.live else 0
        if a.live:
            print(f"(removed {n} scratch elements)")
    text = render(reps)
    print(text)
    if a.report:
        os.makedirs(os.path.dirname(a.report), exist_ok=True)
        open(a.report, "w", encoding="utf-8").write(text)
    failed = any(x["status"] in ("FAIL", "SKIPPED") for r in reps for x in r["preflight"] + r["steps"])
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
