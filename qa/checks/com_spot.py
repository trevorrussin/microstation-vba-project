"""COM-01: an independent evidence path. Every other check reads MicroStation through the WZTC bridge (VBA); a bridge
fault (e.g. the shared result-slot defect found 2026-10-23) would corrupt the change set, the operation record and
the checks together while they all still agree with each other. This check re-reads the same elements through
MicroStation's own COM object model (a different code path) and requires range box and level to match what the
bridge reported.

The elements the task changed are always compared; for the rest of the model a seeded sample (all of it when small).
"""
from __future__ import annotations

import os
import random
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "mcp-server"))

TOL = 0.01
SAMPLE = 300
MAX_SCAN = 30000
MAX_IDS = 25


def _res(status, observed="", expected="", evidence=None, fix=""):
    return {"status": status, "observed": observed, "expected": expected, "evidence": (evidence or [])[:MAX_IDS], "fix": fix}


def read_com(ids: set) -> dict:
    """{element id: {low, high, level}} for the given ids, read through COM (active model)."""
    from win32com.client import Dispatch
    import ms_connect
    app = ms_connect.get_microstation_app()
    sc = Dispatch("MicroStationDGN.ElementScanCriteria")
    sc.ExcludeNonGraphical()
    ee = app.ActiveModelReference.Scan(sc)
    out, n = {}, 0
    while ee.MoveNext() and n < MAX_SCAN:
        n += 1
        el = ee.Current
        el = el[0] if isinstance(el, tuple) else el
        i = str(int(el.ID.Low))
        if i in ids:
            rg = el.Range
            try:
                lvl = el.Level.Name
            except Exception:
                lvl = None
            out[i] = {"low": (rg.Low.X, rg.Low.Y), "high": (rg.High.X, rg.High.Y), "level": lvl}
            if len(out) == len(ids):
                break
    return out


def pick(rows, changed_ids, seed):
    by_id = {str(r["elementId"]): r for r in rows}
    want = [i for i in changed_ids if i in by_id]
    rest = sorted(set(by_id) - set(want))
    if len(rest) > SAMPLE:
        rest = random.Random(seed).sample(rest, SAMPLE)
    return want + rest, by_id


def verify(task, ctx):
    rows = ctx.get("state_rows")
    if not rows:
        return _res("NA", "no model read-back to cross-check")
    changed = [str(e["elementId"]) for e in ctx.get("elements") or []]
    ids, by_id = pick(rows, changed, ctx.get("state_fp") or "seed")
    try:
        com = read_com(set(ids))
    except Exception as ex:
        return _res("UNCERTAIN", f"MicroStation COM read failed: {ex}", "bridge read-back confirmed through COM",
                    fix="Open MicroStation and re-run; without a second evidence path the bridge read-back is unconfirmed.")
    missing = [i for i in ids if i not in com]
    bad = []
    for i in ids:
        c, r = com.get(i), by_id[i]
        if c is None:
            continue
        try:
            d = max(abs(c["low"][0] - float(r["rangeLowX"])), abs(c["low"][1] - float(r["rangeLowY"])),
                    abs(c["high"][0] - float(r["rangeHighX"])), abs(c["high"][1] - float(r["rangeHighY"])))
        except (KeyError, TypeError, ValueError):
            d = 0.0
        if d > TOL:
            bad.append(f"{i}: range differs by {d:.3f} ft (bridge vs COM)")
        elif c["level"] is not None and (r.get("level") or "") and c["level"] != r["level"]:
            bad.append(f"{i}: level {r['level']!r} (bridge) vs {c['level']!r} (COM)")
    if bad or missing:
        return _res("UNCERTAIN", f"bridge read-back disagrees with MicroStation COM for {len(bad)} elements ({len(missing)} not found via COM)",
                    "bridge and COM agree on range box and level", bad + [f"{i}: not found via COM" for i in missing],
                    "The evidence the other checks used may be wrong. Re-run once; if it persists, treat the bridge result as untrusted "
                    "and tell the engineer.")
    return _res("PASS", f"{len(ids)} of {len(rows)} elements cross-checked against MicroStation COM (range + level): bridge evidence confirmed")
