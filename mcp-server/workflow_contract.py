"""Workflow contracts: preflight, mutation guard and verification for engineering workflows (first one: profile-design).

A contract (Data/workflow-contracts/<id>.yaml) says which inputs only the engineer can supply, which ops may run, which are
refused, and what counts as complete. This module is the enforcement:

  preflight(workflow_id, inputs, reader=None)  -> READY | WAITING_FOR_INPUT | BLOCKED | UNVERIFIED  (+ missing, problems, next steps)
  guard_cif_exec(cmd, args)                    -> raises WorkflowRefused when an op would break the contract
                                                  (called from wztc_ops.cif_exec, the shared mutation boundary)
  expected_y / verify_profile                  -> independent re-computation of the profile from the PVI table, compared with a read-back

The guard does not depend on the agent remembering anything: wztc_ops.cif_exec calls it for every CIF edit.
"""
from __future__ import annotations

import hashlib
import json
import math
import os
import time
from pathlib import Path
from typing import Callable

ROOT = Path(__file__).resolve().parents[1]
CONTRACTS = ROOT / "Data" / "workflow-contracts"
INPUTS = ROOT / "Data" / "workflow-inputs"
STATE = ROOT / "outputs" / "workflow"
READY_TTL_S = 12 * 3600
FT_TO_M = 1200.0 / 3937.0          # US survey foot; CIF coordinates, stations and elevations are metres
TEST_PREFIX = "WZTC_TEST_"          # throwaway objects the agent created itself

READY, WAITING, BLOCKED, UNVERIFIED = "READY", "WAITING_FOR_INPUT", "BLOCKED", "UNVERIFIED"


class WorkflowRefused(RuntimeError):
    """Raised instead of running an op that violates a workflow contract."""


def load_contract(workflow_id: str) -> dict:
    import yaml
    p = CONTRACTS / f"{workflow_id}.yaml"
    if not p.exists():
        raise FileNotFoundError(f"no workflow contract '{workflow_id}' in {CONTRACTS}")
    return yaml.safe_load(p.read_text(encoding="utf-8"))


def load_inputs(path: str | Path) -> dict:
    import yaml
    return yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}


# ---------------------------------------------------------------- normalising
def _num(v):
    if isinstance(v, bool) or v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def normalised_pvis(inputs: dict) -> list[dict]:
    """PVI rows in METRES (CIF units). Returns [] when units are missing."""
    u = inputs.get("units")
    k = {"m": 1.0, "ft": FT_TO_M}.get(u)
    if k is None:
        return []
    rows = []
    for r in inputs.get("pvis") or []:
        s, e, c = _num((r or {}).get("station")), _num((r or {}).get("elevation")), _num((r or {}).get("curve_length"))
        rows.append({"station": None if s is None else s * k, "elevation": None if e is None else e * k,
                     "curve_length": None if c is None else c * k})
    return rows


def inputs_hash(inputs: dict) -> str:
    """Hash of exactly what the guarded op will use (alignment + units + PVI table), so the READY token can't be reused for other numbers."""
    key = {"alignment": inputs.get("alignment"), "units": inputs.get("units"),
           "pvis": [[(r or {}).get("station"), (r or {}).get("elevation"), (r or {}).get("curve_length")] for r in (inputs.get("pvis") or [])]}
    return hashlib.sha256(json.dumps(key, sort_keys=True).encode()).hexdigest()[:16]


# ---------------------------------------------------------------- profile maths (independent of OpenRoads)
def expected_y(pvis_m: list[dict], x: float) -> float | None:
    """Elevation of a PVI profile at station x: straight grades with symmetric parabolic vertical curves of length L at each PVI."""
    n = len(pvis_m)
    if n < 2 or x < pvis_m[0]["station"] - 1e-9 or x > pvis_m[-1]["station"] + 1e-9:
        return None

    def grade(i):  # grade between PVI i and i+1
        return (pvis_m[i + 1]["elevation"] - pvis_m[i]["elevation"]) / (pvis_m[i + 1]["station"] - pvis_m[i]["station"])

    for i in range(1, n - 1):                     # inside the curve at PVI i?
        L = pvis_m[i]["curve_length"] or 0.0
        if L > 0:
            bvc = pvis_m[i]["station"] - L / 2
            if bvc <= x <= pvis_m[i]["station"] + L / 2:
                g1, g2 = grade(i - 1), grade(i)
                y_bvc = pvis_m[i]["elevation"] - g1 * L / 2
                return y_bvc + g1 * (x - bvc) + (g2 - g1) / (2 * L) * (x - bvc) ** 2
    for i in range(n - 1):                        # on a tangent
        if pvis_m[i]["station"] - 1e-9 <= x <= pvis_m[i + 1]["station"] + 1e-9:
            return pvis_m[i]["elevation"] + grade(i) * (x - pvis_m[i]["station"])
    return None


# ---------------------------------------------------------------- preflight
def preflight(workflow_id: str, inputs: dict, reader: Callable | None = None) -> dict:
    """reader(cmd, args) -> list of row dicts (cif_read). None = offline: live checks are reported as unverified."""
    c = load_contract(workflow_id)
    missing, problems, unverified = [], [], []
    prohibited = list(c.get("prohibited_assumptions", []))

    align = inputs.get("alignment")
    if not align or not str(align).strip():
        missing.append("alignment")
    units = inputs.get("units")
    if units is None:
        missing.append("units (ft or m)")
    elif units not in ("ft", "m"):
        problems.append(f"units must be 'ft' or 'm', got {units!r}")
    raw = inputs.get("pvis") or []
    if not raw:
        missing.append("pvis (station, elevation, curve_length for each PVI)")
    else:
        for i, r in enumerate(raw, 1):
            for f in ("station", "elevation", "curve_length"):
                if _num((r or {}).get(f)) is None:
                    missing.append(f"pvis[{i}].{f}")
    if inputs.get("synthetic_test") and not str(align or "").startswith(TEST_PREFIX):
        problems.append(f"synthetic_test is only allowed on a throwaway alignment whose name starts with {TEST_PREFIX}")

    pv = normalised_pvis(inputs)
    if not missing and pv and not problems:
        if len(pv) < 2:
            problems.append("at least 2 PVIs are required")
        st = [r["station"] for r in pv]
        if any(b <= a for a, b in zip(st, st[1:])):
            problems.append("PVI stations must be strictly increasing")
        for i, r in enumerate(pv):
            if r["curve_length"] < 0:
                problems.append(f"pvis[{i + 1}].curve_length is negative")
        if pv[0]["curve_length"] != 0 or pv[-1]["curve_length"] != 0:
            problems.append("curve_length must be 0 on the first and last PVI")
        if not problems:
            for i in range(1, len(pv) - 1):
                half = pv[i]["curve_length"] / 2
                if half > 0 and (half >= pv[i]["station"] - pv[i - 1]["station"] or half >= pv[i + 1]["station"] - pv[i]["station"]):
                    problems.append(f"pvis[{i + 1}] curve ({pv[i]['curve_length']:.2f} m) does not fit between its neighbours")
            for i in range(1, len(pv) - 2):    # adjacent curves must not overlap
                if pv[i]["station"] + pv[i]["curve_length"] / 2 > pv[i + 1]["station"] - pv[i + 1]["curve_length"] / 2 + 1e-9:
                    problems.append(f"curves at pvis[{i + 1}] and pvis[{i + 2}] overlap")

    live = {}
    if not missing and not problems:
        if reader is None:
            unverified.append("alignment_exists / stations_within_alignment (OpenRoads not queried)")
        else:
            try:
                rows = reader("get_alignment", {"name": align})
                live = {r["key"]: r["value"] for r in rows}
                s0, ln = _num(live.get("startStation")), _num(live.get("length"))
                if s0 is None or ln is None:
                    unverified.append("alignment stationing could not be read")
                else:
                    live["endStation"] = s0 + ln
                    if s0 < 1000.0 * FT_TO_M - 0.01:      # 10+00.00 in survey feet (alignments playbook S5; BRP horizontal @03:58)
                        problems.append(f"alignment starts at station {s0:.2f} m, below 10+00.00; never start an alignment at 0+00 or in the negatives "
                                        "(Modify > Start Station, or a Station Equation from the first station)")
                    if pv[0]["station"] < s0 - 0.01 or pv[-1]["station"] > s0 + ln + 0.01:
                        problems.append(f"PVI stations {pv[0]['station']:.2f}..{pv[-1]['station']:.2f} m lie outside the alignment "
                                        f"({s0:.2f}..{s0 + ln:.2f} m); check the units and the station values")
            except Exception as ex:  # alignment missing or OpenRoads unreachable: say which
                msg = str(ex)
                if "alignment not found" in msg:
                    problems.append(f"alignment '{align}' is not in the open model")
                else:
                    unverified.append(f"OpenRoads could not be read: {msg[:120]}")

    if missing:
        status = WAITING
    elif problems:
        status = BLOCKED
    elif unverified:
        status = UNVERIFIED
    else:
        status = READY
    safe = []
    if status == WAITING:
        safe = ["supply the missing inputs in Data/workflow-inputs/profile-design.yaml (template: profile-design.template.yaml)",
                "read-only meanwhile: list_alignments / get_alignment to confirm the alignment name and its stationing"]
    elif status == BLOCKED:
        safe = ["fix the listed problems in the inputs file and run preflight again"]
    elif status == UNVERIFIED:
        safe = ["run preflight again with OpenRoads connected"]
    else:
        safe = ["run the agent steps (create_profile_from_pvis, verify_profile_pvis), then the manual steps N1, N4-N7"]
    out = {"status": status, "workflow": workflow_id, "inputs_hash": inputs_hash(inputs), "alignment": align,
           "missing_inputs": missing, "problems": problems, "unverified": unverified,
           "prohibited_actions": prohibited, "safe_next_steps": safe,
           "agent_steps": [s["id"] + " " + s["op"] for s in c.get("agent_steps", [])],
           "manual_steps": [f"{s['id']} ({s['who']}): {s['what']}" for s in c.get("manual_steps", [])]}
    if live:
        out["alignment_stationing_m"] = [live.get("startStation"), live.get("endStation")]
    if status == READY:
        STATE.mkdir(parents=True, exist_ok=True)
        (STATE / f"{workflow_id}.ready.json").write_text(
            json.dumps({"workflow": workflow_id, "inputs_hash": out["inputs_hash"], "alignment": align, "at": time.time(),
                        "synthetic_test": bool(inputs.get("synthetic_test"))}), encoding="utf-8")
    return out


def _ready_token(workflow_id: str) -> dict | None:
    p = STATE / f"{workflow_id}.ready.json"
    if not p.exists():
        return None
    try:
        t = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None
    return t if time.time() - t.get("at", 0) <= READY_TTL_S else None


# ---------------------------------------------------------------- mutation guard
def args_hash(args: dict) -> str:
    """Hash of a create_profile_from_pvis call's numbers, comparable with inputs_hash() of the preflighted inputs (metres, so it re-derives)."""
    return hashlib.sha256(json.dumps({"alignment": args.get("name"), "pvis": str(args.get("pvis"))}, sort_keys=True).encode()).hexdigest()[:16]


def pvis_arg(inputs: dict) -> str:
    """The exact `pvis` string handed to the CIF op (metres)."""
    return ";".join(f"{r['station']!r},{r['elevation']!r},{r['curve_length']!r}" for r in normalised_pvis(inputs))


def guard_cif_exec(cmd: str, args: dict, workflow_id: str = "profile-design") -> None:
    """Called by wztc_ops.cif_exec before anything is sent to OpenRoads."""
    c = load_contract(workflow_id)
    name = str((args or {}).get("name", ""))
    if cmd in (c.get("prohibited_ops") or {}):
        if name.startswith(TEST_PREFIX):
            return                                   # throwaway test object the agent itself created
        raise WorkflowRefused(
            f"refused cif_exec {cmd}: {c['prohibited_ops'][cmd]}. Workflow '{workflow_id}' builds a design profile only from the engineer's PVI table: "
            f"fill Data/workflow-inputs/profile-design.yaml, run `python scripts/workflow_preflight.py {workflow_id}`, then use profile_build / create_profile_from_pvis.")
    g = (c.get("guarded_ops") or {}).get(cmd)
    if not g:
        return
    if g.get("enabled") is False and not name.startswith(TEST_PREFIX):
        raise WorkflowRefused(f"refused cif_exec {cmd}: disabled by contract '{workflow_id}' ({g.get('disabled_reason', 'experimental')}).")
    tok = _ready_token(workflow_id)
    if not tok:
        raise WorkflowRefused(f"refused cif_exec {cmd}: preflight for '{workflow_id}' is not READY (no fresh token). Inputs are missing or unchecked; "
                              f"run `python scripts/workflow_preflight.py {workflow_id}`.")
    if g.get("inputs_hash_must_match") and tok.get("pvi_args_hash") and tok["pvi_args_hash"] != args_hash(args):
        raise WorkflowRefused(f"refused cif_exec {cmd}: the alignment/PVI numbers differ from the ones that passed preflight; run preflight again.")
    if g.get("inputs_hash_must_match") and not tok.get("pvi_args_hash"):
        raise WorkflowRefused(f"refused cif_exec {cmd}: the READY token carries no numbers to compare; run preflight through profile_preflight / scripts/workflow_preflight.py.")
    if tok.get("synthetic_test") and not name.startswith(TEST_PREFIX):
        raise WorkflowRefused(f"refused cif_exec {cmd}: the READY token is a synthetic test and may only target {TEST_PREFIX}* objects.")


def preflight_and_arm(workflow_id: str, inputs: dict, reader: Callable | None = None) -> dict:
    """preflight(), and when READY also bind the token to the exact `pvis` argument string the op will receive."""
    out = preflight(workflow_id, inputs, reader)
    if out["status"] == READY:
        p = STATE / f"{workflow_id}.ready.json"
        t = json.loads(p.read_text(encoding="utf-8"))
        t["pvi_args_hash"] = args_hash({"name": inputs.get("alignment"), "pvis": pvis_arg(inputs)})
        p.write_text(json.dumps(t), encoding="utf-8")
        out["pvis_arg_m"] = pvis_arg(inputs)
    return out


# ---------------------------------------------------------------- verification
# Default tolerance 0.003048 m = 0.01 ft: HDM Ch.21 Table 21-4 (p.19), PVI elevations 0.01 ft.
def verify_profile(inputs: dict, reader: Callable, profile_idx: int, samples_per_span: int = 8, tol_m: float = 0.003048) -> dict:
    """Re-read the profile from OpenRoads and compare with the PVI table. Returns {status: PASS|FAIL, max_dev_m, n, worst}."""
    pv = normalised_pvis(inputs)
    xs = []
    for a, b in zip(pv, pv[1:]):
        xs += [a["station"] + (b["station"] - a["station"]) * j / samples_per_span for j in range(samples_per_span)]
    xs.append(pv[-1]["station"])
    for r in pv[1:-1]:                              # and across each curve
        L = r["curve_length"]
        if L > 0:
            xs += [r["station"] - L / 2, r["station"] - L / 4, r["station"], r["station"] + L / 4, r["station"] + L / 2]
    xs = sorted(set(round(x, 6) for x in xs))
    rows = []
    for x in xs:       # one read per station: sample_profile with x0 == x1
        got = reader("sample_profile", {"name": inputs["alignment"], "profileIdx": profile_idx, "x0": x, "x1": x, "step": 1})
        rows.append((x, got[0]))
    worst, n = (0.0, None), 0
    for x, row in rows:
        want = expected_y(pv, x)
        if row.get("defined") != "Y" or want is None:
            return {"status": "FAIL", "reason": f"profile undefined at station {x:.3f}", "n": n}
        d = abs(float(row["y"]) - want)
        n += 1
        if d > worst[0]:
            worst = (d, x)
    return {"status": "PASS" if worst[0] <= tol_m else "FAIL", "max_dev_m": worst[0], "at_station": worst[1], "n": n, "tol_m": tol_m}


SLOPE_TOL = 1e-4        # HDM Table 21-4: rate of grade 0.01 %
CONTINUITY_EPS = 1e-6   # m either side of a PVC / PVI / PVT


def expected_slope(pvis_m: list[dict], x: float) -> float | None:
    """Grade (m/m) of the PVI profile at station x: tangent grade, or the parabola's grade inside a curve."""
    n = len(pvis_m)
    if n < 2 or x < pvis_m[0]["station"] - 1e-9 or x > pvis_m[-1]["station"] + 1e-9:
        return None

    def grade(i):
        return (pvis_m[i + 1]["elevation"] - pvis_m[i]["elevation"]) / (pvis_m[i + 1]["station"] - pvis_m[i]["station"])
    for i in range(1, n - 1):
        L = pvis_m[i]["curve_length"] or 0.0
        if L > 0 and pvis_m[i]["station"] - L / 2 <= x <= pvis_m[i]["station"] + L / 2:
            return grade(i - 1) + (grade(i) - grade(i - 1)) * (x - (pvis_m[i]["station"] - L / 2)) / L
    for i in range(n - 1):
        if pvis_m[i]["station"] - 1e-9 <= x <= pvis_m[i + 1]["station"] + 1e-9:
            return grade(i)
    return None


def _dense_stations(pv: list[dict], step: float = 2.0) -> tuple[list[float], list[float]]:
    """(all stations to sample, break stations where tangent meets curve or PVI)."""
    breaks = [pv[0]["station"], pv[-1]["station"]]
    for r in pv[1:-1]:
        L = r["curve_length"] or 0.0
        breaks += [r["station"] - L / 2, r["station"], r["station"] + L / 2] if L > 0 else [r["station"]]
    xs = set(round(b, 6) for b in breaks)
    for a, b in zip(pv, pv[1:]):
        k = max(1, int((b["station"] - a["station"]) / step))
        xs.update(round(a["station"] + (b["station"] - a["station"]) * j / k, 6) for j in range(k + 1))
    lo, hi = pv[0]["station"], pv[-1]["station"]
    for b in breaks:                                    # one point either side of every break, for continuity
        for e in (-CONTINUITY_EPS, CONTINUITY_EPS):
            if lo - 1e-12 <= b + e <= hi + 1e-12:
                xs.add(round(b + e, 9))
    return sorted(xs), sorted(set(round(b, 6) for b in breaks))


def profile_ranges(inputs: dict, reader: Callable) -> list[list[float]]:
    """[x0, x1] of every profile on the alignment, in order (the baseline for 'originals unchanged')."""
    out = []
    for r in reader("list_profiles", {}):
        if r.get("alignment") == inputs.get("alignment"):
            try:
                out.append([float(r["x0"]), float(r["x1"])])
            except (KeyError, ValueError):
                out.append([float("nan"), float("nan")])
    return out


def save_baseline(workflow_id: str, inputs: dict, reader: Callable) -> Path:
    STATE.mkdir(parents=True, exist_ok=True)
    p = STATE / f"{workflow_id}.before.json"
    p.write_text(json.dumps({"alignment": inputs.get("alignment"), "at": time.time(), "ranges": profile_ranges(inputs, reader)}), encoding="utf-8")
    return p


def load_baseline(workflow_id: str, inputs: dict) -> list | None:
    p = STATE / f"{workflow_id}.before.json"
    if not p.exists():
        return None
    try:
        b = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None
    return b["ranges"] if b.get("alignment") == inputs.get("alignment") else None


def verify_profile_full(inputs: dict, reader: Callable, profile_idx: int, tol_m: float = 0.003048, baseline: list | None = None) -> dict:
    """verify_profile (sampled elevations) plus: dense elevation + grade comparison, exact endpoints, continuity at every PVC/PVI/PVT,
    exactly one profile with the PVI range, the profile is the alignment's ACTIVE profile, original profiles unchanged (when a baseline exists)."""
    pv = normalised_pvis(inputs)
    checks = {}
    base = verify_profile(inputs, reader, profile_idx, tol_m=tol_m)
    checks["elevation_samples"] = {"status": base["status"], "max_dev_m": base.get("max_dev_m"), "n": base.get("n"), "tol_m": tol_m}
    xs, breaks = _dense_stations(pv)
    rows = reader("sample_profile", {"name": inputs["alignment"], "profileIdx": profile_idx, "x": xs})
    byx = {}
    bad_def = 0
    for x, row in zip(xs, rows):
        if row.get("defined") != "Y":
            bad_def += 1
            continue
        byx[x] = (float(row["y"]), float(row["slope"]) if row.get("slope") not in (None, "", "?") else None)
    dev = max((abs(y - expected_y(pv, x)) for x, (y, _) in byx.items()), default=float("inf"))
    checks["dense_elevation"] = {"status": "PASS" if bad_def == 0 and dev <= tol_m else "FAIL", "max_dev_m": dev, "n": len(byx), "undefined_stations": bad_def}
    sl = [abs(s - expected_slope(pv, x)) for x, (_, s) in byx.items() if s is not None and expected_slope(pv, x) is not None]
    checks["grade"] = {"status": "PASS" if sl and max(sl) <= SLOPE_TOL else "FAIL", "max_dev": max(sl) if sl else None, "tol": SLOPE_TOL, "n": len(sl)}
    e0, e1 = byx.get(round(pv[0]["station"], 6)), byx.get(round(pv[-1]["station"], 6))
    ok_ends = bool(e0 and e1 and abs(e0[0] - pv[0]["elevation"]) <= tol_m and abs(e1[0] - pv[-1]["elevation"]) <= tol_m)
    checks["endpoints"] = {"status": "PASS" if ok_ends else "FAIL", "start": e0[0] if e0 else None, "end": e1[0] if e1 else None}
    jumps = []
    for b in breaks:
        lo, hi = byx.get(round(b - CONTINUITY_EPS, 9)), byx.get(round(b + CONTINUITY_EPS, 9))
        if lo and hi and abs(hi[0] - lo[0]) > tol_m:
            jumps.append(b)
    checks["continuity"] = {"status": "PASS" if not jumps else "FAIL", "breaks_checked": len(breaks), "jumps_at": jumps}
    mine = [r for r in profile_ranges(inputs, reader) if abs(r[0] - pv[0]["station"]) < 0.01 and abs(r[1] - pv[-1]["station"]) < 0.01]
    checks["single_profile"] = {"status": "PASS" if len(mine) == 1 else "FAIL", "profiles_with_this_range": len(mine)}
    try:
        act = {r["key"]: r["value"] for r in reader("active_profile", {"name": inputs["alignment"]})}
        a0, _, a1 = (act.get("activeRange") or "").partition("..")
        is_act = act.get("activeIsNull") == "N" and abs(float(a0) - pv[0]["station"]) < 0.01 and abs(float(a1) - pv[-1]["station"]) < 0.01
        checks["active_profile"] = {"status": "PASS" if is_act else "FAIL", "active_range": act.get("activeRange", ""), "active_is_null": act.get("activeIsNull")}
    except Exception as ex:
        checks["active_profile"] = {"status": "UNVERIFIED", "why": str(ex)[:120]}
    if baseline is None:
        checks["originals_unchanged"] = {"status": "UNVERIFIED", "why": "no baseline snapshot (profile_build saves one before the profile exists)"}
    else:
        now = profile_ranges(inputs, reader)
        left = list(now)
        missing = []
        for r in baseline:
            hit = next((i for i, q in enumerate(left) if abs(q[0] - r[0]) < 1e-6 and abs(q[1] - r[1]) < 1e-6), None)
            if hit is None:
                missing.append(r)
            else:
                left.pop(hit)
        ours = [q for q in left if abs(q[0] - pv[0]["station"]) < 0.01 and abs(q[1] - pv[-1]["station"]) < 0.01]
        rest = [q for q in left if q not in ours]
        stub_known = any(abs(r[0]) < 1e-6 and abs(r[1] - 100.0) < 1e-6 for r in baseline)
        if not stub_known and rest and abs(rest[0][0]) < 1e-6 and abs(rest[0][1] - 100.0) < 1e-6:
            rest = rest[1:]          # OpenRoads adds its 0..100 stub profile when an alignment gets its first profile
        checks["originals_unchanged"] = {"status": "PASS" if not missing and len(ours) == 1 and not rest else "FAIL",
                                         "missing_or_changed": missing, "new_profiles": len(ours), "unexpected_extra_profiles": rest}
    st = [c["status"] for c in checks.values()]
    status = "FAIL" if "FAIL" in st else ("UNVERIFIED" if "UNVERIFIED" in st else "PASS")
    return {"status": status, "checks": checks, "max_dev_m": base.get("max_dev_m"), "at_station": base.get("at_station"), "n": base.get("n"), "tol_m": tol_m}


def _detail_scale(txt) -> float | None:
    """'1" = 40\\'' -> 480 (inches per foot * 12 / ...): detail scale denominator for a 1"=N' scale."""
    import re
    m = re.search(r"(\d+(?:\.\d+)?)\s*\"?\s*=\s*(\d+(?:\.\d+)?)\s*'", str(txt))
    return float(m.group(2)) * 12.0 / float(m.group(1)) if m else None


def verify_profile_sheet(inputs: dict, reader: Callable, workflow_id: str = "profile-design") -> dict:
    """After the engineer runs N5-N7: read back the Civil Profile named boundary, the Drawing and Sheet models and the sheet's attachment, and
    compare with the PDF settings (contract inputs.named_boundary) and the PVI range. Items OpenRoads does not expose are listed as engineer-confirmed."""
    c = load_contract(workflow_id)
    nb = dict((c.get("inputs", {}).get("named_boundary", {}) or {}).get("fields", {}))
    nb.update(inputs.get("named_boundary") or {})
    pv = normalised_pvis(inputs)
    rng = inputs.get("profile_range_m")                        # explicit range for a profile that did not come from the PVI table
    if rng and len(rng) == 2:
        pv = [{"station": float(rng[0])}, {"station": float(rng[1])}]
    checks = {}
    uor = 10000.0                                              # UOR per master unit (survey ft); boundary data is stored in UORs
    name = str(inputs.get("profile_name") or "").strip()
    bargs = {"name": inputs["alignment"]}
    if name:                                                   # the boundary lives in the Drawing model Create Drawing made, named after the profile
        try:
            for r in reader("list_models", {}):
                row = r.get("model", "")
                mn = row.split("Name=", 1)[1].split(";", 1)[0] if "Name=" in row else ""
                if "ModelType=Drawing" in row and mn and name in mn:
                    bargs["model"] = mn
                    break
        except Exception:
            pass
    try:
        d = {r["key"]: r["value"] for r in reader("profile_boundary_data", bargs) if "key" in r}
    except Exception as ex:
        d = {}
        checks["boundary_exists"] = {"status": "UNVERIFIED", "why": str(ex)[:120]}
    if d and d.get("profileBoundary") != "present":
        checks["boundary_exists"] = {"status": "FAIL", "why": "no Civil Profile named boundary on this alignment's profile model (N5 not done?)"}
    elif d:
        checks["boundary_exists"] = {"status": "PASS", "name": d.get("name"), "group": d.get("group")}
        want_ve = _num(nb.get("vertical_exaggeration"))
        bad = []
        if want_ve is not None and abs(float(d["verticalExaggeration"]) - want_ve) > 1e-9:
            bad.append(f"vertical exaggeration {d['verticalExaggeration']} != {want_ve}")
        if abs(float(d["verticalAxisSpacing"]) - float(nb.get("elevation_datum_spacing", 5))) > 1e-9:
            bad.append(f"elevation datum spacing {d['verticalAxisSpacing']} != {nb.get('elevation_datum_spacing', 5)}")
        if abs(float(d["horizontalAxisSpacing"]) - float(nb.get("station_datum_spacing", 100))) > 1e-9:
            bad.append(f"station datum spacing {d['horizontalAxisSpacing']} != {nb.get('station_datum_spacing', 100)}")
        if nb.get("use_active_vertical", True) and not (d.get("displayVertical") == "True" and d.get("displayTerrain") == "False"):
            bad.append("Use Active Vertical not set (PDF: tick Use Active Vertical, not Use Terrain)")
        if d.get("assocAlignment") != inputs["alignment"]:
            bad.append(f"boundary belongs to alignment '{d.get('assocAlignment')}'")
        checks["boundary_settings"] = {"status": "FAIL" if bad else "PASS", "problems": bad}
        if pv:
            m_per_ft = FT_TO_M
            s0, s1 = float(d["startDistance"]) / uor * m_per_ft, float(d["stopDistance"]) / uor * m_per_ft
            ok = s0 <= pv[0]["station"] + 0.05 and s1 >= pv[-1]["station"] - 0.05
            checks["boundary_range"] = {"status": "PASS" if ok else "FAIL", "boundary_m": [s0, s1], "profile_m": [pv[0]["station"], pv[-1]["station"]]}
        checks["engineer_values"] = {"status": "INFO", "profile_height_ft": float(d["profileHeight"]) / uor,
                                     "top_clearance_ft": float(d["topClearance"]) / uor, "bottom_clearance_ft": float(d["botClearance"]) / uor,
                                     "note": "PDF seeds height 50 / clearances 1 but their units depend on the dialog; compare with what you set"}
    try:
        models = [r.get("model", "") for r in reader("list_models", {})]
        atts_by_model = [(r.get("model", ""), r.get("attachment", "")) for r in reader("list_attachments", {})]
    except Exception as ex:
        models, atts_by_model = [], []
        checks["models"] = {"status": "UNVERIFIED", "why": str(ex)[:120]}
    if models or atts_by_model:
        if not name:
            checks["models"] = {"status": "UNVERIFIED", "why": "profile_name not supplied, so the Drawing/Sheet models cannot be identified"}
        else:
            def mname(row):
                return row.split("Name=", 1)[1].split(";", 1)[0] if "Name=" in row else ""
            drw = [mname(r) for r in models if "ModelType=Drawing" in r and name in mname(r)]
            sht = [mname(r) for r in models if "ModelType=Sheet" in r and name in mname(r)]
            problems = []
            if not drw:
                problems.append("no Drawing model named after the profile (N6/N7 not done?)")
            if not sht:
                problems.append("no Sheet model named after the profile (N7: Sheets (New))")
            if "'" in name:
                problems.append("the name contains an apostrophe (Create Drawing reports Invalid name)")
            checks["models"] = {"status": "FAIL" if problems else "PASS", "drawing_models": drw, "sheet_models": sht, "problems": problems}
            if drw and sht:
                want = _detail_scale(nb.get("detail_scale", "1\" = 40'"))
                hit = [(m, a) for m, a in atts_by_model if m in sht and any(f"model={x};" in a for x in drw)]
                if not hit:
                    checks["sheet_attachment"] = {"status": "FAIL", "why": "the Sheet model does not attach the Drawing model"}
                else:
                    import re
                    sc = float(re.search(r"scale=([0-9.eE+-]+)", hit[0][1]).group(1))
                    ok = want is not None and abs(sc - 1.0 / want) < 1e-6
                    checks["sheet_attachment"] = {"status": "PASS" if ok else "FAIL", "attachment_scale": sc, "expected": (1.0 / want) if want else None}
    st = [v["status"] for v in checks.values() if v["status"] != "INFO"]
    status = "FAIL" if "FAIL" in st else ("UNVERIFIED" if ("UNVERIFIED" in st or not st) else "PASS")
    return {"status": status, "checks": checks,
            "engineer_confirms": ["N4 Element Annotation with the Profile Annotation group (not readable)",
                                  "N5 Detail Scale 1\"=40' and Drawing Seed Profile Only (seed is not readable)",
                                  "N7 NYSDOT border, title block updated, one sheet per DGN, update server copy and check in (ProjectWise is not visible)"]}


def find_profile_idx(inputs: dict, reader: Callable) -> int | None:
    """Index (within the alignment's Profiles) of the profile whose station range equals the PVI table's first/last station."""
    pv = normalised_pvis(inputs)
    if len(pv) < 2:
        return None
    i = 0
    for r in reader("list_profiles", {}):
        if r.get("alignment") != inputs.get("alignment"):
            continue
        try:
            x0, x1 = float(r["x0"]), float(r["x1"])
        except (KeyError, ValueError):
            i += 1
            continue
        if abs(x0 - pv[0]["station"]) < 0.01 and abs(x1 - pv[-1]["station"]) < 0.01:
            return i
        i += 1
    return None
