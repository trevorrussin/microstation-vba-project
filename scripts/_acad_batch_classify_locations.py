"""Batch Straight/Depressed/Corner classify for HWP24QC1 Location SGD DWGs.

NW rules (from loc01-curb-classify skill):
  Depressed — PROP-D flare-to-flare on INNER curb (split verts at flare edges)
  Corner    — outside Depressed; short-chord curve (LONG_SEG >= 2.5 ft = Straight)
  Straight  — outside Depressed; colinear / long segs
  RDP 0.1 on dense FL samples; hide INNER sources; never Delete; no flare draw.

Discovers C-TOPO-FEAT AeccDbFeatureLine + PROP-D-ROAD-RAMPS per DWG (no Loc01 handles).

Usage:
  python scripts/_acad_batch_classify_locations.py --only 05-08
  python scripts/_acad_batch_classify_locations.py --all
  python scripts/_acad_batch_classify_locations.py --all --skip-done
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import time
import traceback
from collections import defaultdict
from datetime import datetime

import pythoncom
from win32com.client import Dispatch, GetObject, VARIANT

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
import _acad_reclassify_loc01_v7_keep as v7
import _acad_reclassify_others_nw_process as nw
import _acad_site_zone as site_zone

SCREENSHOTS = r"c:\repos\microstation-vba-project\Screenshots"
OUT_DIR = r"c:\repos\microstation-vba-project\outputs"
LOG_PATH = os.path.join(OUT_DIR, "curb-batch-classify.jsonl")
MANIFEST_PATH = os.path.join(OUT_DIR, "curb-batch-manifest.json")
REGIONS_PATH = r"c:\repos\microstation-vba-project\Data\corner-cad-regions.json"
PDF_CORNERS = r"c:\repos\microstation-vba-project\Data\hwp24qc1-corner-ids-from-pdfs.json"

RAMP_LAYER = "PROP-D-ROAD-RAMPS"
# Ped-ramp blocks with flare solids may sit on PROP-D *or* PRO-S / S-CURB-DROP
# (Loc10: *U255–*U260 on PRO-S-CURB-LINE — never PROP-D). Miss → Straight in drops.
RAMP_LAYERS = (
    "PROP-D-ROAD-RAMPS",
    "PRO-S-CURB-LINE",
    "PROP-S-CURB-LINE",
    "S-CURB-DROP",
)
# Inner curb may live on more than C-TOPO-FEAT (engineer 2026-09-30)
FL_LAYERS = (
    "C-TOPO-FEAT",
    "PRO-S-CURB-LINE",
    "PROP-S-CURB-LINE",
    "PRO-S-CURB-DLINE",
)
PAIR_CENTER_FT = 10.0  # Loc09 inner/outer centers were ~5 ft apart
MIN_INNER_DIAG = 18.0  # short face/cap FLs are not island inners
ISLAND_RAMP_FT = 28.0
FL_NEAR_ISLAND_FT = 55.0
SAMPLE_STEP = 0.35
RDP_EPS = 0.10
# NEVER hide or delete PROP-D ramps / source FLs / outers.
# Only delete our own steel AcDbLine segs when re-clearing classify.
HIDE_SOURCES = False


def log(msg: str) -> None:
    print(msg, flush=True)

RPC_E_CALL_REJECTED = -2147418111


def com_retry(fn, *args, tries: int = 12, label: str = "COM", **kwargs):
    """Retry on AutoCAD 'call rejected by callee' / busy."""
    last = None
    for attempt in range(tries):
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        try:
            return fn(*args, **kwargs)
        except Exception as ex:
            last = ex
            msg = str(ex)
            busy = (
                getattr(ex, "args", None)
                and len(ex.args) > 0
                and ex.args[0] == RPC_E_CALL_REJECTED
            ) or "rejected" in msg.lower() or "busy" in msg.lower() or "<unknown>" in msg
            if not busy and attempt > 2:
                raise
            log(f"  {label} retry {attempt+1}/{tries}: {ex}")
            time.sleep(1.5 + attempt * 0.75)
    raise RuntimeError(f"{label} failed after retries: {last}")


def close_all_except_keep(acad, keep_substrings=("01-02", "RECOVERED")) -> None:
    """Ensure only Loc01 (or keep list) stays open — leftover docs cause COM rejects."""
    for _ in range(8):
        try:
            n = acad.Documents.Count
        except Exception:
            time.sleep(2.0)
            continue
        closed_any = False
        for i in range(n - 1, -1, -1):
            try:
                d = acad.Documents.Item(i)
                name = d.Name
            except Exception:
                continue
            if any(k.upper() in name.upper() for k in keep_substrings):
                continue
            try:
                if not d.Saved:
                    # never discard unsaved work we did not make
                    continue
            except Exception:
                continue
            try:
                d.Close(False)
                log(f"  preclose leftover {name}")
                closed_any = True
            except Exception as ex:
                log(f"  preclose FAIL {name}: {ex}")
            time.sleep(0.5)
        if not closed_any:
            break
        time.sleep(1.0)



def append_jsonl(row: dict) -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(row, default=str) + "\n")


def list_dwgs(only: str | None) -> list[str]:
    files = []
    for name in sorted(os.listdir(SCREENSHOTS)):
        if not name.lower().endswith(".dwg"):
            continue
        if not name.upper().startswith("HWP24QC1_LOCATION"):
            continue
        if "CLASSIFIED" in name.upper():
            continue
        # skip Loc01 recovered / 01-02 / 03-04 — already handled separately
        up = name.upper()
        if "01-02" in up or "RECOVERED" in up:
            continue
        # 03-04 has no steel yet (09-30 BLOCKED) — batch only when named explicitly
        if "03-04" in up and not (only and "03-04" in only):
            continue
        if only:
            key = only.replace("_SGD", "").replace(".dwg", "").upper()
            if key not in up:
                continue
        elif "_SGD" not in up:
            continue  # --all: SGD hosts only (bare plan DWGs carry no curb)
        # Prefer SGD hosts for curb classify (steel lives on SGD)
        if "_SGD" not in up and any(
            x.upper().startswith("HWP24QC1_LOCATION") and "_SGD" in x.upper()
            for x in os.listdir(SCREENSHOTS)
            if x.lower().endswith(".dwg")
        ):
            # If a matching SGD exists for this stem, skip bare Location DWG
            stem = name.upper().replace(".DWG", "")
            sgd_name = stem + "_SGD.DWG"
            # Location_10 vs Location_10_SGD
            if not stem.endswith("_SGD"):
                candidate = name[:-4] + "_SGD.dwg"
                if os.path.exists(os.path.join(SCREENSHOTS, candidate)):
                    continue
        files.append(os.path.join(SCREENSHOTS, name))
    return files


def parse_sheet_locations(filename: str) -> list[int]:
    """Extract location numbers from HWP24QC1_Location_05-08_SGD.dwg → [5,6,7,8]."""
    base = os.path.basename(filename)
    m = re.search(r"Location[_\s-]+(.+?)(?:_SGD|_SDG)?\.dwg$", base, re.I)
    if not m:
        return []
    token = m.group(1).replace(" ", "").replace(",", "-")
    # drop trailing _SGD-like fragments already handled
    token = re.sub(r"_?(SGD|SDG)$", "", token, flags=re.I)
    locs: list[int] = []
    for part in re.split(r"[_]", token):
        if not part:
            continue
        # 89_96 or 90_91 already split; also 05-08, 58-60, 96-98
        if "-" in part:
            a, b = part.split("-", 1)
            if a.isdigit() and b.isdigit():
                locs.extend(range(int(a), int(b) + 1))
                continue
        if part.isdigit():
            locs.append(int(part))
    # unique preserve order
    seen = set()
    out = []
    for n in locs:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def load_pdf_corner_map() -> dict[int, list[dict]]:
    """location_no -> [{corner_id, pdf}, ...]"""
    with open(PDF_CORNERS, encoding="utf-8") as f:
        rows = json.load(f)
    by_loc: dict[int, list[dict]] = defaultdict(list)
    for r in rows:
        loc = int(r["location"])
        by_loc[loc].append({"corner_id": str(r["corner_id"]), "pdf": r.get("pdf", "")})
    # Master list = Working estimate 4.09 AD (read-only). PDF catalog covers ~38 locs
    # and has truncated-ID parse errors (e.g. 104405 under loc 1) — use only as fallback.
    try:
        import openpyxl

        xl = os.path.join(os.path.dirname(SCREENSHOTS), "2026.09.28_HWP24QC1 - Pre-Final Cost Estimate_Working.xlsx")
        wb = openpyxl.load_workbook(xl, read_only=True, data_only=True)
        xl_map: dict[int, list[dict]] = defaultdict(list)
        for row in wb["4.09 AD"].iter_rows(max_col=3, values_only=True):
            loc, cid = row[1], row[2]
            if isinstance(loc, (int, float)) and isinstance(cid, (int, float)):
                if not any(c["corner_id"] == str(int(cid)) for c in xl_map[int(loc)]):
                    xl_map[int(loc)].append({"corner_id": str(int(cid)), "pdf": "4.09 AD"})
        wb.close()
        for loc, lst in xl_map.items():
            by_loc[loc] = lst
    except Exception as ex:
        log(f"  corner map: 4.09 AD unavailable, PDF catalog only: {ex}")
    return by_loc


def bbox_center(e):
    mn, mx = e.GetBoundingBox()
    return (
        (float(mn[0]) + float(mx[0])) / 2,
        (float(mn[1]) + float(mx[1])) / 2,
        float(mx[0]) - float(mn[0]),
        float(mx[1]) - float(mn[1]),
    )


def collect_ramps(doc) -> list[dict]:
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbBlockReference":
                continue
            if (e.Layer or "") != RAMP_LAYER:
                continue
            cx, cy, w, h = bbox_center(e)
            rd = v7.ramp_flare_data(doc, e.Handle)
            out.append(
                {
                    "handle": e.Handle,
                    "name": e.Name or "",
                    "cx": cx,
                    "cy": cy,
                    "flares": rd["flares"],
                    "dws": rd["dws"],
                    "n_flares": len(rd["flares"]),
                }
            )
        except Exception:
            continue
    return out


def collect_fls(doc) -> list[dict]:
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if (e.Layer or "") != FL_LAYERS[0]:
                continue
            if e.ObjectName not in (
                "AeccDbFeatureLine",
                "AcDb3dPolyline",
                "AcDbPolyline",
                "AcDb2dPolyline",
            ):
                continue
            cx, cy, w, h = bbox_center(e)
            diag = math.hypot(w, h)
            out.append(
                {
                    "handle": e.Handle,
                    "type": e.ObjectName,
                    "cx": cx,
                    "cy": cy,
                    "w": w,
                    "h": h,
                    "diag": diag,
                }
            )
        except Exception:
            continue
    return out


def cluster_islands(ramps: list[dict]) -> list[list[dict]]:
    unused = list(ramps)
    islands = []
    while unused:
        seed = unused.pop(0)
        island = [seed]
        changed = True
        while changed:
            changed = False
            for r in list(unused):
                if any(
                    math.hypot(r["cx"] - s["cx"], r["cy"] - s["cy"]) < ISLAND_RAMP_FT
                    for s in island
                ):
                    island.append(r)
                    unused.remove(r)
                    changed = True
        islands.append(island)
    return islands


def pick_inner_fl(fls: list[dict], island: list[dict], exclude: set[str] | None = None) -> tuple[dict | None, list[str]]:
    """Pick INNER curb FL — Loc01: inland loop next to PROP-D ramps, never OUTER.

    Among a center-paired FL pair, INNER = closer to ramp cluster.
    Never fall back to the largest FL (Loc09 outer mistake 2026-09-30).
    """
    icx = sum(r["cx"] for r in island) / len(island)
    icy = sum(r["cy"] for r in island) / len(island)
    exclude = exclude or set()
    near = []
    for f in fls:
        if f["handle"] in exclude:
            continue
        d = math.hypot(f["cx"] - icx, f["cy"] - icy)
        if d > FL_NEAR_ISLAND_FT:
            continue
        if f["diag"] < MIN_INNER_DIAG:
            continue
        rd = min(math.hypot(f["cx"] - r["cx"], f["cy"] - r["cy"]) for r in island)
        near.append((rd, d, f))
    if not near:
        return None, []
    near.sort(key=lambda t: (t[0], t[1]))
    pool = [f for _, _, f in near]

    seen = set()
    best = None  # (rd_inner, inner, [outer])
    for f in pool:
        if f["handle"] in seen:
            continue
        mates = [
            g
            for g in pool
            if g["handle"] != f["handle"]
            and math.hypot(g["cx"] - f["cx"], g["cy"] - f["cy"]) <= PAIR_CENTER_FT
        ]
        if not mates:
            continue
        mates.sort(key=lambda g: -g["diag"])
        a, b = f, mates[0]
        rd_a = min(math.hypot(a["cx"] - r["cx"], a["cy"] - r["cy"]) for r in island)
        rd_b = min(math.hypot(b["cx"] - r["cx"], b["cy"] - r["cy"]) for r in island)
        inner, outer = (a, b) if rd_a <= rd_b else (b, a)
        seen.add(inner["handle"])
        seen.add(outer["handle"])
        rd_i = min(math.hypot(inner["cx"] - r["cx"], inner["cy"] - r["cy"]) for r in island)
        if best is None or rd_i < best[0]:
            best = (rd_i, inner, [outer["handle"]])
    if best is not None:
        return best[1], best[2]

    # No pair: closest to ramps among substantial loops — NEVER largest
    return pool[0], []


LABEL_SCOPE_FT = 60.0
TEMPLATE_KIT_FT = 150.0
OUTLINE_LAYERS = ("PRO-S-CURB-LINE", "PROP-S-CURB-LINE")
OUTLINE_NEAR_FT = 20.0
JOIN_TOL = 0.15


def _poly_pts(doc, handle) -> list[tuple[float, float]]:
    """LWPolyline vertices (bulges densified to ~0.5 ft chords)."""
    from win32com.client.dynamic import Dispatch as DynDispatch

    e = DynDispatch(doc.HandleToObject(handle))
    if e.ObjectName == "AcDbLine":
        sp, ep = e.StartPoint, e.EndPoint
        return [(sp[0], sp[1]), (ep[0], ep[1])]
    if e.ObjectName == "AcDbArc":
        cpt, r = e.Center, float(e.Radius)
        a0, a1 = float(e.StartAngle), float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        k = max(2, int((a1 - a0) * r / 0.5))
        return [(cpt[0] + r * math.cos(a0 + (a1 - a0) * j / k), cpt[1] + r * math.sin(a0 + (a1 - a0) * j / k)) for j in range(k + 1)]
    c = list(e.Coordinates)
    v = [(c[i], c[i + 1]) for i in range(0, len(c) - 1, 2)]
    n = len(v)
    if e.Closed and n > 2:
        v.append(v[0])
    out = [v[0]]
    for i in range(len(v) - 1):
        try:
            b = float(e.GetBulge(i % n))
        except Exception:
            b = 0.0
        p0, p1 = v[i], v[i + 1]
        if abs(b) > 1e-6:
            th = 4 * math.atan(b)
            ch = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
            r = ch / (2 * math.sin(abs(th) / 2))
            mx, my = (p0[0] + p1[0]) / 2, (p0[1] + p1[1]) / 2
            d = math.sqrt(max(r * r - (ch / 2) ** 2, 0))
            ux, uy = (p1[0] - p0[0]) / ch, (p1[1] - p0[1]) / ch
            sgn = 1 if (b > 0) == (abs(th) < math.pi) else -1
            cx, cy = mx - sgn * uy * d, my + sgn * ux * d
            a0 = math.atan2(p0[1] - cy, p0[0] - cx)
            k = max(2, int(abs(th) * r / 0.5))
            for j in range(1, k):
                a = a0 + th * j / k
                out.append((cx + r * math.cos(a), cy + r * math.sin(a)))
        out.append(p1)
    return out


def outline_chain(doc, fls, island, local_ramps, exclude):
    """Engineer 2026-10-01: where no Loc01-style FL exists, the INNER curb is the tan
    PRO-S-CURB-LINE outline around the ramp. Join touching pieces near the island."""
    icx = sum(r["cx"] for r in island) / len(island)
    icy = sum(r["cy"] for r in island) / len(island)
    cands = [
        f for f in fls
        if f.get("layer") in OUTLINE_LAYERS and f["type"] in ("AcDbPolyline", "AcDbLine", "AcDbArc")
        and f["handle"] not in exclude
        and math.hypot(f["cx"] - icx, f["cy"] - icy) <= OUTLINE_NEAR_FT
    ]
    if not cands:
        return None, []
    pieces = {f["handle"]: _poly_pts(doc, f["handle"]) for f in cands}
    fv = [p for r in local_ramps for fl in r["flares"] for p in fl]

    def contact(pts):
        if not fv:
            return 0
        sv = v7.stations([(x, y, 0.0) for x, y in pts])
        return sum(1 for p in fv if v7.project_station([(x, y, 0.0) for x, y in pts], sv, p[0], p[1])[1] <= v7.CURB_NEAR)

    seed = max(pieces, key=lambda h: (contact(pieces[h]), len(pieces[h])))
    if fv and contact(pieces[seed]) == 0:
        return None, []  # no curb piece touches a flare
    chain = list(pieces.pop(seed))
    used = [seed]
    near = lambda a, b: math.hypot(a[0] - b[0], a[1] - b[1]) <= JOIN_TOL
    grew = True
    while grew and pieces:
        grew = False
        for h, pts in list(pieces.items()):
            if near(chain[-1], pts[0]):
                chain += pts[1:]
            elif near(chain[-1], pts[-1]):
                chain += pts[-2::-1]
            elif near(chain[0], pts[-1]):
                chain = pts[:-1] + chain
            elif near(chain[0], pts[0]):
                chain = pts[:0:-1] + chain
            else:
                continue
            used.append(h)
            pieces.pop(h)
            grew = True
    return [(x, y, 0.0) for x, y in chain], used


def _ss_by_layer(doc, ss_name: str, layer: str):
    """acSelectionSetAll filtered to one layer (DXF 8)."""
    try:
        doc.SelectionSets.Item(ss_name).Delete()
    except Exception:
        pass
    ss = doc.SelectionSets.Add(ss_name)
    ft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])
    fd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [layer])
    # 5 = acSelectionSetAll
    ss.Select(5, None, None, ft, fd)
    return ss


def clear_steel(doc, steel_handles: list[str] | None = None) -> int:
    if steel_handles is None:
        steel_handles = []
        for lyr in v7.COLORS:
            try:
                ss = _ss_by_layer(doc, "WZTC_STEEL_CLR", lyr)
                for i in range(ss.Count):
                    steel_handles.append(ss.Item(i).Handle)
                ss.Delete()
            except Exception:
                continue
    for h in steel_handles:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(steel_handles)


def _block_solid_count(doc, e) -> int:
    """Count AcDbSolid in a block ref (flare geometry). Late-bind Name."""
    from win32com.client.dynamic import Dispatch as DynDispatch

    try:
        br = DynDispatch(e)
        if br.ObjectName != "AcDbBlockReference":
            return 0
        blk = doc.Blocks.Item(br.Name)
        n = 0
        for j in range(blk.Count):
            try:
                if blk.Item(j).ObjectName == "AcDbSolid":
                    n += 1
            except Exception:
                continue
        return n
    except Exception:
        return 0


def _block_name(e) -> str:
    from win32com.client.dynamic import Dispatch as DynDispatch

    try:
        return DynDispatch(e).Name or ""
    except Exception:
        return ""


def scan_doc(doc) -> tuple[list[dict], list[dict], list[str]]:
    """Layer SelectionSets → ramps, FLs, existing steel handles (fast)."""
    t0 = time.time()
    ramps_raw = []
    fls = []
    steel_handles = []
    seen_ramp = set()
    log("  SelectionSet scan (TOPO + ramp layers + steel) …")

    # Unhide any previously hidden FLs / ramps (never leave color-coded content off)
    try:
        ms = doc.ModelSpace
        for i in range(ms.Count):
            try:
                e = ms.Item(i)
                lyr = e.Layer or ""
                if lyr in RAMP_LAYERS or lyr in FL_LAYERS or "Steel Faced" in lyr:
                    if hasattr(e, "Visible") and not e.Visible:
                        e.Visible = True
            except Exception:
                continue
    except Exception as ex:
        log(f"    unhide pass warn: {ex}")

    for lyr in FL_LAYERS:
        try:
            ss = _ss_by_layer(doc, "WZTC_FL", lyr)
            log(f"    {lyr} hits={ss.Count}")
            for i in range(ss.Count):
                try:
                    e = ss.Item(i)
                    on = e.ObjectName or ""
                    if on not in (
                        "AeccDbFeatureLine",
                        "AcDb3dPolyline",
                        "AcDbPolyline",
                        "AcDb2dPolyline",
                    ) and not (on in ("AcDbLine", "AcDbArc") and lyr in OUTLINE_LAYERS):
                        continue  # Line/Arc pieces: outline chains only (Loc13)
                    cx, cy, w, h = bbox_center(e)
                    fls.append(
                        {
                            "handle": e.Handle,
                            "type": on,
                            "layer": lyr,
                            "cx": cx,
                            "cy": cy,
                            "w": w,
                            "h": h,
                            "diag": math.hypot(w, h),
                        }
                    )
                except Exception:
                    continue
            ss.Delete()
        except Exception as ex:
            log(f"    {lyr} SelectionSet FAIL: {ex}")

    # Ped ramps = block refs with flare solids (≥2), any RAMP_LAYERS.
    # PROP-D-ROAD-RAMPS: always take inserts (Loc01/05–08).
    # Other layers: only inserts with ≥2 solids (Loc10 *U* on PRO-S-CURB-LINE).
    for lyr in RAMP_LAYERS:
        try:
            ss = _ss_by_layer(doc, "WZTC_RAMP", lyr)
            log(f"    {lyr} hits={ss.Count}")
            n_keep = 0
            for i in range(ss.Count):
                try:
                    e = ss.Item(i)
                    if e.ObjectName != "AcDbBlockReference":
                        continue
                    h = e.Handle
                    if h in seen_ramp:
                        continue
                    solids = _block_solid_count(doc, e)
                    if lyr != RAMP_LAYER and solids < 2:
                        continue
                    cx, cy, w, hbox = bbox_center(e)
                    # Skip paper-space / border-scale junk (sheet coords ~tens of ft)
                    if abs(cx) < 500 and abs(cy) < 500:
                        continue
                    seen_ramp.add(h)
                    ramps_raw.append(
                        {
                            "handle": h,
                            "name": _block_name(e),
                            "cx": cx,
                            "cy": cy,
                            "layer": lyr,
                            "n_solids": solids,
                        }
                    )
                    n_keep += 1
                except Exception:
                    continue
            log(f"      -> ramp blocks kept={n_keep}")
            ss.Delete()
        except Exception as ex:
            log(f"    {lyr} SelectionSet FAIL: {ex}")

    for lyr in v7.COLORS:
        try:
            ss = _ss_by_layer(doc, "WZTC_STEEL", lyr)
            for i in range(ss.Count):
                try:
                    steel_handles.append(ss.Item(i).Handle)
                except Exception:
                    continue
            ss.Delete()
        except Exception:
            continue

    log(
        f"  scan done {time.time()-t0:.1f}s "
        f"ramps={len(ramps_raw)} fls={len(fls)} steel={len(steel_handles)}"
    )
    for r in ramps_raw:
        try:
            rd = com_retry(v7.ramp_flare_data, doc, r["handle"], label=f"flare {r['handle']}")
            r["flares"] = rd["flares"]
            r["dws"] = rd["dws"]
            r["n_flares"] = len(rd["flares"])
            if not r.get("name"):
                r["name"] = rd.get("name") or r["handle"]
        except Exception as ex:
            log(f"  ramp {r['handle']} flare read fail: {ex}")
            r["flares"] = []
            r["dws"] = []
            r["n_flares"] = 0
    # Drop inserts that could not yield flare measure (except keep for island seed if PROP-D)
    measurable = [r for r in ramps_raw if r.get("n_flares", 0) >= 1 or r.get("flares")]
    if measurable:
        log(f"  ramps with flare geometry={len(measurable)}/{len(ramps_raw)}")
        ramps_raw = measurable
    return ramps_raw, fls, steel_handles


def already_classified(doc) -> bool:
    n_steel = 0
    for lyr in v7.COLORS:
        try:
            ss = _ss_by_layer(doc, "WZTC_STEEL_CHK", lyr)
            n_steel += ss.Count
            ss.Delete()
            if n_steel >= 20:
                return True
        except Exception:
            continue
    return False


def classify_doc(doc, sheet_locs: list[int], pdf_map: dict) -> dict:
    # Clear any stuck command before SelectionSets
    try:
        doc.SendCommand(chr(27) + chr(27))
        time.sleep(0.3)
    except Exception:
        pass
    log("  ensure_layers…")
    com_retry(v7.ensure_layers, doc, label="ensure_layers")
    log("  ensure_layers done")
    ramps_raw, fls, steel_handles = scan_doc(doc)
    # Site zone: drop template ramp kit / E52F-FC58 copies that sit off the survey
    cells = site_zone.survey_cells(doc, log=log)
    if cells is None:
        log("  VERIFY FAIL (UNCERTAIN): no loaded survey xref — cannot establish site zone")
        return {
            "cleared_steel": 0, "n_ramps": len(ramps_raw), "n_fls": len(fls),
            "n_islands": 0, "n_islands_ok": 0, "verify_ok": False,
            "reason": "no_survey_xref", "inners": [], "outers_ref": [],
            "totals": {k: 0.0 for k in v7.COLORS}, "counts": {k: 0 for k in v7.COLORS},
            "islands": [], "total_lf": 0.0,
        }
    # Template kit signature: PRO-S E52F + FC58 pair at fixed offset (66.5, 69.7) moved with
    # the copied PROP-D kit. Kit ramps sit within ~150 ft; real sites were 350-700 ft away.
    # Loc61: FC58 absent, E52F alone + kit circle E51A (r~97 ft) — either anchor is enough.
    tpl = {f["handle"]: f for f in fls if f["handle"] in ("E52F", "FC58") and f.get("layer") == "PRO-S-CURB-LINE"}
    if tpl:
        kit = [r for r in ramps_raw
               if min(math.hypot(r["cx"] - t["cx"], r["cy"] - t["cy"]) for t in tpl.values()) <= TEMPLATE_KIT_FT]
        if kit:
            log(f"  template-kit dropped ramps={[(r['name'], round(r['cx']), round(r['cy'])) for r in kit]}")
        ramps_raw = [r for r in ramps_raw if r not in kit]
        fls = [f for f in fls if f["handle"] not in tpl]
    off_r = [r for r in ramps_raw if not site_zone.on_site(cells, r["cx"], r["cy"])]
    off_f = [f for f in fls if not site_zone.on_site(cells, f["cx"], f["cy"])]
    ramps_raw = [r for r in ramps_raw if r not in off_r]
    fls = [f for f in fls if f not in off_f]
    log(
        f"  site-zone kept ramps={len(ramps_raw)} fls={len(fls)}; "
        f"dropped off-site ramps={[(r['name'], round(r['cx']), round(r['cy'])) for r in off_r]} "
        f"fls={[f['handle'] for f in off_f]}"
    )
    cleared = clear_steel(doc, steel_handles)
    log(f"  cleared steel segs={cleared}")
    ramp_data = []
    for r in ramps_raw:
        ramp_data.append(
            {
                "handle": r["handle"],
                "name": r["name"],
                "flares": r.get("flares") or [],
                "dws": r.get("dws") or [],
                "area": None,
                "cx": r["cx"],
                "cy": r["cy"],
            }
        )
    islands = cluster_islands(ramps_raw) if ramps_raw else []

    # orphan FLs (no ramps): treat each large FL pair cluster as island later
    if not islands and fls:
        unused = [f for f in fls if f["diag"] >= MIN_INNER_DIAG]
        while unused:
            seed = unused.pop(0)
            grp = [seed]
            changed = True
            while changed:
                changed = False
                for f in list(unused):
                    if any(
                        math.hypot(f["cx"] - s["cx"], f["cy"] - s["cy"]) < 40
                        for s in grp
                    ):
                        grp.append(f)
                        unused.remove(f)
                        changed = True
            islands.append(
                [
                    {
                        "handle": f"synth-{grp[0]['handle']}",
                        "name": "synth",
                        "cx": sum(g["cx"] for g in grp) / len(grp),
                        "cy": sum(g["cy"] for g in grp) / len(grp),
                        "flares": [],
                        "dws": [],
                        "n_flares": 0,
                    }
                ]
            )

    # sort islands N→S then W→E for Corner ID assignment
    islands.sort(
        key=lambda isl: (
            -sum(r["cy"] for r in isl) / len(isl),
            sum(r["cx"] for r in isl) / len(isl),
        )
    )
    log(f"  islands={len(islands)}")

    totals = {k: 0.0 for k in v7.COLORS}
    counts = {k: 0 for k in v7.COLORS}
    claimed = set()
    used_inners: list[str] = []
    used_outers: list[str] = []
    island_rows = []
    ms = doc.ModelSpace

    corner_queue: list[dict] = []
    for loc in sheet_locs:
        for c in pdf_map.get(loc, []):
            corner_queue.append({"location_no": loc, **c})

    labels = site_zone.corner_labels(doc, log=log)
    # Plan CORNER ID labels = in-scope corners. Template ramp kit can land inside the
    # survey zone (Loc21: ~400 ft from labels) — drop islands with no label nearby.
    if labels:
        keep, drop = [], []
        for isl in islands:
            ix = sum(r["cx"] for r in isl) / len(isl)
            iy = sum(r["cy"] for r in isl) / len(isl)
            dmin = min(math.hypot(lx - ix, ly - iy) for _, lx, ly in labels)
            (keep if dmin <= LABEL_SCOPE_FT else drop).append((isl, ix, iy, dmin))
        if drop:
            log(f"  label-scope dropped islands (no CORNER ID label within {LABEL_SCOPE_FT:.0f} ft): "
                f"{[(round(x), round(y), round(d)) for _, x, y, d in drop]}")
        islands = [k[0] for k in keep]
        kept_ramps = {r["handle"] for isl in islands for r in isl}
        ramps_raw = [r for r in ramps_raw if r["handle"] in kept_ramps]
        log(f"  islands in label scope={len(islands)}")
    else:
        log("  WARN: no CORNER ID labels in plan xref — island scope from site zone only")
    # Merge islands whose flares touch the same curb line (Loc58-60: one 70 ft PRO-S
    # outline with a ramp at each end clustered as 2 islands; reuse guard starved one).
    _cache: dict = {}

    def _contact_fl(isl):
        fvv = [p for r in isl for fl in (r.get("flares") or []) for p in fl]
        if not fvv:
            return None
        best = None
        for f in fls:
            if f["type"] in ("AcDbLine", "AcDbArc"):
                continue
            if min(math.hypot(f["cx"] - r["cx"], f["cy"] - r["cy"]) for r in isl) > FL_NEAR_ISLAND_FT:
                continue
            if f["handle"] not in _cache:
                try:
                    d = v7.sample_fl_vlax(doc, f["handle"], step=SAMPLE_STEP)
                    d3 = [(q[0], q[1], 0.0) for q in d]
                    _cache[f["handle"]] = (d3, v7.stations(d3)) if len(d3) >= 3 else None
                except Exception:
                    _cache[f["handle"]] = None
            c3 = _cache[f["handle"]]
            if not c3:
                continue
            n = sum(1 for q in fvv if v7.project_station(c3[0], c3[1], q[0], q[1])[1] <= v7.CURB_NEAR)
            if n and (best is None or n > best[0]):
                best = (n, f)
        return best[1] if best else None

    by_fl: dict = {}
    for isl in islands:
        f = _contact_fl(isl)
        by_fl.setdefault(f["handle"] if f else id(isl), []).append((isl, f))
    shared_pref: dict = {}
    merged = []
    for key, grp in by_fl.items():
        if len(grp) > 1 and grp[0][1] is not None:
            m = [r for isl, _ in grp for r in isl]
            merged.append(m)
            shared_pref[id(m)] = grp[0][1]
            log(f"  merged {len(grp)} islands sharing curb {key}")
        else:
            merged.extend(isl for isl, _ in grp)
    islands = merged
    for ii, island in enumerate(islands):
        icx = sum(r["cx"] for r in island) / len(island)
        icy = sum(r["cy"] for r in island) / len(island)
        ramp_handles = [r["handle"] for r in island if not str(r["handle"]).startswith("synth")]
        local_ramps = [rd for rd in ramp_data if rd["handle"] in ramp_handles]

        if id(island) in shared_pref:
            inner, outers = shared_pref[id(island)], []
        else:
            inner, outers = pick_inner_fl(
                [f for f in fls if f["type"] not in ("AcDbLine", "AcDbArc")],
                island, exclude=set(used_inners) | set(used_outers))
        dense = None
        fv = [p for r in local_ramps for fl in r["flares"] for p in fl]
        if inner is not None:
            dense = v7.sample_fl_vlax(doc, inner["handle"], step=SAMPLE_STEP)
            if fv and len(dense) >= 3:
                d3 = [(p[0], p[1], 0.0) for p in dense]
                sv = v7.stations(d3)
                touch = sum(1 for p in fv if v7.project_station(d3, sv, p[0], p[1])[1] <= v7.CURB_NEAR)
                if touch == 0:
                    log(f"  island{ii+1}: FL {inner['handle']} touches no flare — not INNER; try PRO-S outline")
                    inner, outers, dense = None, [], None
        if inner is None:
            chain, pieces = outline_chain(doc, fls, island, local_ramps, set(used_inners) | set(used_outers))
            if chain and len(chain) >= 3:
                tag = "+".join(pieces)
                used_inners.extend(pieces)
                log(f"  island{ii+1} center=({icx:.1f},{icy:.1f}) ramps={len(local_ramps)} "
                    f"OUTLINE inner={tag} verts={len(chain)}")
                ok, used = nw.classify_nw_rules(ms, chain, local_ramps, totals, counts, tag, claimed)
                if ok:
                    claimed.update(used)
                corner_meta = None
                allowed = {str(c["corner_id"]): c for c in corner_queue}
                best_lbl = None
                for cid, lx, ly in labels:
                    if cid in allowed:
                        d = math.hypot(lx - icx, ly - icy)
                        if d <= 40.0 and (best_lbl is None or d < best_lbl[0]):
                            best_lbl = (d, cid)
                if best_lbl:
                    corner_meta = allowed[best_lbl[1]]
                elif len(islands) == 1 and len(allowed) == 1:
                    corner_meta = next(iter(allowed.values()))
                island_rows.append({"island": ii + 1, "center": [icx, icy], "inner": tag,
                                    "outers": [], "ramps": ramp_handles, "ok": ok,
                                    "corner": corner_meta, "mode": "outline"})
                continue
            log(f"  island{ii+1} ({icx:.0f},{icy:.0f}): no inner FL")
            island_rows.append(
                {"island": ii + 1, "center": [icx, icy], "ok": False, "reason": "no_inner"}
            )
            continue

        used_outers.extend(outers)
        h = inner["handle"]
        used_inners.append(h)  # reserve immediately — never reuse across islands
        log(
            f"  island{ii+1} center=({icx:.1f},{icy:.1f}) "
            f"ramps={len(local_ramps)} inner={h} diag={inner['diag']:.1f}"
            f" layer={inner.get('layer','?')}"
        )

        if dense is None:
            dense = v7.sample_fl_vlax(doc, h, step=SAMPLE_STEP)
        log(f"    sampled n={len(dense)}")
        if len(dense) < 3:
            island_rows.append(
                {
                    "island": ii + 1,
                    "center": [icx, icy],
                    "inner": h,
                    "ok": False,
                    "reason": "sample_fail",
                }
            )
            continue
        pts = nw.rdp(dense, RDP_EPS)
        pts = [(p[0], p[1], p[2] if len(p) > 2 else 0.0) for p in pts]
        ok, used = nw.classify_nw_rules(
            ms, pts, local_ramps, totals, counts, h, claimed
        )
        if ok:
            claimed.update(used)
            # used_inners already reserved at pick time
            # Do NOT hide sources/ramps — engineer needs color-coded geometry visible
            if HIDE_SOURCES:
                v7.hide_entity(doc, h, f"island{ii+1} classified")
        # Corner ID by nearest plan label arrow (sort-order queue mis-assigns multi-corner sheets)
        corner_meta = None
        allowed = {str(c["corner_id"]): c for c in corner_queue}
        best_lbl = None
        for cid, lx, ly in labels:
            if cid not in allowed:
                continue
            d = math.hypot(lx - icx, ly - icy)
            if d <= 40.0 and (best_lbl is None or d < best_lbl[0]):
                best_lbl = (d, cid)
        if best_lbl:
            corner_meta = allowed[best_lbl[1]]
        elif len(islands) == 1 and len(allowed) == 1:
            corner_meta = next(iter(allowed.values()))  # 1 island / 1 Corner ID: unambiguous
        else:
            log(f"    island{ii+1}: no Corner ID label within 40 ft — region not assigned")
        island_rows.append(
            {
                "island": ii + 1,
                "center": [icx, icy],
                "inner": h,
                "outers": outers,
                "ramps": ramp_handles,
                "ok": ok,
                "corner": corner_meta,
            }
        )

    try:
        doc.Regen(1)
    except Exception:
        pass

    n_ok = sum(1 for r in island_rows if r.get("ok"))
    dep_lf = float(totals.get(v7.LAYER_DEPRESSED, 0.0))
    # Every ramp island must classify — Loc09 one-island pass is a FAIL
    # Zero on-site islands = nothing classified → UNCERTAIN (FAIL), never silent PASS
    verify_ok = len(islands) > 0 and n_ok == len(islands) and sum(totals.values()) > 1.0
    if not islands:
        log("  VERIFY FAIL (UNCERTAIN): 0 on-site islands — investigate sheet content")
    # Loc01 rule: if ped-ramp flares exist, Depressed must be measured.
    # Zero Depressed with ramps present → Straight wrongly fills drop zones (Loc10).
    if len(ramps_raw) > 0 and dep_lf < 0.5:
        verify_ok = False
        log(
            f"  VERIFY FAIL Depressed={dep_lf:.2f} with {len(ramps_raw)} ramp(s) — "
            f"Straight must not occupy flare-to-flare spans"
        )
    if not verify_ok:
        log(
            f"  VERIFY FAIL islands_ok={n_ok}/{len(islands)} "
            f"ramps={len(ramps_raw)} Depressed={dep_lf:.1f} "
            f"total_lf={sum(totals.values()):.1f}"
        )
    else:
        log(
            f"  VERIFY PASS islands_ok={n_ok}/{len(islands)} "
            f"Depressed={dep_lf:.1f}"
        )

    return {
        "cleared_steel": cleared,
        "n_ramps": len(ramps_raw),
        "n_fls": len(fls),
        "n_islands": len(islands),
        "n_islands_ok": n_ok,
        "verify_ok": verify_ok,
        "inners": used_inners,
        "outers_ref": used_outers,
        "totals": {k: round(v, 2) for k, v in totals.items()},
        "counts": counts,
        "islands": island_rows,
        "total_lf": round(sum(totals.values()), 2),
    }

def update_regions(result: dict, sheet: str) -> int:
    if not os.path.isfile(REGIONS_PATH):
        data = {"_comment": "", "corners": {}}
    else:
        with open(REGIONS_PATH, encoding="utf-8") as f:
            data = json.load(f)
    corners = data.setdefault("corners", {})
    n = 0
    for isl in result.get("islands", []):
        meta = isl.get("corner")
        if not meta or not isl.get("ok"):
            continue
        cid = str(meta["corner_id"])
        cx, cy = isl["center"]
        corners[cid] = {
            "location_no": meta["location_no"],
            "center": [round(cx, 1), round(cy, 1)],
            "radius_ft": 40.0,
            "note": f"auto from {os.path.basename(sheet)} island{isl['island']}",
            "inner_handle": isl.get("inner"),
        }
        n += 1
    data["_updated"] = datetime.now().isoformat(timespec="seconds")
    with open(REGIONS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.write("\n")
    return n


def open_doc(acad, path: str):
    key = os.path.basename(path).lower()
    close_all_except_keep(acad)
    last = None
    for attempt in range(12):
        try:
            pythoncom.PumpWaitingMessages()
            n = acad.Documents.Count
            for i in range(n):
                d = acad.Documents.Item(i)
                if d.Name.lower() == key:
                    d.Activate()
                    time.sleep(0.5)
                    return d, "already_open"
            d = acad.Documents.Open(path, False)
            time.sleep(1.0)
            try:
                d.Activate()
            except Exception:
                pass
            return d, "opened"
        except Exception as ex:
            last = ex
            log(f"  Open/enum retry {attempt+1}/12: {ex}")
            time.sleep(2.5 + attempt * 0.8)
            close_all_except_keep(acad)
    raise RuntimeError(f"Open failed after retries: {last}")


def try_save(doc, path: str) -> tuple[bool, str]:
    """Save in place to the same DWG only (no SaveAs copies)."""
    # Preserve prior DWG + .bak (AutoCAD save overwrites .bak) — never lose a file
    import shutil

    bdir = os.path.join(SCREENSHOTS, "_presave_backup_" + datetime.now().strftime("%Y%m%d"))
    os.makedirs(bdir, exist_ok=True)
    stamp = datetime.now().strftime("%H%M%S")
    for src in (path, os.path.splitext(path)[0] + ".bak"):
        if os.path.exists(src):
            b, ext = os.path.splitext(os.path.basename(src))
            shutil.copy2(src, os.path.join(bdir, f"{b}_{stamp}{ext}"))
    for attempt in range(4):
        try:
            doc.Save()
            log(f"  saved in place -> {os.path.basename(path)}")
            return True, path
        except Exception as ex:
            log(f"  Save retry {attempt+1}/4: {ex}")
            time.sleep(1.0 + attempt)
    log("  SAVE FAIL — leaving unsaved; will still close")
    return False, ""


def close_doc(doc) -> None:
    try:
        doc.Close(False)  # False = do not save again (already tried Save)
        log("  closed")
    except Exception as ex:
        log(f"  close FAIL: {ex}")
    time.sleep(2.5)
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def process_one(acad, path: str, pdf_map: dict, skip_done: bool, save: bool, keep_open: bool = False) -> dict:
    name = os.path.basename(path)
    log(f"\n======== {name} ========")
    t0 = time.time()
    doc, how = open_doc(acad, path)
    # ModelSpace can reject while PW/regen settles
    ms_count = None
    for attempt in range(6):
        try:
            ms_count = doc.ModelSpace.Count
            break
        except Exception as ex:
            log(f"  ModelSpace retry {attempt+1}/6: {ex}")
            time.sleep(1.5 + attempt)
    log(f"  open={how} ms={ms_count}")
    if skip_done and already_classified(doc):
        log("  skip_done (steel already present)")
        row = {
            "file": name,
            "status": "skip_done",
            "elapsed_s": round(time.time() - t0, 1),
        }
        append_jsonl(row)
        close_doc(doc)
        return row

    sheet_locs = parse_sheet_locations(name)
    log(f"  locations={sheet_locs}")
    saved = False
    saved_path = ""
    try:
        result = classify_doc(doc, sheet_locs, pdf_map)
        nreg = 0
        verify_ok = bool(result.get("verify_ok", True))
        # Never persist Straight-in-depressed / incomplete island work
        if verify_ok:
            nreg = update_regions(result, path)
            if save:
                saved, saved_path = try_save(doc, path)
        else:
            log("  NOT SAVED — verify failed (fix before next sheet)")
            # Strip bad steel so the open DWG is not left looking "done"
            try:
                cleared_bad = clear_steel(doc)
                log(f"  cleared bad steel after verify fail: {cleared_bad}")
            except Exception as ex:
                log(f"  clear-after-fail warn: {ex}")
        row = {
            "file": name,
            "status": "ok" if verify_ok else "verify_fail",
            "saved": saved,
            "saved_path": saved_path,
            "locations": sheet_locs,
            "regions_updated": nreg,
            "elapsed_s": round(time.time() - t0, 1),
            **result,
        }
        log(
            f"  LF Straight={result['totals'].get(v7.LAYER_STRAIGHT, 0):.1f} "
            f"Depressed={result['totals'].get(v7.LAYER_DEPRESSED, 0):.1f} "
            f"Corner={result['totals'].get(v7.LAYER_CORNER, 0):.1f} "
            f"TOTAL={result['total_lf']:.1f}  regions+={nreg}"
        )
    except Exception as ex:
        row = {
            "file": name,
            "status": "error",
            "error": str(ex),
            "traceback": traceback.format_exc()[-1500:],
            "elapsed_s": round(time.time() - t0, 1),
        }
        log(f"  ERROR: {ex}")
    append_jsonl(row)
    # Leave DWG open on blocking failure so agent/engineer can investigate
    st = row.get("status")
    if st in ("verify_fail", "error"):
        log(
            f"  STATUS=BLOCKED — DWG left open for Failure Recovery Loop "
            f"(do NOT open next sheet)"
        )
        return row
    if keep_open:
        log("  kept open for visual QA")
        return row
    close_doc(doc)
    return row


def _write_blocked(path: str, row: dict) -> None:
    pause = os.path.join(OUT_DIR, "curb-batch-paused.txt")
    with open(pause, "w", encoding="utf-8") as f:
        f.write("blocked\n")
        f.write(f"file={row.get('file')}\n")
        f.write(f"status={row.get('status')}\n")
        f.write(f"error={row.get('error', '')}\n")
        f.write(
            "ACTION: Failure Recovery Loop — diagnose vs known-good Loc; "
            "do not resume batch until PASS or engineer UNRESOLVED.\n"
        )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="Substring match e.g. 05-08 or 14")
    ap.add_argument("--all", action="store_true", help="All Screenshots Location DWGs")
    ap.add_argument("--skip-done", action="store_true", help="Skip if steel lines already present")
    ap.add_argument("--force", action="store_true", help="Reclassify even if steel present")
    ap.add_argument("--no-save", action="store_true")
    ap.add_argument("--keep-open", action="store_true", help="Leave DWG open after PASS (visual QA)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument(
        "--start-after",
        help="Skip files until after this substring (e.g. 13 to start at 14)",
    )
    ap.add_argument(
        "--exclude",
        action="append",
        default=[],
        help="Substring to exclude (repeatable), e.g. --exclude 13",
    )
    ap.add_argument(
        "--continue-on-fail",
        action="store_true",
        help="DANGEROUS: keep batching after verify_fail (default: STOP — Failure Recovery)",
    )
    args = ap.parse_args()
    if not args.all and not args.only:
        ap.error("pass --only NAME or --all")

    os.makedirs(OUT_DIR, exist_ok=True)
    pdf_map = load_pdf_corner_map()
    files = list_dwgs(args.only)
    if args.exclude:
        files = [
            f
            for f in files
            if not any(ex.upper() in os.path.basename(f).upper() for ex in args.exclude)
        ]
    if args.start_after:
        key = args.start_after.upper()
        out = []
        seen = False
        for f in files:
            if seen:
                out.append(f)
            elif key in os.path.basename(f).upper():
                seen = True  # skip this one; take following
        files = out
    if args.limit:
        files = files[: args.limit]
    log(f"files={len(files)}")
    for f in files:
        log(f"  {os.path.basename(f)}")

    pythoncom.CoInitialize()
    try:
        acad = GetObject(None, "AutoCAD.Application")
    except Exception:
        acad = Dispatch("AutoCAD.Application")
    acad.Visible = True

    summary = []
    blocked = None
    for path in files:
        try:
            row = process_one(
                acad, path, pdf_map, skip_done=(args.skip_done and not args.force), save=not args.no_save,
                keep_open=args.keep_open,
            )
        except Exception as ex:
            log(f"  FATAL sheet error: {ex}")
            row = {
                "file": os.path.basename(path),
                "status": "error",
                "error": str(ex),
                "elapsed_s": 0,
            }
            append_jsonl(row)
        summary.append(
            {
                "file": row.get("file"),
                "status": row.get("status"),
                "total_lf": row.get("total_lf"),
                "elapsed_s": row.get("elapsed_s"),
            }
        )
        st = row.get("status")
        if st in ("verify_fail", "error") and not args.continue_on_fail:
            blocked = row
            log(
                f"\n===== BATCH STOPPED (STATUS=BLOCKED) on {row.get('file')} =====\n"
                f"Failure Recovery Loop required before any next sheet.\n"
                f"Do not confuse batch progress with engineering progress."
            )
            _write_blocked(path, row)
            break
        time.sleep(3.0)  # let AutoCAD settle between sheets

    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "when": datetime.now().isoformat(timespec="seconds"),
                "n": len(summary),
                "blocked": blocked,
                "summary": summary,
            },
            f,
            indent=2,
        )
    ok = sum(1 for s in summary if s["status"] == "ok")
    skip = sum(1 for s in summary if s["status"] == "skip_done")
    err = sum(1 for s in summary if s["status"] in ("error", "verify_fail"))
    log(f"\n===== BATCH DONE ok={ok} skip={skip} err={err} blocked={bool(blocked)} =====")
    log(f"log: {LOG_PATH}")
    log(f"manifest: {MANIFEST_PATH}")
    if blocked or err:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
