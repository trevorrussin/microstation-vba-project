"""Patch _acad_batch_classify_locations.py INNER pick + multi-layer + no-hide."""
from __future__ import annotations

from pathlib import Path

P = Path(r"c:\repos\microstation-vba-project\scripts\_acad_batch_classify_locations.py")
text = P.read_text(encoding="utf-8")

NEW_PICK = r'''def pick_inner_fl(fls: list[dict], island: list[dict]) -> tuple[dict | None, list[str]]:
    """Pick INNER curb FL — Loc01: inland loop next to PROP-D ramps, never OUTER.

    Among a center-paired FL pair, INNER = closer to ramp cluster.
    Never fall back to the largest FL (Loc09 outer mistake 2026-09-30).
    """
    icx = sum(r["cx"] for r in island) / len(island)
    icy = sum(r["cy"] for r in island) / len(island)
    near = []
    for f in fls:
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


'''

start = text.index("def pick_inner_fl(")
end = text.index("def _ss_by_layer(")
text = text[:start] + NEW_PICK + text[end:]

# Multi-layer FL scan: replace C-TOPO-FEAT-only block
old_scan_topo = '''    try:
        ss = _ss_by_layer(doc, "WZTC_TOPO", TOPO_LAYER)
        log(f"    C-TOPO-FEAT hits={ss.Count}")
        for i in range(ss.Count):
            try:
                e = ss.Item(i)
                on = e.ObjectName or ""
                if on not in (
                    "AeccDbFeatureLine",
                    "AcDb3dPolyline",
                    "AcDbPolyline",
                    "AcDb2dPolyline",
                ):
                    continue
                cx, cy, w, h = bbox_center(e)
                fls.append(
                    {
                        "handle": e.Handle,
                        "type": on,
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
        log(f"    TOPO SelectionSet FAIL: {ex}")
'''

new_scan_topo = '''    # Unhide any previously hidden FLs / ramps (never leave color-coded content off)
    try:
        ms = doc.ModelSpace
        for i in range(ms.Count):
            try:
                e = ms.Item(i)
                lyr = e.Layer or ""
                if lyr == RAMP_LAYER or lyr in FL_LAYERS or "Steel Faced" in lyr:
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
                    ):
                        continue
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
'''

if old_scan_topo not in text:
    raise SystemExit("scan topo block not found")
text = text.replace(old_scan_topo, new_scan_topo, 1)

# Remove hide_entity calls
text = text.replace(
    """        if ok:
            claimed.update(used)
            used_inners.append(h)
            v7.hide_entity(doc, h, f"island{ii+1} classified")
""",
    """        if ok:
            claimed.update(used)
            used_inners.append(h)
            # Do NOT hide sources/ramps — engineer needs color-coded geometry visible
            if HIDE_SOURCES:
                v7.hide_entity(doc, h, f"island{ii+1} classified")
""",
)

# Fix TOPO_LAYER reference if any remain
text = text.replace("TOPO_LAYER", "FL_LAYERS[0]")

# Add --force flag near skip-done
if "--force" not in text:
    text = text.replace(
        'ap.add_argument("--skip-done", action="store_true", help="Skip if steel lines already present")',
        'ap.add_argument("--skip-done", action="store_true", help="Skip if steel lines already present")\n'
        '    ap.add_argument("--force", action="store_true", help="Reclassify even if steel present")',
    )
    text = text.replace(
        "skip_done=args.skip_done, save=not args.no_save",
        "skip_done=(args.skip_done and not args.force), save=not args.no_save",
    )

P.write_text(text, encoding="utf-8")
print("patched OK", P)
