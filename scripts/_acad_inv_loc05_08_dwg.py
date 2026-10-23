"""Read-only COM inventory: HWP24QC1 Location 05-08 vs Loc01 curb structure."""
from __future__ import annotations

import re
import sys
from collections import defaultdict

import pythoncom
from win32com.client import GetObject

PATH = r"c:\repos\microstation-vba-project\Screenshots\HWP24QC1_Location_05-08_SGD.dwg"

STEEL_LAYERS = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]

TOPO_LIKE = re.compile(r"topo|feat|c-topo", re.I)
LAYER_FILTER = re.compile(r"curb|steel|topo|feat|ramp|flare|dws|prop", re.I)

FEAT_TYPES = {
    "AeccDbFeatureLine",
    "AcDb3dPolyline",
    "AcDbPolyline",
    "AcDb2dPolyline",
    "AcDbLine",
}

XDATA_APPS = ("HWP24QC1", "CORNER_ID", "AECOM", "CIVIL", "AECEXPLORER")


def approx_length(e) -> float | None:
    try:
        return float(e.Length)
    except Exception:
        pass
    try:
        coords = e.Coordinates
        if coords is None:
            return None
        pts = list(coords)
        if len(pts) < 4:
            return None
        # 2d or 3d flat list
        step = 3 if len(pts) % 3 == 0 else 2
        total = 0.0
        for i in range(step, len(pts), step):
            dx = pts[i] - pts[i - step]
            dy = pts[i + 1] - pts[i - step + 1]
            if step == 3:
                dz = pts[i + 2] - pts[i - step + 2]
                total += (dx * dx + dy * dy + dz * dz) ** 0.5
            else:
                total += (dx * dx + dy * dy) ** 0.5
        return total
    except Exception:
        return None


def vertex_count(e) -> int | str:
    for attr in ("NumberOfVertices", "NumVertices"):
        try:
            return int(getattr(e, attr))
        except Exception:
            pass
    try:
        c = e.Coordinates
        if c is None:
            return "?"
        n = len(list(c))
        if n % 3 == 0:
            return n // 3
        if n % 2 == 0:
            return n // 2
        return n
    except Exception:
        return "?"


def get_xdata_summary(e) -> list[str]:
    out = []
    try:
        xd = e.GetXData("")
        if xd is None:
            return out
        # xd is variant array: (type, value) pairs
        raw = list(xd) if hasattr(xd, "__iter__") else []
        app = None
        for i in range(0, len(raw) - 1, 2):
            code = raw[i]
            val = raw[i + 1]
            if code == 1001:
                app = str(val)
            elif app and app in XDATA_APPS:
                out.append(f"{app}:{val}")
    except Exception:
        pass
    return out


def open_readonly(acad, path: str):
    name_key = path.replace("/", "\\").lower()
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        if d.FullName.replace("/", "\\").lower() == name_key:
            d.Activate()
            return d, "already_open"
    try:
        d = acad.Documents.Open(path, True)
        return d, "opened_readonly"
    except Exception as e1:
        try:
            d = acad.Documents.Open(path, False)
            return d, f"opened_rw ({e1})"
        except Exception as e2:
            raise RuntimeError(f"Open failed: {e1}; {e2}") from e2


def main():
    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    doc, how = open_readonly(acad, PATH)
    print("=== FILE ===", flush=True)
    print("path:", doc.FullName)
    print("open:", how)
    print("saved:", doc.Saved)

    layer_counts: dict[str, int] = defaultdict(int)
    layer_by_type: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    object_name_counts: dict[str, int] = defaultdict(int)
    proxy_aecc = 0
    aecc_names: set[str] = set()

    topo_feat_entities: list[tuple] = []
    feat_line_candidates: list[tuple] = []

    prop_d_hits: list[tuple] = []
    ramp_hits: list[tuple] = []
    xdata_samples: list[tuple] = []
    island_name_hits: list[tuple] = []

    ms = doc.ModelSpace
    n = ms.Count
    print("modelspace_count:", n, flush=True)

    for i in range(n):
        if i and i % 25000 == 0:
            print(f"  scan {i}/{n}...", flush=True)
        try:
            e = ms.Item(i)
        except Exception:
            continue
        try:
            lyr = e.Layer or ""
            oname = e.ObjectName or "?"
        except Exception:
            continue

        layer_counts[lyr] += 1
        layer_by_type[lyr][oname] += 1
        object_name_counts[oname] += 1

        if oname.startswith("Aecc") or "Aecc" in oname:
            aecc_names.add(oname)
        if oname == "AcDbProxyEntity":
            try:
                desc = e.ProxyMetafile.Type if hasattr(e, "ProxyMetafile") else ""
            except Exception:
                desc = ""
            if "Aecc" in str(desc) or "Feature" in str(desc):
                proxy_aecc += 1

        if LAYER_FILTER.search(lyr):
            pass  # counted via layer_counts filter below

        is_topo_layer = TOPO_LIKE.search(lyr) is not None or lyr.upper().startswith("C-TOPO")
        if is_topo_layer and oname in FEAT_TYPES | {"AeccDbFeatureLine"}:
            topo_feat_entities.append((oname, lyr))

        is_feat = oname in ("AeccDbFeatureLine",) or (
            oname in ("AcDb3dPolyline", "AcDbPolyline", "AcDb2dPolyline")
            and (is_topo_layer or "curb" in lyr.lower() or "feat" in lyr.lower())
        )
        if is_feat:
            try:
                h = e.Handle
                col = getattr(e, "Color", "?")
                L = approx_length(e)
                vc = vertex_count(e)
                feat_line_candidates.append((h, lyr, col, L, vc, oname))
            except Exception:
                pass

        ol = lyr.lower()
        onl = oname.lower()
        try:
            blk = getattr(e, "Name", "") or ""
        except Exception:
            blk = ""
        blob = f"{lyr} {oname} {blk}".lower()
        if "prop-d" in blob or "prop_d" in blob:
            prop_d_hits.append((e.Handle, lyr, oname, blk))
        if "ramp" in ol or "*u" in blk.lower() or "prop" in ol and "d" in ol:
            if "ramp" in ol or blk.startswith("*U") or "PROP" in lyr.upper():
                ramp_hits.append((e.Handle, lyr, oname, blk))

        if any(k in blob for k in ("island", "corner id", "corner_id")):
            island_name_hits.append((e.Handle, lyr, oname, blk))

    print("\n=== LAYERS (curb/steel/topo/feat/ramp/flare/dws/prop) ===", flush=True)
    matched_layers = sorted(
        [(lyr, layer_counts[lyr]) for lyr in layer_counts if LAYER_FILTER.search(lyr)],
        key=lambda x: (-x[1], x[0]),
    )
    for lyr, cnt in matched_layers:
        types = layer_by_type[lyr]
        top_types = ", ".join(f"{t}:{c}" for t, c in sorted(types.items(), key=lambda x: -x[1])[:4])
        print(f"  {cnt:5d}  {lyr}")
        print(f"         {top_types}")

    print("\n=== C-TOPO / FEAT linear objects ===")
    type_totals = defaultdict(int)
    for oname, lyr in topo_feat_entities:
        type_totals[oname] += 1
    for t in sorted(type_totals.keys()):
        print(f"  {t}: {type_totals[t]}")
    similar_layers = sorted({lyr for _, lyr in topo_feat_entities})
    print("  layers:", ", ".join(similar_layers[:15]), ("..." if len(similar_layers) > 15 else ""))

    print("\n=== STEEL CLASSIFY LAYERS ===")
    for sl in STEEL_LAYERS:
        c = layer_counts.get(sl, 0)
        print(f"  {'EXISTS' if c else 'MISSING':7s}  {c:5d}  {sl}")

    print("\n=== SAMPLE FEATURE LINES (up to 10) ===")
    # prefer Aecc then 3dpoly on C-TOPO-FEAT
    def sort_key(row):
        h, lyr, col, L, vc, oname = row
        pri = 0
        if oname == "AeccDbFeatureLine":
            pri = 0
        elif oname == "AcDb3dPolyline":
            pri = 1
        else:
            pri = 2
        if lyr.upper() == "C-TOPO-FEAT":
            pri -= 1
        return (pri, -(L or 0))

    samples = sorted(feat_line_candidates, key=sort_key)[:10]
    for row in samples:
        h, lyr, col, L, vc, oname = row
        Ls = f"{L:.1f}" if L is not None else "?"
        print(f"  {h}  {oname}  lyr={lyr}  color={col}  len~{Ls}ft  verts={vc}")

    for row in samples:
        if len(xdata_samples) >= 15:
            break
        try:
            e = doc.HandleToObject(row[0])
            xd = get_xdata_summary(e)
            if xd:
                xdata_samples.append((row[0], row[1], row[5], xd))
        except Exception:
            pass

    print("\n=== PROP-D / RAMP SOLIDS (sample) ===")
    print(f"  prop-d pattern hits: {len(prop_d_hits)}")
    for row in prop_d_hits[:8]:
        print(f"    {row}")
    print(f"  ramp/prop/U-block hits: {len(ramp_hits)}")
    for row in ramp_hits[:8]:
        print(f"    {row}")

    print("\n=== ISLAND ID (names / xdata) ===")
    print(f"  name/layer island hints: {len(island_name_hits)}")
    for row in island_name_hits[:8]:
        print(f"    {row}")
    print(f"  xdata samples: {len(xdata_samples)}")
    for row in xdata_samples[:10]:
        print(f"    {row}")

    print("\n=== CIVIL / Aecc OBJECT ENABLER ===")
    aecc_all = sorted(k for k in object_name_counts if "Aecc" in k or k.startswith("Aecc"))
    print(f"  Aecc ObjectName types in file: {len(aecc_all)}")
    for t in aecc_all[:20]:
        print(f"    {t}: {object_name_counts[t]}")
    print(f"  AcDbProxyEntity (possible missing enabler): {object_name_counts.get('AcDbProxyEntity', 0)}")
    print(f"  proxy_aecc_guess: {proxy_aecc}")
    if aecc_all and object_name_counts.get("AcDbProxyEntity", 0) == 0:
        print("  verdict: Aecc types resolve — Civil Object Enabler likely loaded.")
    elif object_name_counts.get("AcDbProxyEntity", 0) > 0 and not aecc_all:
        print("  verdict: Proxies only — install/load Civil Object Enabler for AeccDbFeatureLine access.")
    elif not aecc_all and not feat_line_candidates:
        print("  verdict: No feature lines found; check file content or enabler.")
    else:
        print("  verdict: Mixed — inspect proxy count vs Aecc* types above.")

    print("\n=== TOP ObjectName counts (all) ===")
    for t, c in sorted(object_name_counts.items(), key=lambda x: -x[1])[:12]:
        print(f"  {c:6d}  {t}")

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
