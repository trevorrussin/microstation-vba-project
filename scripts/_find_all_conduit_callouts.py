"""Find all TRAFFIC SIGNAL CONDUIT callouts including cells."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())


def tn_lines(tn):
    return [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]


def walk(el, out, model_name, depth=0):
    try:
        if el.IsTextNodeElement:
            lines = tn_lines(el.AsTextNodeElement)
            t = "\n".join(lines)
            if "TRAFFIC SIGNAL CONDUIT" in t.upper() or (
                "TRAFFIC SIGNAL CABLE" in t.upper() and len(lines) >= 3
            ):
                rng = el.Range
                out.append(
                    {
                        "model": model_name,
                        "id": el.ID.Low,
                        "type": "TN",
                        "gg": getattr(el, "GraphicGroup", None),
                        "x": float(rng.Low.X),
                        "y": float(rng.Low.Y),
                        "lines": lines,
                        "depth": depth,
                    }
                )
        if int(el.Type) in (2, 7):  # cell / complex
            try:
                ee = el.AsComplexElement.GetSubElements()
                while ee.MoveNext():
                    walk(ee.Current, out, model_name, depth + 1)
            except Exception:
                pass
    except Exception:
        pass


hits = []
for m in app.ActiveDesignFile.Models:
    try:
        o = Dispatch("MicroStationDGN.ElementScanCriteria")
        o.ExcludeNonGraphical()
        ee = m.Scan(o)
    except Exception:
        continue
    while ee.MoveNext():
        walk(ee.Current, hits, m.Name)

# de-dupe by id
seen = set()
uniq = []
for h in hits:
    if h["id"] in seen:
        continue
    seen.add(h["id"])
    uniq.append(h)

print(f"found {len(uniq)}")
for h in sorted(uniq, key=lambda r: (-r["y"], r["x"])):
    print(
        f"\nid={h['id']} model={h['model']} gg={h['gg']} "
        f"xy=({h['x']:.1f},{h['y']:.1f}) lines={len(h['lines'])}"
    )
    for ln in h["lines"]:
        print(" ", repr(ln))
