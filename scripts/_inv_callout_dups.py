"""Inventory callout GGs and any duplicate text near them."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference

GGs = list(range(6114, 6121))
KNOWN = {
    6114: (6107, 6117, 6138),
    6115: (6108, 6120, 6139),
    6116: (6109, 6123, 6140),
    6117: (6110, 6126, 6141),
    6118: (6111, 6129, 6142),
    6119: (6112, 6132, 6143),
    6120: (6113, 6135, 6144),
}

def text_of(e):
    try:
        if e.IsTextNodeElement:
            tn = e.AsTextNodeElement
            lines = []
            for i in range(1, int(tn.TextLinesCount) + 1):
                lines.append(tn.TextLine(i))
            return " | ".join(lines)
        if e.IsTextElement:
            return e.AsTextElement.Text
    except Exception:
        pass
    return ""

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
by_gg = {g: [] for g in GGs}
all_tn = []
while ee.MoveNext():
    e = ee.Current
    try:
        gg = int(e.GraphicGroup)
    except Exception:
        gg = 0
    info = {
        "id": int(e.ID64),
        "type": int(e.Type),
        "gg": gg,
        "level": e.Level.Name if e.Level else "?",
        "text": text_of(e)[:80],
    }
    if e.IsTextNodeElement:
        p = e.AsTextNodeElement.Origin
        info["xy"] = (round(float(p.X), 3), round(float(p.Y), 3))
        all_tn.append(info)
    if gg in by_gg:
        by_gg[gg].append(info)

print("=== GG inventory ===")
for g in GGs:
    ids = sorted(x["id"] for x in by_gg[g])
    print(f"GG {g}: {ids} expected {KNOWN.get(g)}")
    for x in sorted(by_gg[g], key=lambda z: z["id"]):
        print(f"  {x['id']} type={x['type']} {x['text']!r}")

# Find callout-like textnodes (APS / STREETLIGHT / VEHICULAR / CONTROL CABINET)
keys = ("APS", "STREETLIGHT", "VEHICULAR", "CONTROL CABINET", "PUSH BUTTON")
print("=== callout-like TextNodes ===")
callouts = []
for t in all_tn:
    if any(k in (t["text"] or "").upper() for k in keys):
        callouts.append(t)
        print(t)

# Cluster by rounded xy
from collections import defaultdict
clusters = defaultdict(list)
for t in callouts:
    clusters[t.get("xy")].append(t)
print("=== clusters (same xy) ===")
for xy, items in clusters.items():
    if len(items) > 1:
        print("DUP at", xy, "count", len(items))
        for it in items:
            print(" ", it)

print("total callout-like TN", len(callouts))
print("done")
