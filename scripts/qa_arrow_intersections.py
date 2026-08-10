"""QA captures + list striping arrow cells."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
OUT = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
LIB = r"c:\pwworking\usny\d0119091\ny_plan_striping.cel"

wztc_ops.attach_cell_library(LIB)
rows = wztc_ops.list_cells(name_contains="SA")
print("SA cells", len(rows) if isinstance(rows, list) else rows)
for c in rows or []:
    print(f"{c.get('name') or c.get('cellName')}\t{c.get('description')}")

for q in ("only", "left straight", "straight right", "left right", "three"):
    found = wztc_ops.find_cell(q, library_path=LIB)
    print("find", q, found if not isinstance(found, list) else f"{len(found)} hits")
    if isinstance(found, list):
        for c in found[:12]:
            print(f"  {c.get('cellName')}\t{c.get('description')}\t{c.get('libraryPath')}")

for name, cx, cy, w in [
    ("qa_cont", 24650, 291200, 280),
    ("qa_ded", 25150, 291200, 300),
    ("qa_old", 23800, 292500, 280),
    ("qa_cont_west", 24570, 291200, 110),
    ("qa_cont_south", 24650, 291120, 110),
    ("qa_cont_north", 24650, 291280, 110),
    ("qa_cont_east", 24730, 291200, 110),
    ("qa_ded_west", 25050, 291200, 140),
]:
    wztc_ops.adjust_view(center_x=cx, center_y=cy, width=w, height=w)
    p = Path(wztc_ops.capture_view()["path"])
    (OUT / f"{name}.png").write_bytes(p.read_bytes())
    print("saved", name)
