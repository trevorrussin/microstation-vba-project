"""Compare agent-placed sign vs user reference assembly to the south."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import Dispatch

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import view_capture

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference

out = Path(r"c:\repos\microstation-vba-project\Bridge\_sign_compare.txt")
lines: list[str] = []


def scan():
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    return mr.Scan(oScan)


signs = []
ee = scan()
while ee.MoveNext():
    el = ee.Current
    if not el.IsCellElement:
        continue
    c = el.AsCellElement
    if c.Name not in ("W20-01RA", "TWZSGN_P"):
        continue
    ox, oy = float(c.Origin.X), float(c.Origin.Y)
    if ox < 1021800 or ox > 1022500:
        continue
    if oy < 217000 or oy > 217500:
        continue
    rng = el.Range
    signs.append(
        {
            "name": c.Name,
            "x": ox,
            "y": oy,
            "sx": float(c.Scale.X),
            "sy": float(c.Scale.Y),
            "w": float(rng.High.X - rng.Low.X),
            "h": float(rng.High.Y - rng.Low.Y),
            "color": el.Color,
            "wt": el.LineWeight,
            "lvl": el.Level.Name if el.Level else None,
        }
    )

signs.sort(key=lambda s: (-s["y"], s["x"], s["name"]))
lines.append("=== cells near corridor ===")
for s in signs:
    lines.append(
        f"{s['name']:10} xy=({s['x']:.2f},{s['y']:.2f}) "
        f"scale=({s['sx']:.1f},{s['sy']:.1f}) wh=({s['w']:.1f},{s['h']:.1f}) "
        f"c={s['color']} wt={s['wt']} lvl={s['lvl']}"
    )

# Cluster by y roughly
ys = sorted({round(s["y"], 0) for s in signs}, reverse=True)
lines.append(f"\n=== distinct Y clusters: {ys} ===")

# Wider search south of agent sign (agent at y=217353.75)
lines.append("\n=== all W20/TWZSGN with y in 217100..217360, x in 1021000..1022500 ===")
signs2 = []
ee = scan()
while ee.MoveNext():
    el = ee.Current
    if not el.IsCellElement:
        continue
    c = el.AsCellElement
    if c.Name not in ("W20-01RA", "TWZSGN_P"):
        continue
    ox, oy = float(c.Origin.X), float(c.Origin.Y)
    if not (1021000 <= ox <= 1022500 and 217100 <= oy <= 217360):
        continue
    rng = el.Range
    signs2.append((c.Name, ox, oy, float(c.Scale.X), float(rng.High.X - rng.Low.X),
                   float(rng.High.Y - rng.Low.Y), el.Color, el.LineWeight,
                   el.Level.Name if el.Level else None))
signs2.sort(key=lambda t: (-t[2], t[1], t[0]))
for t in signs2:
    lines.append(
        f"{t[0]:10} xy=({t[1]:.2f},{t[2]:.2f}) scaleX={t[3]:.1f} "
        f"wh=({t[4]:.1f},{t[5]:.1f}) c={t[6]} wt={t[7]} lvl={t[8]}"
    )

# Lines near those origins (post stems)
lines.append("\n=== lines near sign posts ===")
ee = scan()
while ee.MoveNext():
    el = ee.Current
    if not el.IsLineElement:
        continue
    try:
        s = el.AsLineElement.StartPoint
        e = el.AsLineElement.EndPoint
    except Exception:
        continue
    sx, sy = float(s.X), float(s.Y)
    ex, ey = float(e.X), float(e.Y)
    if not (1021000 <= sx <= 1022500 and 217100 <= sy <= 217360):
        continue
    length = ((ex - sx) ** 2 + (ey - sy) ** 2) ** 0.5
    if length < 5 or length > 100:
        continue
    lines.append(
        f"LINE ({sx:.2f},{sy:.2f})->({ex:.2f},{ey:.2f}) len={length:.1f} "
        f"c={el.Color} wt={el.LineWeight} lvl={el.Level.Name if el.Level else None}"
    )

# Texts
lines.append("\n=== text near corridor ===")
ee = scan()
while ee.MoveNext():
    el = ee.Current
    try:
        is_text = el.IsTextElement or el.IsTextNodeElement
    except Exception:
        is_text = False
    if not is_text:
        continue
    rng = el.Range
    cx = (float(rng.High.X) + float(rng.Low.X)) / 2
    cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
    if not (1021000 <= cx <= 1022500 and 217100 <= cy <= 217360):
        continue
    content = "?"
    try:
        if el.IsTextElement:
            content = el.AsTextElement.Text.replace("\n", " | ")
        else:
            content = "[textnode]"
    except Exception:
        pass
    lines.append(
        f"TEXT '{content[:60]}' xy=({cx:.2f},{cy:.2f}) "
        f"wh=({float(rng.High.X-rng.Low.X):.1f},{float(rng.High.Y-rng.Low.Y):.1f}) "
        f"c={el.Color} lvl={el.Level.Name if el.Level else None}"
    )

out.write_text("\n".join(lines), encoding="utf-8")
print(out.read_text(encoding="utf-8"))

# Screenshots: overview of both, then each cluster
# Agent at ~217353.75; look for southern cluster
agent_y = 217353.75
south = [s for s in signs2 if s[2] < agent_y - 5]
north = [s for s in signs2 if abs(s[2] - agent_y) < 5]
print("\nagent-ish count", len(north), "south count", len(south))

cap = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
# Wide view covering both
view_capture.navigate_view(1022130, 217280, 400, 250)
p1 = view_capture.capture_microstation(cap / "compare_signs_both.png")
print("saved", p1)

# Agent assembly closeup
view_capture.navigate_view(1022130.87, 217353.75, 180, 120)
p2 = view_capture.capture_microstation(cap / "compare_sign_agent.png")
print("saved", p2)

# Southern reference — use mean of southern posts if any
if south:
    sx = sum(s[1] for s in south) / len(south)
    sy = sum(s[2] for s in south) / len(south)
    print(f"south center ({sx:.2f},{sy:.2f})")
    view_capture.navigate_view(sx, sy, 180, 120)
    p3 = view_capture.capture_microstation(cap / "compare_sign_user_ref.png")
    print("saved", p3)
else:
    # maybe slightly south of agent on same tick
    view_capture.navigate_view(1022130.87, 217300, 220, 160)
    p3 = view_capture.capture_microstation(cap / "compare_sign_south_area.png")
    print("no south cluster found; saved area", p3)
