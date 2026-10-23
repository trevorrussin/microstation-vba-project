"""Guided flare fix via general acad_ask — prompt in AutoCAD when unsure.

NW reference solids are never deleted.
For each selected pocket: alert + 3 getpoints (TIP, LEFT, TAN) + place triangle.
"""
from __future__ import annotations

import math
import sys
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
from acad_ask import ask_choice, ask_message, ask_points

LAYER = "PROP-D-ROAD-RAMPS"
NW_CENTER = (1028470.0, 187478.0)

# Pockets we do NOT invent — engineer clicks blue-sketch corners
POCKETS = [
    {
        "id": "MidN-north",
        "label": "MidN NORTH (35.01 / over green arrow)",
        "zoom": (1028502.0, 187483.0, 14.0),
        "center": (1028500.0, 187487.0),
        "radius": 12.0,
        "hint": (
            "Fill your blue triangle:\n"
            "TIP = 35.01 where green topo meets tan\n"
            "LEFT = green topo / DWS side\n"
            "TAN = ON tan curb PAST the green slope arrow"
        ),
    },
    {
        "id": "MidN-south",
        "label": "MidN SOUTH",
        "zoom": (1028506.0, 187478.0, 12.0),
        "center": (1028507.0, 187479.0),
        "radius": 10.0,
        "hint": (
            "South MidN pocket:\n"
            "TIP = outer/tan junction at this end\n"
            "LEFT = along outer / DWS\n"
            "TAN = ON tan PRO-S past any slope arrow in the gap"
        ),
    },
    {
        "id": "MidS-north",
        "label": "MidS NORTH end",
        "zoom": (1028525.0, 187435.0, 14.0),
        "center": (1028526.0, 187435.0),
        "radius": 12.0,
        "hint": (
            "MidS north pocket between green topo, tan PRO-S, and depressed curb:\n"
            "TIP / LEFT / TAN like NW — right edge ON tan"
        ),
    },
    {
        "id": "MidS-south",
        "label": "MidS SOUTH end (diagonal tan)",
        "zoom": (1028524.0, 187420.0, 14.0),
        "center": (1028524.0, 187420.0),
        "radius": 12.0,
        "hint": (
            "MidS south pocket:\n"
            "TIP / LEFT / TAN — yellow contained to tan diagonal"
        ),
    },
    {
        "id": "E-north",
        "label": "East island NORTH end",
        "zoom": (1028585.0, 187451.0, 14.0),
        "center": (1028585.0, 187451.0),
        "radius": 12.0,
        "hint": (
            "East north pocket:\n"
            "TIP / LEFT / TAN — right edge ON tan PRO-S"
        ),
    },
    {
        "id": "E-south",
        "label": "East island SOUTH end",
        "zoom": (1028592.0, 187442.0, 14.0),
        "center": (1028592.0, 187442.0),
        "radius": 12.0,
        "hint": (
            "East south pocket:\n"
            "TIP / LEFT / TAN — right edge ON tan PRO-S"
        ),
    },
]


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def delete_near(doc, cx, cy, radius, keep_nw=True):
    ms = doc.ModelSpace
    handles = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER:
                continue
            mn, mx = e.GetBoundingBox()
            x = (float(mn[0]) + float(mx[0])) / 2
            y = (float(mn[1]) + float(mx[1])) / 2
            if keep_nw and math.hypot(x - NW_CENTER[0], y - NW_CENTER[1]) < 28:
                continue
            if math.hypot(x - cx, y - cy) <= radius:
                handles.append(e.Handle)
        except Exception:
            continue
    for h in handles:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(handles)


def place_tri(ms, tip, left, tan):
    sol = ms.AddSolid(vpt(*tip), vpt(*left), vpt(*tan), vpt(*tan))
    sol.Layer = LAYER
    sol.Color = 2
    return sol.Handle


def fix_pocket(doc, pocket: dict) -> str | None:
    ask_message(
        title=pocket["label"],
        message=pocket["hint"] + "\n\nClick OK, then pick TIP, LEFT, TAN.",
    )
    n = delete_near(doc, pocket["center"][0], pocket["center"][1], pocket["radius"])
    print(f"  deleted {n} old solids near {pocket['id']}")
    pts = ask_points(
        title=pocket["label"],
        message=pocket["hint"],
        prompts=[
            "1/3 TIP (apex / topo-tan junction)",
            "2/3 LEFT (green topo / DWS side)",
            "3/3 TAN (ON tan curb, past green arrow if present)",
        ],
        zoom=pocket["zoom"],
        banner=True,
    )
    tip, left, tan = pts
    h = place_tri(doc.ModelSpace, tip, left, tan)
    try:
        doc.Regen(1)
    except Exception:
        pass
    print(f"  placed {h} tip={tip} left={left} tan={tan}")
    return h


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    choice = ask_choice(
        title="Flare guided fix",
        message=(
            "NW flares stay untouched.\n"
            "I will NOT invent shapes — you click TIP / LEFT / TAN on tan "
            "like your blue sketch.\n\n"
            "Suggested options below are optional — you can type anything "
            "(e.g. 'only the 35.01 MidN north one', 'skip MidS', 'show me MidS first')."
        ),
        options=[
            "All non-NW pockets (6)",
            "MidN only (north+south)",
            "MidS only (north+south)",
            "East only (north+south)",
            "MidN NORTH only",
            "Stop",
        ],
        allow_free_text=True,
    )
    print("choice:", choice)
    low = choice.lower().strip()
    if low in ("stop", "cancel", "quit", "none", "6"):
        ask_message(title="Stopped", message="No flares changed.")
        return

    if choice in [p for p in [
        "All non-NW pockets (6)",
        "MidN only (north+south)",
        "MidS only (north+south)",
        "East only (north+south)",
        "MidN NORTH only",
    ]] or low.startswith("all"):
        if "midn north" in low or choice == "MidN NORTH only":
            selected = [p for p in POCKETS if p["id"] == "MidN-north"]
        elif choice.startswith("MidN only") or (low.startswith("midn") and "north only" not in low):
            selected = [p for p in POCKETS if p["id"].startswith("MidN")]
        elif choice.startswith("MidS") or low.startswith("mids"):
            selected = [p for p in POCKETS if p["id"].startswith("MidS")]
        elif choice.startswith("East") or low.startswith("east") or low == "e":
            selected = [p for p in POCKETS if p["id"].startswith("E")]
        elif low.startswith("all") or choice.startswith("All"):
            selected = POCKETS
        else:
            selected = None
    else:
        selected = None

    # Free-text path: map common phrases, else ask which pockets by free text again
    if selected is None:
        # Heuristic map
        if "35.01" in low or ("midn" in low and "north" in low):
            selected = [p for p in POCKETS if p["id"] == "MidN-north"]
        elif "midn" in low and "south" in low:
            selected = [p for p in POCKETS if p["id"] == "MidN-south"]
        elif "midn" in low:
            selected = [p for p in POCKETS if p["id"].startswith("MidN")]
        elif "mids" in low:
            selected = [p for p in POCKETS if p["id"].startswith("MidS")]
        elif "east" in low or low == "e":
            selected = [p for p in POCKETS if p["id"].startswith("E")]
        elif "all" in low:
            selected = POCKETS
        else:
            # Echo free text back — list pocket ids and ask them to name one/more
            pocket_list = "\n".join(f"  - {p['id']}: {p['label']}" for p in POCKETS)
            named = ask_text(
                title="Got your note",
                message=(
                    f"You wrote:\n\"{choice}\"\n\n"
                    f"Which pocket id(s) should we click now?\n{pocket_list}\n\n"
                    "Type ids separated by commas (e.g. MidN-north, MidS-south),\n"
                    "or 'all', or 'stop'."
                ),
                prompt="Pocket id(s)",
            )
            print("named:", named)
            nlow = named.lower().strip()
            if nlow in ("stop", "cancel", "quit", "none"):
                return
            if nlow == "all":
                selected = POCKETS
            else:
                ids = [x.strip() for x in named.replace(";", ",").split(",") if x.strip()]
                selected = []
                for i in ids:
                    hit = [p for p in POCKETS if p["id"].lower() == i.lower() or i.lower() in p["id"].lower()]
                    selected.extend(hit)
                # dedupe
                seen = set()
                selected = [p for p in selected if not (p["id"] in seen or seen.add(p["id"]))]
                if not selected:
                    ask_message(
                        title="No pockets matched",
                        message=f"Could not match: {named}\nTry again later.",
                    )
                    return

    results = []
    for pocket in selected:
        print("===", pocket["id"])
        try:
            h = fix_pocket(doc, pocket)
            results.append((pocket["id"], h))
        except Exception as ex:
            print("FAILED", pocket["id"], ex)
            again = ask_choice(
                title="Pocket failed / cancelled",
                message=f"{pocket['label']}\n{ex}\nContinue with remaining?",
                options=["Continue", "Stop"],
                allow_free_text=True,
            )
            if str(again).lower().startswith("stop"):
                break

    summary = "\n".join(f"{i}: {h}" for i, h in results) or "(none)"
    ask_message(title="Flare guided fix done", message=summary)
    print("DONE", results)


if __name__ == "__main__":
    main()
