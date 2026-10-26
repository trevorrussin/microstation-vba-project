"""Draw the bottom (station) axis ticks and labels of a profile Drawing model by hand.

Why: for WZTC_TEST_ALG2 the Profile Grid annotation made by Create Drawing draws the elevation (y) axis only. The engineer's own CL1
Drawing model also has the station axis: DrawLineAnnotationCenter tick cells every 50 ft below the bottom axis (long at each 100 ft station,
short between) and a label ("66+00") under every 100 ft tick. Tried and ruled out as the cause: scale, Profile Shifts, round start station,
Named Boundaries route instead of the inline dialog, E_Stationing annotation of the alignment (annotate_profile status 0, nothing drawn).
These are plain lines and text, NOT ORD annotation cells; say so when delivering.

Geometry copied from CL1 (annotation scale 480) and scaled to the model's annotation scale S (CL1 value * S/480):
  long tick  = vertical line, 12 ft at S=1200, centred on the bottom axis line      (CL1: 4.8 ft)
  short tick = vertical line, 3 ft at S=1200, from the axis up into the grid         (CL1: 1.2 ft)
  label      = station text centred under the long tick, centre 14.5 ft below the axis at S=1200 (CL1: label box 3.3-8.3 ft below at 480)
Text size = NYSDOT standard text at plot ratio S: 1/16 in -> 6.252 ft high, 4.8 ft wide at 1200 (HDM Exhibit 22-27).
Levels: ticks on XAX_P, labels on XAXTXT_P_T (the levels of the existing y-axis ticks and labels); colour/weight ByLevel.
Station label = distance along the alignment + OFFSET (profile starts at 65+61.67 for distance 3280.83: OFFSET 3280.84).

ROOT CAUSE STATUS (2026-10-25): still unknown why Create Drawing leaves out the bottom axis ticks, station labels and the XGLV_P vertical grid
for these profiles (CL1, made from the NYSDOT sample, has all three). Compared by profile_boundary_data: every boundary setting equals CL1's
(Seed, VE 5, elevation spacing 5, station datum spacing 100, clearances) except shiftLocation (CL1 1 = Datum Stations, ours 0), and earlier
tests with Profile Shifts on did not change it. This script is the sanctioned fallback; `--check` reports whether it is needed/complete.

Usage (needs an open cad.drafting QA task on the scratch file; the model is activated in view 1):
  python scripts/ord_profile_xaxis.py "<Drawing model name>" [--scale 1200] [--offset 3280.84] [--dry] [--labels-only] [--skip N]
      [--grid | --grid-only] [--normalise-only] [--check] [--y0 <axis y>]     (--check is read-only: exit 1 when the grid/ticks/labels are missing)
Needs the VBA project to COMPILE: a compile error shows up as the first bridge op breaking into [break] and this script hanging
(2026-10-25: Modules/WZTCExec.bas ExecPlaceTextLabel used te.TextStyle.IsAnnotationText, not in this type library; now late bound).
Check first with: python scripts/vba_compile_recovery.py   (must say "likely clean").
Each placement is one bridge round trip (about 15 s); run it in the background for a whole sheet.
Known problems (2026-10-25): (1) y0 is the lowest XAX_P range, so a re-run after ticks exist measures y0 from the ticks (6 ft low): pass
--y0 <axis y> on any re-run. (2) FIXED 2026-10-25: text height/width are PAPER units in an annotation-scale model (0.00521 / 0.004; x1200 = 6.252 / 4.8 ft).
Labels must also be BOLD with justification 7 like CL1's own (2026-10-25: first versions were regular weight, justification 0).
Passing 6.252 made 7500 ft text; the MCP place_text_label tool has no height/width parameters, so it used the active style (0.01,
12 ft = twice standard). Correct label box at 1200 = 18.5 x 12.5 ft. Sheet 1 axis y 270, sheet 2 y 280 (pass --y0 on re-runs).
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "mcp-server"))
sys.path.insert(0, str(HERE))
import ms_connect  # noqa: E402


def station_text(label_ft: float) -> str:
    v = int(round(label_ft))
    return "%d+%02d" % (v // 100, v % 100)


def plan(x0: float, x1: float, y0: float, scale: float, offset: float) -> dict:
    k = scale / 480.0
    long_h, short_h, label_dy = 4.8 * k, 1.2 * k, 5.8 * k
    first = int(-(-(x0 + offset) // 50))                  # first 50 ft station at or after x0
    last = int((x1 + offset) // 50)
    ticks, labels = [], []
    for n in range(first, last + 1):
        st = n * 50.0
        x = st - offset
        if x < x0 - 1e-6 or x > x1 + 1e-6:
            continue
        if n % 2 == 0:                                      # multiple of 100 ft
            ticks.append([[x, y0 - long_h / 2], [x, y0 + long_h / 2]])
            labels.append((station_text(st), x, y0 - label_dy))
        else:
            ticks.append([[x, y0], [x, y0 + short_h]])
    # annotative text: height/width are PAPER units (x annotation scale on screen); 0.00521 x 1200 = 6.252 ft
    return {"ticks": ticks, "labels": labels, "text_h": 0.00521, "text_w": 0.004}


def grid_xs(p: dict, x0: float, x1: float) -> list:
    """x of every station tick that lies inside the frame (not on its edge lines)."""
    return sorted({round(t[0][0], 4) for t in p["ticks"] if x0 + 1.0 < t[0][0] < x1 - 1.0})


def normalise(app, name: str) -> None:
    # HDM 20.5.1: symbology ByLevel. The label op places colour 0 / weight 0, so set colour and weight ByLevel (-1) on my labels,
    # and on my axis ticks (short verticals on XAX_P; ORD's own frame lines are not touched).
    sc2 = app.CreateObjectInMicroStation("MicroStationDGN.ElementScanCriteria")
    sc2.ExcludeNonGraphical()
    en2 = app.ActiveDesignFile.Models.Item(name).Scan(sc2)
    fix = []
    while en2.MoveNext():
        e = en2.Current
        try:
            lv = e.Level.Name
            if e.Type == 17 and lv == "XAXTXT_P_T" and (e.LineWeight != -1 or e.Color != -1):
                fix.append(e)
            elif e.Type == 3 and lv == "XAX_P" and (e.Range.High.Y - e.Range.Low.Y) <= 13 and (e.Range.High.X - e.Range.Low.X) < 0.01 \
                    and (e.Color != -1 or e.LineWeight != -1):
                fix.append(e)
        except Exception:
            pass
    for e in fix:
        e.Color = -1
        e.LineWeight = -1
        e.Rewrite()
    print(f"set {len(fix)} of my axis ticks/labels to ByLevel colour and weight")


def check(app, name: str, p: dict, x0: float, x1: float) -> int:
    """Read-only: is the vertical grid / bottom-axis fallback complete in this Drawing model? Exit 1 when anything is missing."""
    import re
    sc3 = app.CreateObjectInMicroStation("MicroStationDGN.ElementScanCriteria")
    sc3.ExcludeNonGraphical()
    en3 = app.ActiveDesignFile.Models.Item(name).Scan(sc3)
    grid = ticks = labels = 0
    while en3.MoveNext():
        e = en3.Current
        try:
            lv = e.Level.Name
            if e.Type == 3 and lv == "XGLV_P":
                grid += 1
            elif e.Type == 3 and lv == "XAX_P" and (e.Range.High.Y - e.Range.Low.Y) <= 13 and (e.Range.High.X - e.Range.Low.X) < 0.01:
                ticks += 1
            elif e.Type == 17 and lv == "XAXTXT_P_T" and re.match(r"^\d+\+\d\d$", e.AsTextElement.Text.strip()):
                labels += 1
        except Exception:
            pass
    want_grid = len(grid_xs(p, x0, x1))
    rows = [("vertical grid XGLV_P", grid, want_grid), ("bottom axis ticks XAX_P", ticks, len(p["ticks"])), ("station labels XAXTXT_P_T", labels, len(p["labels"]))]
    bad = 0
    for label, have, want in rows:
        ok = have == want
        bad += 0 if ok else 1
        print(f"{'PASS' if ok else 'FAIL'} {label}: {have} found, {want} expected")
    return 1 if bad else 0


def main(argv: list[str]) -> int:
    name = argv[1]
    scale = float(argv[argv.index("--scale") + 1]) if "--scale" in argv else 1200.0
    offset = float(argv[argv.index("--offset") + 1]) if "--offset" in argv else 3280.84
    dry = "--dry" in argv
    app = ms_connect.get_microstation_app()
    model = app.ActiveDesignFile.Models.Item(name)
    sc = app.CreateObjectInMicroStation("MicroStationDGN.ElementScanCriteria")
    sc.ExcludeNonGraphical()
    en = model.Scan(sc)
    x0, x1, y0, y1 = 1e18, -1e18, 1e18, -1e18
    while en.MoveNext():
        e = en.Current
        try:
            if e.Level.Name != "XAX_P":
                continue
            r = e.Range
        except Exception:
            continue
        x0, x1, y0, y1 = min(x0, r.Low.X), max(x1, r.High.X), min(y0, r.Low.Y), max(y1, r.High.Y)
    if x0 > x1:
        raise RuntimeError("no XAX_P axis in " + name)
    if "--y0" in argv:
        y0 = float(argv[argv.index("--y0") + 1])
    p = plan(x0, x1, y0, scale, offset)
    print(f"{name}: axis x {x0:.2f}..{x1:.2f} y0 {y0:.2f}; {len(p['ticks'])} ticks, {len(p['labels'])} labels, text {p['text_h']:.3f} x {p['text_w']:.3f}")
    for t in p["labels"][:3]:
        print("  label", t)
    if dry:
        return 0
    if "--check" in argv:
        return check(app, name, p, x0, x1)
    import wztc_ops
    from bridge_client import bridge
    wztc_ops.set_bridge(bridge)
    app.ActiveDesignFile.Models.Item(name).Activate()
    time.sleep(2.0)
    if "--normalise-only" in argv:
        normalise(app, name)
        return 0
    if "--grid" in argv or "--grid-only" in argv:
        # Vertical grid at EVERY 50 ft station (the same x as the axis ticks), inside the frame. CL1's lines sit at frame-left + 50k, which
        # equals the stations only because CL1's frame starts on a round station; these profiles start at 65+61.67, so frame-left + 50k
        # put the grid 10 ft off the ticks (vision review caught it, 2026-10-25). ORD did not generate them for these profiles.
        gxs = grid_xs(p, x0, x1)
        if "--replace-grid" in argv:
            sc4 = app.CreateObjectInMicroStation("MicroStationDGN.ElementScanCriteria")
            sc4.ExcludeNonGraphical()
            en4 = app.ActiveDesignFile.Models.Item(name).Scan(sc4)
            old = []
            while en4.MoveNext():
                e4 = en4.Current
                try:
                    if e4.Type == 3 and e4.Level.Name == "XGLV_P":
                        old.append(e4.ID.Low)
                except Exception:
                    pass
            for oid in old:
                wztc_ops.delete_element(str(oid), own_element_only=False, reason="replace vertical grid at the station tick positions")
            print(f"deleted {len(old)} old XGLV_P lines")
        app.CadInputQueue.SendKeyin("ACTIVE LEVEL XGLV_P")
        time.sleep(0.8)
        for gx in gxs:
            wztc_ops.place_polyline([[gx, y0], [gx, y1]], reason="profile vertical grid line at a station (HDM ByLevel)")
        print(f"placed {len(gxs)} vertical grid lines on XGLV_P")
        if "--grid-only" in argv:
            return 0
    if "--labels-only" not in argv:
        app.CadInputQueue.SendKeyin("ACTIVE LEVEL XAX_P")
        time.sleep(0.8)
        # HDM 20.5.1: elements are placed ByLevel (active colour stays ByLevel; no ACTIVE COLOR key-in)
        for v in p["ticks"]:
            wztc_ops.place_polyline(v, reason="profile station axis tick (CL1 pattern)")
    app.CadInputQueue.SendKeyin("ACTIVE LEVEL XAXTXT_P_T")
    time.sleep(0.8)
    skip = int(argv[argv.index("--skip") + 1]) if "--skip" in argv else 0
    for text, x, y in p["labels"][skip:]:
        # CL1's own station labels: Arial, bold, justification 7 (checked with the style probe 2026-10-25)
        wztc_ops.place_text_label(text, x, y, height=p["text_h"], width=p["text_w"], bold=True, justification=7,
                                  reason="profile station axis label (CL1 pattern)")
    normalise(app, name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
