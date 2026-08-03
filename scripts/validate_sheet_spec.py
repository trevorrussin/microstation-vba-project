"""Validate a Data/sheet-specs/*.json sheet spec. Sheet-generic: reads table
roles, corridor zones and inputs from the spec itself rather than assuming
619-311's specific table numbers, so this is the same script for every sheet.

Three passes:

1. Structural — every zone / table / sign cross-reference resolves.
2. Transcription invariants — relationships that hold on any 619 WZTC lane/
   shoulder-taper sheet (skip line = 40 ft, devices = skips + 1, taper length
   monotonic in speed and lateral shift) and catch a mistyped table value.
   Sheets outside that family should still pass the structural pass; the
   invariant pass only runs the checks relevant to the tableRoles the sheet
   actually declares.
3. Resolution — resolve a worked case end to end via mcp-server/sheet_spec.py
   (not a re-implementation) and print the station table, proving the spec is
   executable and not just well-formed.

Usage:
    python scripts/validate_sheet_spec.py Data/sheet-specs/619-311.json
    python scripts/validate_sheet_spec.py Data/sheet-specs/619-311.json \
        --speed 55 --lane-width 11 --shoulder ">= 8 ft" --area RURAL
"""
from __future__ import annotations

import argparse
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))
import sheet_spec  # noqa: E402

DEFAULT_SKIP_LINE_FT = 40.0


class Problems:
    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def err(self, msg: str) -> None:
        self.errors.append(msg)

    def warn(self, msg: str) -> None:
        self.warnings.append(msg)

    @property
    def ok(self) -> bool:
        return not self.errors


# ------------------------------------------------------------------ structure

def check_reference_library(spec: dict, p: Problems) -> None:
    """Structural check for a sheet.kind == 'referenceLibrary' spec (e.g.
    619-011): no corridor/orderTable/signs, just tables + tableRoles + legend.
    Confirms every tableRoles entry resolves and every table has a non-empty
    rows list -- the analog of check_structure's cross-reference checks for a
    sheet that isn't a plan."""
    tables = spec.get("tables", {})
    for role, table_id in spec.get("tableRoles", {}).items():
        if role == "note":
            continue
        if table_id not in tables:
            p.err(f"tableRoles.{role} {table_id!r} not in tables")
    for tid, t in tables.items():
        if not t.get("rows"):
            p.err(f"table {tid}: no rows")
    if not spec.get("legend", {}).get("items"):
        p.warn("no legend.items present")


def check_structure(spec: dict, p: Problems) -> None:
    zones = {z["id"]: z for z in spec["corridor"]["zones"]}
    tables = spec["tables"]
    roles = spec.get("tableRoles", {})
    sign_codes = {s["signCode"] for s in spec["signs"]["items"]}

    size_codes: set[str] = set()
    size_table_id = roles.get("signSizes")
    if size_table_id:
        if size_table_id not in tables:
            p.err(f"tableRoles.signSizes {size_table_id!r} not in tables")
        else:
            size_codes = {r["signCode"] for r in tables[size_table_id]["rows"]}
    else:
        p.warn("tableRoles.signSizes not set — sign size cross-checks skipped")

    def zone_ref(zid, where):
        if zid is None:
            return
        # "bufferSpace..workArea" is a span shorthand for a device run
        for part in str(zid).split(".."):
            if part and part not in zones:
                p.err(f"{where}: unknown zone id {part!r}")

    orders = [z["order"] for z in spec["corridor"]["zones"]]
    if orders != sorted(orders) or len(set(orders)) != len(orders):
        p.err("corridor.zones: 'order' must be unique and ascending")

    for z in spec["corridor"]["zones"]:
        ls = z.get("lengthSource")
        if ls and "table" in ls:
            if ls["table"] not in tables:
                p.err(f"zone {z['id']}: lengthSource table {ls['table']!r} not in tables")
        if z["kind"] == "sign" and z.get("signCode") not in sign_codes:
            p.err(f"zone {z['id']}: signCode {z.get('signCode')!r} not in signs.items")
        if z["kind"] in ("gap", "taper", "buffer", "clearance") and not ls:
            p.err(f"zone {z['id']}: kind {z['kind']} requires a lengthSource")

    for al in spec["orderTable"]["alignments"]:
        nums = [r["rowNum"] for r in al["rows"]]
        if nums != list(range(1, len(nums) + 1)):
            p.err(f"alignment {al['alignIdx']}: rowNum must be 1..n in walk order")
        for r in al["rows"]:
            zone_ref(r.get("zone"), f"alignment {al['alignIdx']} row {r['rowNum']}")
            zone_ref(r.get("spacingZone"), f"alignment {al['alignIdx']} row {r['rowNum']}")
            if r["type"] == "Sign" and not r.get("spacingZone"):
                p.err(f"alignment {al['alignIdx']} row {r['rowNum']}: sign row needs a spacingZone")

    placed = {r.get("zone") for al in spec["orderTable"]["alignments"] for r in al["rows"]}
    placed |= {r.get("spacingZone") for al in spec["orderTable"]["alignments"] for r in al["rows"]}
    overlays = {o["zone"] for al in spec["orderTable"]["alignments"]
                for o in al.get("overlayZones", [])}
    for al in spec["orderTable"]["alignments"]:
        for o in al.get("overlayZones", []):
            zone_ref(o["zone"], f"alignment {al['alignIdx']} overlay")
            zone_ref(o["anchor"]["zone"], f"alignment {al['alignIdx']} overlay anchor")
            if zones.get(o["zone"], {}).get("consumesStation") is not False:
                p.err(f"overlay {o['zone']}: zone must be marked consumesStation=false")
    for z in spec["corridor"]["zones"]:
        if z.get("lengthSource") and z["id"] not in placed | overlays:
            p.warn(f"zone {z['id']} has a length but appears in no order-table walk")
        if z.get("consumesStation") is False and z["id"] in placed:
            p.err(f"zone {z['id']} is an overlay but also a sequential order-table row")

    for s in spec["signs"]["items"]:
        zone_ref(s.get("corridorZone"), f"sign {s['signCode']}")
        if size_table_id and s["signCode"] not in size_codes:
            p.err(f"sign {s['signCode']}: no size row in table {size_table_id}")
        sub = s.get("legendSubstitution")
        if sub:
            tbl = tables.get(sub["table"], {})
            cols = {c for r in tbl.get("rows", []) for c in r}
            if sub["column"] not in cols:
                p.err(f"sign {s['signCode']}: substitution column {sub['column']!r} "
                      f"not in table {sub['table']}")
        if s.get("postMounted") is False and s.get("mountedOn") is None:
            p.err(f"sign {s['signCode']}: not post mounted but no mountedOn given")

    if size_table_id:
        for code in size_codes:
            if code not in sign_codes:
                p.err(f"table {size_table_id} lists {code!r} with no entry in signs.items")

    for sym in spec["symbols"]["items"]:
        anchor = sym.get("stationAnchor")
        if anchor:
            zone_ref(anchor.get("zone"), f"symbol {sym['id']}")
        for run in sym.get("runs", []):
            zone_ref(run.get("zone"), f"symbol {sym['id']} run {run['id']}")

    for d in spec["annotations"]["dimensions"]:
        zone_ref(d["zone"], "annotations.dimensions")
        z = zones.get(d["zone"], {})
        if z and not z.get("dimensioned"):
            p.err(f"annotations dimensions {d['zone']}: zone is not marked dimensioned")
    dimmed = {d["zone"] for d in spec["annotations"]["dimensions"]}
    for z in spec["corridor"]["zones"]:
        if z.get("dimensioned") and z["id"] not in dimmed:
            p.err(f"zone {z['id']} is marked dimensioned but has no annotations entry")

    for r in spec["rules"]:
        for f in ("id", "severity", "source", "assert", "commonFailure"):
            if not r.get(f):
                p.err(f"rule {r.get('id')}: missing {f}")
        if r["severity"] not in ("error", "warning"):
            p.err(f"rule {r['id']}: bad severity {r['severity']!r}")
    ids = [r["id"] for r in spec["rules"]]
    if len(set(ids)) != len(ids):
        p.err("rules: duplicate rule id")


# ------------------------------------------------- transcription invariants

def check_tables(spec: dict, p: Problems) -> None:
    if "applicability" not in spec:
        # Reference-library sheets (sheet.kind == "referenceLibrary") have no
        # single set of inputs -- their tables are keyed by duration/roadType
        # combinations instead of one speed-range applicability block. The
        # per-cell arithmetic invariants (skip line = 40 ft, devices = skips+1)
        # still hold and are worth checking generically where a table's rows
        # carry laneTaper/shoulderTaper/longitudinalBufferSpace shapes, but the
        # speed-range/lane-width/shoulder-band cross-checks below assume a
        # single applicability block a plan sheet has and a library doesn't.
        roles = spec.get("tableRoles", {})
        tables = spec["tables"]
        skip_ft = DEFAULT_SKIP_LINE_FT

        def triplet(entry, where):
            ft, skips, devs = entry["ft"], entry["skipLines"], entry.get("devices")
            if abs(ft - skips * skip_ft) > 1e-6:
                p.err(f"{where}: {ft} ft is not {skips} x {skip_ft:.0f} ft skip lines")
            if devs is not None and devs != skips + 1:
                p.err(f"{where}: {devs} devices should be skipLines + 1 = {skips + 1}")

        taper_id = roles.get("taper") or roles.get("taperAndBuffer")
        if taper_id and taper_id in tables:
            for row in tables[taper_id]["rows"]:
                s = row.get("speedMph", row.get("speedBand"))
                for w, e in row.get("laneTaper", {}).items():
                    triplet(e, f"table {taper_id} speed {s} laneTaper {w} ft")
                for b, e in row.get("shoulderTaper", {}).items():
                    triplet(e, f"table {taper_id} speed {s} shoulderTaper {b}")
        roll_id = roles.get("rollAheadDistance")
        if roll_id and roll_id in tables:
            for row in tables[roll_id]["rows"]:
                for op in ("moving", "stationary"):
                    if op in row:
                        for k in ("min", "max"):
                            if k in row[op]:
                                triplet({**row[op][k], "devices": None},
                                        f"table {roll_id} {row.get('speedBand')} {op} {k}")
                if "min" in row and "max" in row:
                    triplet({**row["min"], "devices": None}, f"table {roll_id} {row.get('speedBand')} min")
                    triplet({**row["max"], "devices": None}, f"table {roll_id} {row.get('speedBand')} max")
        return

    roles = spec.get("tableRoles", {})
    tables = spec["tables"]
    skip_ft = spec["applicability"].get("skipLineFt", DEFAULT_SKIP_LINE_FT)

    def triplet(entry, where):
        ft, skips, devs = entry["ft"], entry["skipLines"], entry.get("devices")
        if abs(ft - skips * skip_ft) > 1e-6:
            p.err(f"{where}: {ft} ft is not {skips} x {skip_ft:.0f} ft skip lines")
        if devs is not None and devs != skips + 1:
            p.err(f"{where}: {devs} devices should be skipLines + 1 = {skips + 1}")

    taper_id = roles.get("taperAndBuffer")
    if taper_id and taper_id in tables:
        t02 = tables[taper_id]
        lane_widths = [str(w) for w in (spec["applicability"].get("laneWidthFt") or [])]
        bands = spec["applicability"].get("shoulderWidthBands", [])

        speeds = [r["speedMph"] for r in t02["rows"]]
        if speeds != sorted(speeds):
            p.err(f"table {taper_id}: rows must be in ascending speed order")
        rng = spec["applicability"]["speedRangeMph"]
        if "allowed" in rng:
            expected = list(rng["allowed"])
        else:
            expected = list(range(rng["min"], rng["max"] + 1, rng["increment"]))
        if speeds != expected:
            p.err(f"table {taper_id} speeds {speeds} disagree with applicability.speedRangeMph {expected}")

        for row in t02["rows"]:
            s = row["speedMph"]
            if lane_widths and "laneTaper" in row and set(row["laneTaper"]) != set(lane_widths):
                p.err(f"table {taper_id} speed {s}: laneTaper widths {sorted(row['laneTaper'])} "
                      f"!= {lane_widths}")
            if bands and "shoulderTaper" in row and set(row["shoulderTaper"]) != set(bands):
                p.err(f"table {taper_id} speed {s}: shoulderTaper bands {sorted(row['shoulderTaper'])} "
                      f"!= {bands}")
            for w, e in row.get("laneTaper", {}).items():
                triplet(e, f"table {taper_id} speed {s} laneTaper {w} ft")
            for b, e in row.get("shoulderTaper", {}).items():
                triplet(e, f"table {taper_id} speed {s} shoulderTaper {b}")
            buf = row.get("longitudinalBufferSpace")
            if buf and abs(buf["ft"] / skip_ft - buf["skipLines"]) > 0.75:
                p.err(f"table {taper_id} speed {s}: buffer {buf['ft']} ft vs "
                      f"{buf['skipLines']} skip lines is off by more than rounding")

        # Taper length must not shrink as speed rises or as the shift widens.
        for w in lane_widths:
            vals = [r["laneTaper"][w]["ft"] for r in t02["rows"] if "laneTaper" in r]
            if vals and vals != sorted(vals):
                p.err(f"table {taper_id}: laneTaper at {w} ft is not monotonic in speed: {vals}")
        for row in t02["rows"]:
            if "laneTaper" in row and lane_widths:
                vals = [row["laneTaper"][w]["ft"] for w in lane_widths]
                if vals != sorted(vals):
                    p.err(f"table {taper_id} speed {row['speedMph']}: laneTaper is not monotonic "
                          f"in lane width: {vals}")
            if "shoulderTaper" in row and bands:
                vals = [row["shoulderTaper"][b]["ft"] for b in bands]
                if vals != sorted(vals):
                    p.err(f"table {taper_id} speed {row['speedMph']}: shoulderTaper is not monotonic "
                          f"in shoulder width: {vals}")
    else:
        p.warn("tableRoles.taperAndBuffer not set — taper/buffer invariants skipped")

    roll_id = roles.get("rollAheadDistance")
    if roll_id and roll_id in tables:
        for row in tables[roll_id]["rows"]:
            label = row.get("speedBand") or row.get("gvwBand") or "?"
            for k in ("min", "max"):
                if k in row:
                    triplet({**row[k], "devices": None}, f"table {roll_id} {label} {k}")
            # Equal min==max is valid (sheet prints a single value for both
            # GVW columns on some mowing/mulching bands, e.g. Family 9 <=40).
            if "min" in row and "max" in row and row["min"]["ft"] > row["max"]["ft"]:
                p.err(f"table {roll_id} {label}: min is above max")
    else:
        p.warn("tableRoles.rollAheadDistance not set — roll-ahead invariants skipped")

    spacing_id = roles.get("advanceWarningSpacing")
    if spacing_id and spacing_id in tables:
        for row in tables[spacing_id]["rows"]:
            if "A" in row and "B" in row and "C" in row and not (row["A"] == row["B"] == row["C"]):
                p.warn(f"table {spacing_id} {row.get('areaType')} {row.get('speedBand')}: A/B/C differ "
                       f"({row['A']}/{row['B']}/{row['C']}) — verify against the sheet")

    size_id = roles.get("signSizes")
    if size_id and size_id in tables:
        classes = {i["id"]: i for i in spec["inputs"]}.get("signSizeClass", {}).get("allowed", [])
        for row in tables[size_id]["rows"]:
            for c in classes:
                if c not in row:
                    p.err(f"table {size_id} {row['signCode']}: missing size for {c}")


# ----------------------------------------------------------------- resolution

def default_case(spec: dict, overrides: dict) -> dict:
    """Fill in a worked case from the spec's own declared inputs where the
    caller didn't specify one, defaulting to the input's declared 'default'
    or its first allowed value. Printed so a reviewer can see exactly which
    case was resolved."""
    by_id = {i["id"]: i for i in spec["inputs"]}
    case = dict(overrides)

    def fill(key, *input_ids):
        if case.get(key) is not None:
            return
        for input_id in input_ids:
            inp = by_id.get(input_id)
            if inp:
                case[key] = inp.get("default") or (inp["allowed"][0] if inp.get("allowed") else None)
                return

    fill("speed", "preconstructionPostedSpeedMph")
    fill("laneWidth", "laneWidthFt")
    fill("shoulder", "shoulderWidthBand")
    # Sheet's own name for the advance-warning-spacing key input varies:
    # 619-311 calls it 'areaType' (URBAN/RURAL); 619-302 calls it
    # 'roadTypeForSignSpacing' (URBAN/RURAL/FREEWAY, same table role).
    fill("area", "areaType", "roadTypeForSignSpacing")
    fill("closure", "closureType")
    fill("exposure", "exposureCondition")
    return case


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("spec", type=pathlib.Path)
    ap.add_argument("--speed", type=int, default=None)
    ap.add_argument("--lane-width", type=int, default=None)
    ap.add_argument("--shoulder", default=None)
    ap.add_argument("--area", default=None)
    ap.add_argument("--closure", default=None)
    ap.add_argument("--exposure", default=None)
    args = ap.parse_args()

    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    p = Problems()
    is_ref_library = spec.get("sheet", {}).get("kind") == "referenceLibrary"

    if is_ref_library:
        check_reference_library(spec, p)
    else:
        check_structure(spec, p)
    check_tables(spec, p)

    print(f"{spec['sheet']['number']}  {spec['sheet']['title']}")
    if is_ref_library:
        print(f"schema {spec['schemaVersion']}   kind=referenceLibrary   "
              f"{len(spec['tables'])} tables   "
              f"{len(spec.get('legend', {}).get('items', []))} legend items")
    else:
        print(f"schema {spec['schemaVersion']}   "
              f"{len(spec['tables'])} tables   "
              f"{len(spec['corridor']['zones'])} zones   "
              f"{len(spec['signs']['items'])} signs   "
              f"{len(spec['rules'])} rules")

    for w in p.warnings:
        print(f"  WARN  {w}")
    for e in p.errors:
        print(f"  ERROR {e}")
    if not p.ok:
        print(f"\nFAILED — {len(p.errors)} error(s)")
        return 1
    print(f"  structure + transcription OK ({len(p.warnings)} warning(s))")

    if is_ref_library:
        print("\n(reference-library sheet -- no corridor/order-table to resolve, stopping here)")
        return 0

    case = default_case(spec, {
        "speed": args.speed, "laneWidth": args.lane_width, "shoulder": args.shoulder,
        "area": args.area, "closure": args.closure, "exposure": args.exposure,
    })
    needs_area = bool(spec.get("tableRoles", {}).get("advanceWarningSpacing"))
    needs_lane = bool(spec.get("applicability", {}).get("laneWidthFt"))
    required = ["speed", "shoulder"]
    if needs_lane:
        required.append("laneWidth")
    if needs_area:
        required.append("area")
    missing = [k for k in required if case.get(k) is None]
    if missing:
        print(f"\nCannot resolve a worked case — spec has no default/allowed value for: {missing}")
        return 1
    # Shoulder-only sheets still pass a dummy lane width through to resolve().
    if case.get("laneWidth") is None:
        case["laneWidth"] = 12

    try:
        res = sheet_spec.resolve(spec, case["speed"], case["laneWidth"], case["shoulder"],
                                  case.get("area"), case.get("closure"), case.get("exposure"))
    except sheet_spec.SpecError as e:
        print(f"\nFAILED resolving worked case {case}: {e}")
        return 1

    print(f"\nWorked case: {case['speed']} mph, "
          + (f"{case['laneWidth']} ft lane, " if needs_lane else "")
          + f"{case['shoulder']} shoulder"
          + (f", {case['area']}" if needs_area else "")
          + (f", closure={case['closure']!r}, exposure={case['exposure']!r}"
             if case.get("closure") else ""))
    for k, v in res.items():
        print(f"  {k:<20} {v}")

    walk = sheet_spec.station_walk(spec, res)
    print("\n  {:<11} {:<8} {:<28} {:>10} {:>12}".format(
        "ALIGNMENT", "ROW", "ITEM", "LENGTH FT", "STATION FT"))
    for r in walk:
        row_num = r["rowNum"] if r["rowNum"] is not None else r.get("note", "overlay")
        print("  {:<11} {:<8} {:<28} {:>10} {:>12}".format(
            r["alignName"], str(row_num), str(r["item"]), f"{r['lengthFt']:g}",
            f"{r['stationFt']:g}"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
