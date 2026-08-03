"""Live build check: call BUILD_WZTC_ORDER_TABLE through the bridge with a
sheet spec driving the stations/signs, then diff the returned rows against
what the spec itself says should and should not appear. Sheet-generic --
expected upstream/downstream labels and forbidden labels are derived from the
spec's orderTable section, not hardcoded per sheet (see Bridge/_live_order_table.py,
which this generalizes and supersedes for 619-311).

Requires MicroStation open with WZTCBridge polling.

Usage:
    python scripts/live_build_check.py 619-311 --speed 45 --lane-width 12 \
        --shoulder "12 ft" --area URBAN --road-type "Non-Freeway" \
        --category "Multilane Undivided"
"""
from __future__ import annotations

import argparse
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))
import sheet_spec  # noqa: E402
from bridge_client import Bridge  # noqa: E402


def expected_labels(spec: dict, resolved: dict) -> dict:
    """Derive what should and should not appear per alignment, straight from
    the spec -- no per-sheet literals. 'Present' labels for Sign rows must be
    the RESOLVED SignLibrary key (e.g. 'W04-02R', 'W20-01RF'), not the
    sheet's own bare signCode ('W4-2R') -- the bridge returns the resolved
    key, and 'W4-2R' is not even a substring of 'W04-02R' (different digit
    padding), so comparing against the bare code always fails. Caught live
    against 619-302 -- the first real MicroStation run this script did.
    'Forbidden' labels come from excludedRows, the direct antidote to
    WZTCRules.GetDefaultUpstreamItems emitting the same rows for every sheet."""
    signs = {s["signCode"]: s for s in spec["signs"]["items"]}
    out = {}
    for al in spec["orderTable"]["alignments"]:
        present = []
        for r in al["rows"]:
            if r["type"] == "Sign":
                present.append(sheet_spec.sign_library_key(signs[r["signCode"]], resolved))
            else:
                present.append(r.get("label"))
        forbidden = [x["label"] for x in al.get("excludedRows", [])]
        # Overlay zones are drawn but must NOT show up as a sequential station row.
        overlay_forbidden = []
        zones = {z["id"]: z for z in spec["corridor"]["zones"]}
        for o in al.get("overlayZones", []):
            overlay_forbidden.append(zones[o["zone"]]["sheetLabel"])
        out[al["alignIdx"]] = {
            "name": al["name"],
            "present": present,
            "forbidden": forbidden + overlay_forbidden,
            "expectedRowCount": len(al["rows"]),
        }
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet")
    ap.add_argument("--speed", type=int, required=True)
    ap.add_argument("--lane-width", type=int, required=True)
    ap.add_argument("--shoulder", required=True)
    ap.add_argument("--area", required=True)
    ap.add_argument("--closure", default=None)
    ap.add_argument("--exposure", default=None)
    ap.add_argument("--road-type", required=True)
    ap.add_argument("--category", required=True)
    args = ap.parse_args()

    spec = sheet_spec.load(args.sheet)
    if not spec:
        print(f"no spec for {args.sheet}")
        return 1

    resolved = sheet_spec.resolve(spec, args.speed, args.lane_width, args.shoulder,
                                   args.area, args.closure, args.exposure)
    payload = sheet_spec.order_table_rows(spec, resolved)
    exp = expected_labels(spec, resolved)

    non_sign = "|".join(payload["nonSignRows"])
    sign = "|".join(payload["signRows"])
    lane = resolved.get("laneTaper") or {}
    sh = resolved.get("shoulderTaper") or {}
    roll = resolved.get("rollAheadFt") or {}
    overrides = "|".join([
        f"bufferSpace={resolved.get('bufferFt', '')}",
        f"mergingTaper={lane.get('ft', '')}",
        f"shoulderTapers={sh.get('ft', '')}",
        f"rollAhead={roll.get('min', '')}",
        f"laneTaperSkips={lane.get('skipLines', '')}",
        f"shoulderTaperSkips={sh.get('skipLines', '')}",
        f"laneTaperDevices={lane.get('devices', '')}",
        f"shoulderTaperDevices={sh.get('devices', '')}",
    ])

    print(f"calling BUILD_WZTC_ORDER_TABLE for {args.sheet}...")
    print(f"  nonSign: {non_sign}")
    print(f"  signs:   {sign}")

    bridge = Bridge()
    resp = bridge.call(
        "BUILD_WZTC_ORDER_TABLE",
        category=args.category,
        sheetNum=args.sheet,
        speed=args.speed,
        roadType=args.road_type,
        laneWidth=args.lane_width,
        shoulderWidth=args.shoulder,
        signRowsTSV=sign,
        nonSignRowsTSV=non_sign,
        spacingOverridesTSV=overrides,
    )

    print("\nstatus:", resp.get("status") or resp.get("ok") or resp)
    rows = resp.get("rows") or []
    print(f"rows returned: {len(rows)}")

    labels = []
    for r in rows:
        if isinstance(r, dict):
            a = str(r.get("alignIdx", r.get("alignName", "")))
            lab = str(r.get("label") or r.get("signCode") or "").upper()
            labels.append((a, lab))
        elif isinstance(r, (list, tuple)) and len(r) >= 5:
            labels.append((str(r[0]), str(r[4]).upper()))

    print("\n=== checks (derived from the spec's own orderTable) ===")
    all_ok = True
    for align_idx, e in exp.items():
        these = [lab for a, lab in labels if a in (str(align_idx), e["name"])]
        print(f"\n{e['name']} (alignIdx {align_idx}): {these}")

        for want in e["present"]:
            ok = any(want.upper() in lab for lab in these)
            all_ok &= ok
            print(f"  [{'PASS' if ok else 'FAIL'}] expected present: {want}")
        for bad in e["forbidden"]:
            ok = not any(bad.upper() in lab for lab in these)
            all_ok &= ok
            print(f"  [{'PASS' if ok else 'FAIL'}] must NOT appear: {bad}")

    print(f"\n{'ALL CHECKS PASSED' if all_ok else 'SOME CHECKS FAILED'}")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
