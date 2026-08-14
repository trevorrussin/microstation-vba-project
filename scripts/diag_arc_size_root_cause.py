"""Root-cause probe: why msdDimTypeArcSize draws huge/wrong-side arcs.

Hypothesis (from static read of WZTCExec.bas 2026-08-13):
ExecPlaceArcSizeDimension passes (center, tip1, tip2) in raw PATH order with
rot = Matrix3dIdentity and never normalizes sweep direction. The constructed
fallback ExecPlaceCurvedPlanDimension DOES normalize (da wrapped to (-pi,pi],
point order chosen by sign, plus a swap-retry when ArcElement.SweepAngle comes
back reflex) -- proof that arc constructors on this install take the long way
around when the points imply a clockwise sweep. That fix was never carried
back to the ArcSize path.

Each test compares the placed dimension's element range against the
analytically-expected bbox of the intended MINOR arc. A range far larger than
expected == the dimension swept the long way (or referenced a far-off center).

Run with MicroStation open on the WZTC design file. Places probe geometry in
an empty band near (90000, 287000); delete it afterwards yourself.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)

R = 200.0
DIM_GAP = 25.0


def arc_pts(cx: float, cy: float, r: float,
            a_start_deg: float, a_end_deg: float, step: int = 5):
    """Sample the arc from a_start to a_end going the SHORT way."""
    a1, a2 = math.radians(a_start_deg), math.radians(a_end_deg)
    da = (a2 - a1 + math.pi) % (2.0 * math.pi) - math.pi
    n = max(2, int(abs(math.degrees(da)) // step))
    return [(cx + r * math.cos(a1 + da * i / n),
             cy + r * math.sin(a1 + da * i / n)) for i in range(n + 1)]


def expected_bbox(cx: float, cy: float, r: float,
                  a_start_deg: float, a_end_deg: float):
    """Bbox of the intended minor arc PLUS its radial witness lines.

    The witness lines run from the tip arc (r) out to the dim arc (r+DIM_GAP),
    so they widen the range on shallow spans. Measuring the dim arc alone made
    a correct shallow placement look like an 8x blowup.
    """
    pts = arc_pts(cx, cy, r + DIM_GAP, a_start_deg, a_end_deg, step=1)
    tips = arc_pts(cx, cy, r, a_start_deg, a_end_deg, step=1)
    xs = [p[0] for p in pts] + [tips[0][0], tips[-1][0]]
    ys = [p[1] for p in pts] + [tips[0][1], tips[-1][1]]
    return (max(xs) - min(xs), max(ys) - min(ys))


def probe(label: str, cx: float, cy: float,
          a_start_deg: float, a_end_deg: float, swap: bool) -> dict:
    """Place one Arc Size dim; report range vs expected."""
    pts = arc_pts(cx, cy, R, a_start_deg, a_end_deg)
    p1, p2 = pts[0], pts[-1]
    if swap:
        p1, p2 = p2, p1
    a1, a2 = math.radians(a_start_deg), math.radians(a_end_deg)
    da = (a2 - a1 + math.pi) % (2.0 * math.pi) - math.pi
    amid = a1 + 0.5 * da
    ox = cx + (R + DIM_GAP) * math.cos(amid)
    oy = cy + (R + DIM_GAP) * math.sin(amid)

    exp_w, exp_h = expected_bbox(cx, cy, R, a_start_deg, a_end_deg)
    row: dict = {"test": label, "swap": swap,
                 "expectedW": round(exp_w, 1), "expectedH": round(exp_h, 1)}
    try:
        res = ops.place_arc_size_dimension(
            cx, cy, p1[0], p1[1], p2[0], p2[1], ox, oy,
            override_text="120'-0\"", reason=f"arcsize root-cause {label}")
    except Exception as e:
        row["result"] = f"ERROR {e}"
        return row
    eid = str(res.get("elementId") or "")
    row["elementId"] = eid
    if not eid:
        row["result"] = "no elementId"
        return row
    rng = ops.get_elements_range([eid])
    w = float(rng.get("width") or 0.0)
    h = float(rng.get("height") or 0.0)
    row["actualW"] = round(w, 1)
    row["actualH"] = round(h, 1)
    blowup = max(w / exp_w if exp_w else 0.0, h / exp_h if exp_h else 0.0)
    row["blowup"] = round(blowup, 2)
    if w < 1.0 and h < 1.0:
        row["result"] = "EMPTY (no geometry)"
    elif blowup > 1.8:
        row["result"] = "SWEPT LONG WAY / far-side arc"
    else:
        row["result"] = "hugs intended minor arc"
    return row


def main() -> int:
    rows: list[dict] = []

    # A: counter-clockwise 90 deg span (0 -> 90). Path order already CCW.
    rows.append(probe("A ccw path-order", 90000.0, 287000.0, 0.0, 90.0, False))
    rows.append(probe("A ccw swapped", 90600.0, 287000.0, 0.0, 90.0, True))

    # B: clockwise span (90 -> 0) == a real right-hand roadside bend.
    #    This is the case the shipped code never normalizes.
    rows.append(probe("B cw path-order", 91200.0, 287000.0, 90.0, 0.0, False))
    rows.append(probe("B cw swapped", 91800.0, 287000.0, 90.0, 0.0, True))

    # C: shallow bend near due-west travel (crosses the atan2 branch at pi).
    rows.append(probe("C west path-order", 92400.0, 287000.0, 170.0, 190.0, False))
    rows.append(probe("C west swapped", 93000.0, 287000.0, 170.0, 190.0, True))

    print()
    for r in rows:
        print(f"{r['test']:<20} swap={str(r['swap']):<5} "
              f"exp={r.get('expectedW')}x{r.get('expectedH')} "
              f"act={r.get('actualW')}x{r.get('actualH')} "
              f"blowup={r.get('blowup')}  -> {r.get('result')}")

    good = [r for r in rows if r.get("result") == "hugs intended minor arc"]
    print(f"\n{len(good)}/{len(rows)} placements hugged the intended arc.")
    if good:
        print("Arc Size CAN hug the roadside. Winning point orders:")
        for r in good:
            print(f"  - {r['test']} (swap={r['swap']})")
        print("=> fix ExecPlaceArcSizeDimension to normalize sweep like "
              "ExecPlaceCurvedPlanDimension, rather than keeping the "
              "constructed-graphics fallback.")
    else:
        print("No point order hugged the arc. Next lever: DimHeight semantics "
              "and the rot matrix (identity today) -- angular dims may measure "
              "sweep in the dimension frame, not world.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
