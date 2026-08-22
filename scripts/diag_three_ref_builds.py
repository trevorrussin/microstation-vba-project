"""Compare G20 stem gap + dim footprint on the three known 619-311 builds."""
from __future__ import annotations

import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(r"c:\repos\microstation-vba-project\mcp-server")))

from bridge_client import Bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(Bridge())


def fnum(r, *keys):
    for k in keys:
        if k in r and r[k] not in (None, ""):
            try:
                return float(r[k])
            except Exception:
                pass
    return None


def _row_by_id(element_id: str, near_x: float, near_y: float) -> dict:
    for tf in ("CELL", "DIMENSION", "LINE", "TEXT", ""):
        rows = ops.find_elements_near(near_x, near_y, 900.0, type_filter=tf, force=True)
        for r in rows:
            if str(r.get("elementId")) == str(element_id):
                return r
    raise KeyError(element_id)


def analyze(label: str, g20_id: str, post_id: str, near: tuple[float, float],
            dim_ids: list[str]) -> None:
    print(f"=== {label} ===", flush=True)
    nx, ny = near
    g20 = _row_by_id(g20_id, nx, ny)
    post = _row_by_id(post_id, nx, ny)

    px, py = fnum(post, "cx"), fnum(post, "cy")
    gx, gy = fnum(g20, "cx"), fnum(g20, "cy")
    print(f"POST ({px:.1f},{py:.1f})  G20 ({gx:.1f},{gy:.1f})", flush=True)

    dx, dy = gx - px, gy - py
    n = math.hypot(dx, dy) or 1.0
    ux, uy = dx / n, dy / n

    def edge_t(r, want_max: bool) -> float:
        corners = [
            (fnum(r, "rangeLowX"), fnum(r, "rangeLowY")),
            (fnum(r, "rangeLowX"), fnum(r, "rangeHighY")),
            (fnum(r, "rangeHighX"), fnum(r, "rangeLowY")),
            (fnum(r, "rangeHighX"), fnum(r, "rangeHighY")),
        ]
        ts = [(x - px) * ux + (y - py) * uy for x, y in corners]
        return max(ts) if want_max else min(ts)

    t_post_out = edge_t(post, True)
    t_g20_in = edge_t(g20, False)
    gap = t_g20_in - t_post_out
    print(
        f"bbox_gap_along_stem_dir={gap:.2f} ft "
        f"(STEM_GAP target ~50; >55 means visible white gap; <45 means overlap)",
        flush=True,
    )

    mx, my = 0.5 * (px + gx), 0.5 * (py + gy)
    lines = ops.find_elements_near(mx, my, 90.0, type_filter="LINE", force=True)
    stems = []
    for ln in lines:
        x1 = fnum(ln, "rangeLowX")
        y1 = fnum(ln, "rangeLowY")
        x2 = fnum(ln, "rangeHighX")
        y2 = fnum(ln, "rangeHighY")
        if None in (x1, y1, x2, y2):
            continue
        L = math.hypot(x2 - x1, y2 - y1)
        if 40 < L < 120:
            t_a = (x1 - px) * ux + (y1 - py) * uy
            t_b = (x2 - px) * ux + (y2 - py) * uy
            t_lo, t_hi = min(t_a, t_b), max(t_a, t_b)
            gap_post = t_lo - t_post_out
            gap_face = t_g20_in - t_hi
            stems.append(
                {
                    "id": ln.get("elementId"),
                    "L": round(L, 2),
                    "gap_post": round(gap_post, 2),
                    "gap_face": round(gap_face, 2),
                    "t": (round(t_lo, 2), round(t_hi, 2)),
                }
            )
    stems.sort(key=lambda s: abs(s["gap_face"]) + abs(s["gap_post"]))
    print("best_stems", stems[:4], flush=True)

    for did in dim_ids:
        try:
            d = _row_by_id(did, nx, ny)
        except KeyError:
            print(f"dim {did} NOT FOUND", flush=True)
            continue
        w = abs(fnum(d, "rangeHighX") - fnum(d, "rangeLowX"))
        h = abs(fnum(d, "rangeHighY") - fnum(d, "rangeLowY"))
        print(
            f"dim {did} footprint={w:.1f}x{h:.1f} "
            f"at ({fnum(d, 'cx'):.0f},{fnum(d, 'cy'):.0f})",
            flush=True,
        )


def main() -> int:
    analyze("reference_L", "215248", "215245", (92690.0, 300080.0),
            ["215324", "215313", "215315"])
    analyze("earlier_proof", "224955", "224952", (111020.0, 300100.0),
            ["225031", "225020", "225022"])
    analyze("speed_fix", "228026", "228023", (183020.0, 340100.0),
            ["228102", "228091", "228093"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
