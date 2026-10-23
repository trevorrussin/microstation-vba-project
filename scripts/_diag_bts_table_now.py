"""Diagnose estimate-table layout issues."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

DESC_X = 1030272.88
ITEM_X = 1030191.0
HEADER_BOTTOM = 239707.134


def main() -> None:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = None
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            mr = m
            break
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    bands = []
    descs = []
    nodes = []
    lines = []

    while ee.MoveNext():
        el = ee.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        cx = (float(rng.High.X) + float(rng.Low.X)) / 2
        if cy < 239545 or cy > 239725:
            continue
        if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
            continue
        h = float(rng.High.Y) - float(rng.Low.Y)
        w = float(rng.High.X) - float(rng.Low.X)
        ly, hy = float(rng.Low.Y), float(rng.High.Y)

        if lvl == "O_Details_Table Border_P":
            if int(el.Type) == 6 and w > 150 and 5 < h < 14:
                bands.append((ly, hy, cy, h))
            elif int(el.Type) == 3 and h < 1.0 and w > 50:
                lines.append(("hline", ly, hy, cy, w, lvl))
            continue

        if lvl != "TCSH_P_T":
            if int(el.Type) == 3 and abs(h) < 0.5 and w > 80:
                lines.append(("other_hline", ly, hy, cy, w, lvl))
            continue

        if el.IsTextElement and abs(cx - DESC_X) < 50:
            te = el.AsTextElement
            descs.append(
                {
                    "cy": cy,
                    "ly": ly,
                    "hy": hy,
                    "h": h,
                    "w": w,
                    "txt": te.Text.strip(),
                }
            )
        elif el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            t = " ".join(
                tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)
            ).strip()
            nodes.append(
                {
                    "cy": cy,
                    "ly": ly,
                    "hy": hy,
                    "h": h,
                    "cx": cx,
                    "txt": t,
                    "nlines": tn.TextLinesCount,
                }
            )

    bands = sorted(bands, key=lambda b: -b[2])
    descs = sorted(descs, key=lambda d: -d["cy"])
    nodes = sorted(nodes, key=lambda d: -d["cy"])

    print(f"bands={len(bands)} descs={len(descs)} nodes={len(nodes)} lines={len(lines)}")
    print("band heights:", sorted({round(b[3], 2) for b in bands}))
    print("band edges (top->bottom):")
    for ly, hy, cy, h in bands:
        print(f"  [{ly:.3f} .. {hy:.3f}] cy={cy:.3f} h={h:.2f}")

    print("\n2-line desc pairs:")
    for i in range(len(descs) - 1):
        a, b = descs[i], descs[i + 1]
        gap_c = a["cy"] - b["cy"]
        if gap_c >= 5:
            continue
        clear = a["ly"] - b["hy"]
        print(
            f"  pitch={gap_c:.2f} box_clear={clear:.2f} "
            f"h=({a['h']:.2f},{b['h']:.2f}) "
            f"{a['txt'][:36]!r} / {b['txt'][:36]!r}"
        )

    print("\ntext vs nearest band (pad):")
    for d in descs:
        band = min(bands, key=lambda b: abs(b[2] - d["cy"]))
        top = band[1] - d["hy"]
        bot = d["ly"] - band[0]
        # also check if any band EDGE crosses the text range
        crossed = [
            (bly, bhy)
            for bly, bhy, bcy, bh in bands
            if d["ly"] < bhy < d["hy"] or d["ly"] < bly < d["hy"]
        ]
        flag = []
        if top < 0.4 or bot < 0.4:
            flag.append("TIGHT_PAD")
        if crossed:
            flag.append(f"EDGE_IN_TEXT={crossed}")
        if flag or gap_near_two(d, descs):
            print(
                f"  cy={d['cy']:.3f} top={top:.2f} bot={bot:.2f} "
                f"bandh={band[3]:.2f} {' '.join(flag)} {d['txt'][:45]!r}"
            )

    print("\nitem/unit/qty nodes vs bands:")
    for n in nodes:
        if n["txt"] in (
            "ITEM",
            "DESCRIPTION",
            "UNIT",
            "QUANTITY",
            "ESTIMATE OF QUANTITIES",
        ):
            continue
        band = min(bands, key=lambda b: abs(b[2] - n["cy"]))
        top = band[1] - n["hy"]
        bot = n["ly"] - band[0]
        crossed = any(
            n["ly"] < bhy < n["hy"] or n["ly"] < bly < n["hy"]
            for bly, bhy, _, _ in bands
        )
        if top < 0.4 or bot < 0.4 or crossed or abs(band[2] - n["cy"]) > 1.0:
            print(
                f"  cy={n['cy']:.3f} top={top:.2f} bot={bot:.2f} "
                f"band_cy={band[2]:.3f} dcy={n['cy']-band[2]:+.3f} "
                f"crossed={crossed} {n['txt']!r}"
            )

    print("\nextra hlines:", lines[:30])


def gap_near_two(d, descs):
    for o in descs:
        if o is d:
            continue
        if abs(o["cy"] - d["cy"]) < 5:
            return True
    return False


if __name__ == "__main__":
    main()
