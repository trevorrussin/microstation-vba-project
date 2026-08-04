"""Backfill Family 6 verbatim notes.printed from PDF text layers.

Implements Claude's bounded follow-up: restore printed note prose so engineers
can read sheet language. Placement rules[] were already populated.

Strategy:
- rotation=0 pages: vertical note columns (shared x) — high fidelity.
- rotation!=0 pages: column strips under the NOTES: number row on the original
  page (assign by word left-edge into voronoi bounds). Mark confidence
  drawing when residual column-bleed remains; still better than empty.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]

SHEETS = [
    "090", "091", "307", "308", "309", "314", "321", "322", "323", "324",
    "407", "421", "422", "519", "524",
]

MARKER_RE = re.compile(r"^(\d+|N\d+)\.$")
STOP_BLEED = (
    "* PRECONSTRUCTION",
    " PRECONSTRUCTION POSTED",
    " ROAD TYPE",
    " WARNING TYPE",
    " TABLE 0",
)


def norm(s: str) -> str:
    s = s.replace("\u2022", "½").replace("�", "½")
    return re.sub(r"\s+", " ", s).strip()


def cut_bleed(text: str) -> str:
    for stop in STOP_BLEED:
        if stop in text:
            text = text.split(stop)[0]
    return text.strip(" .,")


def markers(words):
    out = []
    for w in words:
        m = MARKER_RE.match(w[4])
        if m:
            out.append({
                "num": m.group(1),
                "x0": w[0], "y0": w[1], "x1": w[2], "y1": w[3],
            })
    return out


def cluster(values, tol: float):
    if not values:
        return []
    values = sorted(values, key=lambda v: v[0])
    clusters = [[values[0]]]
    for v in values[1:]:
        if abs(v[0] - clusters[-1][-1][0]) <= tol:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return clusters


def extract_vertical(words, marks) -> list[str]:
    marks = sorted(marks, key=lambda m: m["y0"])
    x_lo = min(m["x0"] for m in marks) - 5
    x_hi = max(m["x1"] for m in marks) + 340
    col = [w for w in words if x_lo <= w[0] <= x_hi]
    notes = []
    for i, m in enumerate(marks):
        y0 = m["y0"] - 3
        y1 = marks[i + 1]["y0"] - 2 if i + 1 < len(marks) else m["y0"] + 260
        chunk = sorted(
            [w for w in col if y0 <= w[1] < y1],
            key=lambda w: (round(w[1] / 3), w[0]),
        )
        text = cut_bleed(norm(" ".join(w[4] for w in chunk)))
        text = re.sub(rf"^{re.escape(m['num'])}\.?\s*", "", text)
        text = re.sub(r"\bNOTES:?\b", "", text).strip(" :")
        if len(text) < 20:
            continue
        notes.append(f"{m['num']}. {norm(text)}")
    return notes


def extract_rotated(words, title, marks) -> list[str]:
    marks = sorted(marks, key=lambda m: m["x0"])
    xs = [m["x0"] for m in marks] + [title[0]]
    bounds = []
    for i, m in enumerate(marks):
        left = (xs[i - 1] + xs[i]) / 2 if i else xs[i] - 18
        right = (xs[i] + xs[i + 1]) / 2
        bounds.append((m["num"], left, right))
    y0 = title[1] + 4
    y1 = y0 + 420
    cols = {n: [] for n, _, _ in bounds}
    for w in words:
        if not (y0 <= w[1] <= y1):
            continue
        if MARKER_RE.match(w[4]) or w[4] in ("NOTES:", "NOTES"):
            continue
        x = w[0]
        for n, a, b in bounds:
            if a <= x < b:
                cols[n].append(w)
                break
    notes = []
    for n, _, _ in sorted(
        bounds,
        key=lambda t: (0, int(t[0])) if t[0].isdigit() else (1, int(t[0][1:] or 0)),
    ):
        ws = sorted(cols[n], key=lambda w: (w[1], w[0]))
        text = cut_bleed(norm(" ".join(w[4] for w in ws)))
        if len(text) < 15:
            continue
        notes.append(f"{n}. {text}")
    return notes


def prose_score(notes: list[str]) -> int:
    keys = (
        "SHORT-TERM", "INTERMEDIATE", "LONG-TERM", "FLAGGER", "DURATION",
        "CHANNELIZING", "SHALL", "WORK THAT", "SIDEWALK", "PEDESTRIAN",
    )
    return sum(1 for n in notes if any(k in n.upper() for k in keys))


def page_candidates(page) -> list[tuple[str, list[str]]]:
    words = page.get_text("words")
    marks = markers(words)
    out: list[tuple[str, list[str]]] = []

    # Rotated number-row under NOTES:
    for t in words:
        if t[4] not in ("NOTES:", "NOTES"):
            continue
        near = [m for m in marks if abs(m["y0"] - t[1]) < 12 and m["num"].isdigit()]
        if len(near) >= 3:
            notes = extract_rotated(words, t, near)
            if notes:
                out.append(("rot", notes))

    # Upright columns
    for cl in cluster([(m["x0"], m) for m in marks], tol=28):
        ms = [m for _, m in cl]
        if len(ms) < 3:
            continue
        notes = extract_vertical(words, ms)
        if notes:
            out.append(("vert", notes))
    return out


def pick_best(cands: list[tuple[str, list[str]]]) -> tuple[str, list[str]] | None:
    if not cands:
        return None

    def key(item):
        kind, notes = item
        bleed = sum(
            1 for n in notes
            if "NOTES:" in n
            or (n.count("/") >= 3 and "SEE" not in n.upper())
            or re.search(r"\b(PVH|PVL|TMIA)\b", n) and "PROTECTIVE VEHICLE" not in n[:50]
        )
        # Prefer vertical when scores close; prefer plan-prose over table footnotes
        return (
            prose_score(notes) - bleed * 2,
            len(notes),
            1 if kind == "vert" else 0,
            sum(len(n) for n in notes),
        )

    cands = sorted(cands, key=key, reverse=True)
    return cands[0]


def extract_sheet(sheet: str) -> tuple[list[str], str, str]:
    """Returns (notes, geometry_kind, confidence_hint_reason)."""
    doc = fitz.open(str(ROOT / "Bridge" / "captures" / f"619-{sheet}.pdf"))
    all_cands: list[tuple[str, list[str]]] = []
    for page in doc:
        all_cands.extend(page_candidates(page))
    if not all_cands:
        return [], "none", "no candidates"

    def rank(item):
        kind, notes = item
        bleed = sum(
            1 for n in notes
            if "NOTES:" in n
            or re.match(r"^\d+\.\s+(\d+\s+){4,}", n)
            or (n.count("/") >= 3 and "SEE" not in n.upper())
        )
        return (
            prose_score(notes) - bleed * 3,
            len(notes),
            1 if kind == "vert" else 0,
            sum(len(n) for n in notes),
        )

    ranked = sorted(all_cands, key=rank, reverse=True)
    for kind, notes in ranked:
        seen = set()
        merged = []
        for n in notes:
            head = n.split(".", 1)[0]
            if head in seen:
                continue
            seen.add(head)
            merged.append(n)

        def sk(s: str):
            h = s.split(".", 1)[0]
            if h.startswith("N"):
                return (1, int(h[1:] or 0))
            return (0, int(h))

        merged.sort(key=sk)
        ok, conf, why = quality(merged)
        if ok:
            return merged, kind, f"{conf}:{why}"
    # Return best even if weak — caller decides
    kind, notes = ranked[0]
    return notes, kind, "fail:all-candidates-weak"


def quality(notes: list[str]) -> tuple[bool, str, str]:
    """Returns (ok, confidence, reason)."""
    if len(notes) < 2:
        return False, "verbatim", f"only {len(notes)}"
    avg = sum(len(n) for n in notes) / len(notes)
    if avg < 30:
        return False, "verbatim", f"avg {avg:.0f}"
    if any("NOTES:" in n for n in notes):
        return False, "verbatim", "contains NOTES: bleed"
    # Table-grid bleed: note opens with a run of bare integers
    grid = sum(1 for n in notes if re.match(r"^\d+\.\s+(\d+\s+){4,}", n))
    if grid:
        return False, "verbatim", f"table-grid bleed on {grid} notes"
    interleaved = sum(
        1 for n in notes
        if re.search(r"\b[A-Z]{4,}(?:\s+[A-Z]{2,}){0,2}\s+[a-z]{2,5}\s+[A-Z]{4,}", n)
        or re.search(r"\b(SINGLE|DAYLIGHT)\s+(SHORT-TERM|PERIOD|STATIONARY)\b", n)
        and "SHORT-TERM STATIONARY IS DAYTIME" not in n
    )
    # Classic rot=270 interleave: "SINGLE SHORT-TERM DAYLIGHT STATIONARY PERIOD. IS DAYTIME"
    rot_garble = sum(
        1 for n in notes
        if "SINGLE SHORT-TERM DAYLIGHT STATIONARY PERIOD" in n
        or "DAYLIGHT SHORT-TERM PERIOD. STATIONARY" in n
        or re.search(r"\bIN URBAN BE ADJUSTED CONDITIONS\b", n)
        or re.search(r"\bIN SIGN TO URBAN\b", n)
        or re.search(r"\bFLAGGER DRIVEWAYS\b", n)
        or re.search(r"\bFLAGGER REMOVED, SYMBOL\b", n)
        or re.search(r"\bNEEDS NEEDS PRIOR\b", n)
    )
    if rot_garble or interleaved >= max(2, len(notes) // 3):
        return True, "drawing", (
            f"rotated-page column artifacts ({rot_garble or interleaved}/{len(notes)}); "
            "confirm wording against PDF"
        )
    if prose_score(notes) == 0 and avg < 80:
        return True, "drawing", "weak prose score"
    return True, "verbatim", "ok"


def main() -> int:
    report = []
    for sheet in SHEETS:
        path = ROOT / "Data" / "sheet-specs" / f"619-{sheet}.json"
        if not path.exists():
            report.append((sheet, "SKIP", "no spec"))
            continue
        notes, kind, hint = extract_sheet(sheet)
        if hint.startswith("fail") or not notes:
            report.append((sheet, "FAIL", f"{hint}; kind={kind}; {notes[:1]}"))
            continue
        conf = hint.split(":", 1)[0]
        why = hint.split(":", 1)[1] if ":" in hint else ""
        # Re-check quality for confidence
        ok, conf, why = quality(notes)
        if not ok:
            report.append((sheet, "FAIL", f"{why}; kind={kind}; {notes[:1]}"))
            continue

        spec = json.loads(path.read_text(encoding="utf-8"))
        old = spec.get("notes") if isinstance(spec.get("notes"), dict) else {}
        new = {
            "confidence": conf,
            "printed": notes,
            "backfillNote": (
                "Verbatim notes backfilled 2026-08-03 from PDF text layer "
                "(Family 6 follow-up per Claude assessment). "
                "Placement rules[] were already populated; this restores printed prose. "
                f"Extract geometry={kind}."
                + (
                    " Rotated-page notes may retain minor column-order artifacts — "
                    "confirm against PDF for edge-case wording."
                    if conf == "drawing"
                    else ""
                )
            ),
        }
        if old.get("planCallouts"):
            new["planCallouts"] = old["planCallouts"]
        spec["notes"] = new
        path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        report.append((sheet, f"OK/{conf}", f"{len(notes)} notes | {notes[0][:68]}"))

    print("Family 6 notes backfill")
    for sheet, status, detail in report:
        print(f"  {sheet}: {status} — {detail}")
    fails = sum(1 for _, s, _ in report if s.startswith("FAIL"))
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
