"""Family 5/6 sheet recon via PyMuPDF — writes Bridge/_recon_f5_f6.md."""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
CAP = ROOT / "Bridge" / "captures"
OUT = ROOT / "Bridge" / "_recon_f5_f6.md"

FAMILY5 = ["619-318", "619-316", "619-319", "619-113", "619-211", "619-416", "619-417", "619-418", "619-517", "619-518"]
FAMILY6 = [
    "619-307", "619-308", "619-309", "619-314", "619-321", "619-322", "619-323", "619-324",
    "619-407", "619-419", "619-420", "619-421", "619-422", "619-519", "619-520", "619-524",
    "619-090", "619-091",
]

SIGN_PAT = re.compile(
    r"\b(W20[-\dA-Za-z]*|W21[-\dA-Za-z]*|R\d[-\dA-Za-z]*|NY[A-Z0-9][-\dA-Za-z]*|"
    r"G20[-\dA-Za-z]*|W4[-\dA-Za-z]*|W8[-\dA-Za-z]*|W7[-\dA-Za-z]*|W3[-\dA-Za-z]*)\b",
    re.I,
)
TABLE_PAT = re.compile(r"TABLE\s+(\d{3}-\d{2})", re.I)
TABLE_TITLE_PAT = re.compile(r"TABLE\s+(\d{3}-\d{2})\s*:?\s*([^\n]{5,120})", re.I)
SPACING_PAT = re.compile(r"(?<!\d)(1000|1320|1500|2640|500)(?:'| FT| ft)?(?!\d)", re.I)

KEYWORDS = {
    "merging_taper": re.compile(r"MERGING\s+TAPER", re.I),
    "shoulder_taper": re.compile(r"SHOULDER\s+TAPER", re.I),
    "downstream_taper": re.compile(r"DOWNSTREAM\s+TAPER", re.I),
    "flagger": re.compile(r"\bFLAGGER|\bFLAGGING\b", re.I),
    "afad": re.compile(r"\bAFAD\b|AUTOMATED\s+FLAGGER", re.I),
    "sidewalk": re.compile(r"SIDEWALK|PEDESTRIAN\s+DETOUR|PED\s+DETOUR", re.I),
    "crosswalk": re.compile(r"CROSSWALK", re.I),
    "mobile": re.compile(r"\bMOBILE\b|MOVING\s+OPERATION", re.I),
    "ramp": re.compile(r"\bRAMP\b|ENTRANCE\s+RAMP|EXIT\s+RAMP", re.I),
    "twlt": re.compile(r"TWLT|TWO[- ]WAY\s+LEFT[- ]TURN", re.I),
    "centerline": re.compile(r"CENTERLINE|CENTER\s+LINE", re.I),
    "signal": re.compile(r"\bSIGNAL\b|TRAFFIC\s+SIGNAL", re.I),
    "closure": re.compile(r"ROAD\s+CLOSURE|INTERSECTION\s+CLOSURE", re.I),
    "protective_vehicle": re.compile(r"PROTECTIVE\s+VEHICLE", re.I),
    "arrow_panel": re.compile(r"ARROW\s+PANEL", re.I),
    "channelizing": re.compile(r"CHANNELIZ", re.I),
}


def squash(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip())


def role_from_region(rt: str) -> str:
    if "ORDER OF PLACEMENT" in rt or "PLACEMENT ORDER" in rt:
        return "orderOfPlacement"
    if "ROLL AHEAD" in rt or "ROLL-AHEAD" in rt:
        return "rollAhead"
    if "TAPER" in rt and "BUFFER" in rt:
        return "taperAndBuffer"
    if "SIGN SPACING" in rt or ("ADVANCE WARNING" in rt and "DISTANCE" in rt):
        return "signSpacing"
    if "SIGN SIZE" in rt or "619-012" in rt:
        return "signSize"
    if "CHANNELIZ" in rt or "DEVICE SPACING" in rt:
        return "channelizing"
    if "PROTECTIVE VEHICLE" in rt:
        return "protectiveVehicle"
    if "FLAGGER" in rt or "AFAD" in rt:
        return "flagger"
    if "SIDEWALK" in rt or "PEDESTRIAN" in rt:
        return "pedestrian"
    return "unknown"


def table_titles_from_text(text: str) -> list[dict]:
    out = []
    seen = set()
    for m in TABLE_TITLE_PAT.finditer(text):
        tid = m.group(1).upper()
        if tid in seen:
            continue
        seen.add(tid)
        out.append({"id": tid, "title": squash(m.group(2))})
    return out


def title_keywords(text: str) -> list[str]:
    hits = []
    for pat in [
        r"WORK ZONE TRAFFIC CONTROL",
        r"LANE CLOSURE",
        r"SINGLE LANE",
        r"TWO[- ]LANE TWO[- ]WAY",
        r"FLAGGER",
        r"AFAD",
        r"ENTRANCE RAMP",
        r"EXIT RAMP",
        r"SHOULDER",
        r"FREEWAY",
        r"NON[- ]FREEWAY",
        r"INTERMEDIATE",
        r"LONG TERM",
        r"SHORT TERM",
        r"SIDEWALK",
        r"CROSSWALK",
        r"ROAD CLOSURE",
        r"INTERSECTION",
        r"SIGNAL",
        r"MOBILE",
    ]:
        if re.search(pat, text, re.I):
            hits.append(re.sub(r"\\[- \\]", "-", pat).replace("\\b", "").replace("\\s+", " "))
    return hits


def compare_families(r: dict) -> str:
    k = r["keywords"]
    signs = set(r["signs"])
    tables = [t["roleGuess"] for t in r["tableRoles"]]
    score = {"311": 0, "302": 0, "301": 0}

    if k.get("merging_taper"):
        score["302"] += 3
        score["311"] += 2
    if k.get("shoulder_taper") and not k.get("merging_taper"):
        score["301"] += 2
        score["311"] += 1
    if k.get("downstream_taper"):
        score["302"] += 2
        score["311"] += 1
    if "signSpacing" in tables:
        score["302"] += 2
        score["311"] += 2
    if k.get("flagger") or k.get("afad"):
        score["311"] += 1
    if k.get("protective_vehicle") and k.get("arrow_panel"):
        score["302"] += 2
        score["311"] += 1
    if any(s.startswith("NYW8") for s in signs):
        score["302"] += 2
        score["311"] += 1
    if any(s.startswith("W21-") for s in signs):
        score["301"] += 2
    if "1320" in r["spacingTokens"] or "1500" in r["spacingTokens"]:
        score["301"] += 1
    if "1000" in r["spacingTokens"]:
        score["301"] += 1
    if "2640" in r["spacingTokens"]:
        score["301"] += 1
    if k.get("ramp"):
        score["302"] += 1

    best = max(score, key=score.get)
    return f"619-{best} (score {score[best]})"


def operational_flagger(text: str) -> bool:
    u = text.upper()
    if re.search(r"LANE CLOSURE WITH FLAGGER|FLAGGING OPERATION|AUTOMATED FLAGGER|\bAFAD\b", u):
        return True
    if re.search(r"FLAGGER SYMBOL|NON-ILLUMINATED FLAGGER", u):
        return False
    return bool(re.search(r"\bFLAGGER|\bFLAGGING\b", u))


def schema_class(r: dict, family: int) -> str:
    k = r["keywords"]
    tags = []
    if k.get("afad"):
        tags.append("AFAD")
    elif k.get("flagger") and r.get("operationalFlagger"):
        tags.append("flagger")
    if k.get("sidewalk"):
        tags.append("sidewalk detour")
    if k.get("crosswalk"):
        tags.append("crosswalk")
    if k.get("closure"):
        tags.append("road/intersection closure")
    if k.get("signal"):
        tags.append("signal")
    if k.get("mobile"):
        tags.append("mobile-only")
    if k.get("twlt"):
        tags.append("TWLT lane shift")

    if family == 5:
        if tags:
            return "corridor clone + " + ", ".join(tags)
        return "corridor lane-closure clone (ramp-adjacent)"

    if tags:
        novel = [t for t in tags if t not in ("corridor clone")]
        if novel:
            return "NOVEL: " + ", ".join(novel)
    if k.get("merging_taper") and not k.get("flagger"):
        return "corridor lane-closure clone"
    if k.get("flagger") and not k.get("sidewalk") and not k.get("crosswalk"):
        return "flagger corridor (307-base)"
    return "mixed / review"


def analyze(sheet: str) -> dict | None:
    pdf_path = CAP / f"{sheet}.pdf"
    if not pdf_path.exists():
        return None
    doc = fitz.open(pdf_path)
    result = {
        "sheet": sheet,
        "path": str(pdf_path),
        "bytes": pdf_path.stat().st_size,
        "extractableChars": 0,
        "pageCount": len(doc),
        "rotations": [],
        "tablesListed": [],
        "tableTitles": [],
        "tableRoles": [],
        "signs": set(),
        "keywords": {},
        "spacingTokens": set(),
        "titleLines": [],
        "titleKeywords": [],
    }
    all_text = []
    for pi, page in enumerate(doc):
        rot = page.rotation
        result["rotations"].append(rot)
        text = page.get_text("text")
        all_text.append(text)
        for m in TABLE_PAT.finditer(text):
            result["tablesListed"].append({"id": m.group(1).upper(), "page": pi})
        for m in SIGN_PAT.finditer(text):
            result["signs"].add(m.group(1).upper().replace(" ", ""))
        for m in SPACING_PAT.finditer(text):
            result["spacingTokens"].add(m.group(1))
        words = page.get_text("words")
        table_words = [w for w in words if w[4].upper() == "TABLE"]
        seen = set()
        for tw in sorted(table_words, key=lambda w: w[1]):
            line_words = sorted(
                [w for w in words if abs(w[1] - tw[1]) < 4 and w[0] >= tw[0] - 5],
                key=lambda w: w[0],
            )
            line = " ".join(w[4] for w in line_words)
            m = TABLE_PAT.search(line)
            if not m:
                continue
            tid = m.group(1).upper()
            if tid in seen:
                continue
            seen.add(tid)
            y0 = tw[1]
            next_tables = [w for w in table_words if w[1] > tw[1] + 10]
            y1 = min([w[1] - 5 for w in next_tables], default=page.rect.height)
            region_words = [w for w in words if y0 <= w[1] < y1]
            region_text = " ".join(w[4] for w in sorted(region_words, key=lambda w: (w[1], w[0])))
            result["tableRoles"].append({
                "id": tid,
                "page": pi,
                "roleGuess": role_from_region(region_text.upper()),
            })

    full = "\n".join(all_text)
    result["extractableChars"] = len(full.strip())
    result["imageOnly"] = result["extractableChars"] < 200
    for k, pat in KEYWORDS.items():
        result["keywords"][k] = bool(pat.search(full))
    result["operationalFlagger"] = operational_flagger(full)

    result["tableTitles"] = table_titles_from_text(full)
    result["titleKeywords"] = title_keywords(full)
    for line in full.splitlines():
        u = line.upper().strip()
        if not u or len(u) > 100:
            continue
        if any(x in u for x in [
            "WORK ZONE", "LANE CLOSURE", "FLAGGER", "RAMP", "SHOULDER",
            "SIDEWALK", "CROSSWALK", "TWO-LANE", "TWO LANE", "AFAD",
            "ROAD CLOSURE", "INTERSECTION",
        ]):
            result["titleLines"].append(squash(line))

    result["signs"] = sorted(result["signs"])[:20]
    result["spacingTokens"] = sorted(result["spacingTokens"], key=int)
    result["titleLines"] = list(dict.fromkeys(result["titleLines"]))[:6]
    doc.close()
    return result


def kw_line(k: dict, r: dict | None = None) -> str:
    parts = []
    for name, label in [
        ("merging_taper", "MERGING"),
        ("shoulder_taper", "SHOULDER"),
        ("downstream_taper", "DOWNSTREAM"),
        ("afad", "AFAD"),
        ("sidewalk", "SIDEWALK"),
        ("crosswalk", "CROSSWALK"),
        ("mobile", "MOBILE"),
        ("ramp", "RAMP"),
        ("twlt", "TWLT"),
        ("closure", "CLOSURE"),
        ("signal", "SIGNAL"),
    ]:
        if k.get(name):
            parts.append(label)
    if k.get("flagger"):
        if r and r.get("operationalFlagger"):
            parts.append("FLAGGER")
        elif not r:
            parts.append("FLAGGER(notes?)")
    return ", ".join(parts) if parts else "none"


def sheet_brief(r: dict, family: int) -> str:
    if r.get("imageOnly"):
        return (
            f"- **Pages/rotation:** {r['pageCount']} / {r['rotations'][0] if r['rotations'] else '?'}°\n"
            f"- **Status:** image-only PDF ({r['bytes']//1024} KB, {r['extractableChars']} extractable chars) — needs OCR or visual review\n"
            f"- **Schema:** expect crosswalk (321 sibling) — **PDF unreadable by text extract**\n"
        )
    rot = r["rotations"]
    rot_s = f"{rot[0]}°" if len(set(rot)) == 1 else str(rot)
    tables = r["tableTitles"] or [{"id": t["id"], "title": "?"} for t in {t["id"]: t for t in r["tablesListed"]}.values()]
    table_str = "; ".join(f"{t['id']}: {t.get('title', '?')[:60]}" for t in tables[:8])
    if len(tables) > 8:
        table_str += f" (+{len(tables)-8} more)"
    signs = ", ".join(r["signs"][:12])
    if len(r["signs"]) > 12:
        signs += f" (+{len(r['signs'])-12})"
    spacing = ", ".join(f"{s}'" for s in r["spacingTokens"]) or "—"
    closest = compare_families(r)
    schema = schema_class(r, family)
    titles = "; ".join(r["titleLines"][:3]) or "; ".join(r["titleKeywords"][:5])
    return (
        f"- **Pages/rotation:** {r['pageCount']} / {rot_s}\n"
        f"- **Title keywords:** {titles}\n"
        f"- **TABLE titles:** {table_str or 'none detected'}\n"
        f"- **Key signs:** {signs or 'none'}\n"
        f"- **Language flags:** {kw_line(r['keywords'], r)}\n"
        f"- **Spacing tokens:** {spacing}\n"
        f"- **Closest family:** {closest}\n"
        f"- **Schema:** {schema}\n"
    )


def ref_confirmation(r318: dict | None, r307: dict | None) -> str:
    lines = ["## Reference confirmation\n"]
    if r318:
        k = r318["keywords"]
        lines.append("### 619-318 (Family 5 reference)")
        lines.append(
            f"Confirm as F5 anchor: **{'YES' if k.get('merging_taper') and k.get('ramp') else 'REVIEW'}**. "
            f"{r318['pageCount']} pg, tables {[t['id'] for t in r318['tableTitles']]}, "
            f"MERGING={k.get('merging_taper')}, RAMP={k.get('ramp')}, "
            f"DOWNSTREAM={k.get('downstream_taper')}, PV/AP implied via keywords "
            f"PV={k.get('protective_vehicle')} AP={k.get('arrow_panel')}."
        )
        lines.append(
            "Registry expects Tables 318-01..06 with merging/downstream taper + channelizing — "
            "PDF " + ("matches" if len(r318["tableTitles"]) >= 4 else "partial match; verify table count") + "."
        )
    if r307:
        k = r307["keywords"]
        lines.append("\n### 619-307 (Family 6 reference)")
        lines.append(
            f"Confirm as F6 anchor: **{'YES' if k.get('flagger') else 'REVIEW'}**. "
            f"{r307['pageCount']} pg, tables {[t['id'] for t in r307['tableTitles']]}, "
            f"FLAGGER={k.get('flagger')}, AFAD={k.get('afad')}, "
            f"no merging taper expected for base flagger sheet."
        )
        lines.append(
            "Registry: Two-Lane Two-Way flagger operation, Tables 307-01..03 — "
            + ("PDF aligns." if k.get("flagger") and len(r307["tableTitles"]) >= 2 else "verify table stack.")
        )
    return "\n".join(lines)


def main() -> None:
    reports5 = {}
    reports6 = {}
    missing5 = []
    missing6 = []

    for s in FAMILY5:
        r = analyze(s)
        if r:
            r["closest"] = compare_families(r)
            r["schema"] = schema_class(r, 5)
            reports5[s] = r
        else:
            missing5.append(s)

    for s in FAMILY6:
        r = analyze(s)
        if r:
            r["closest"] = compare_families(r)
            r["schema"] = schema_class(r, 6)
            reports6[s] = r
        else:
            missing6.append(s)

    lines = [
        f"# Family 5 & 6 PDF Recon",
        f"",
        f"Generated {date.today().isoformat()} via PyMuPDF on `Bridge/captures/*.pdf`.",
        f"",
        f"**Family 5** (ref 619-318): ramp-adjacent single lane closure — {len(reports5)}/{len(FAMILY5)} PDFs available.",
        f"**Family 6** (ref 619-307): two-lane two-way / flagger / pedestrian — {len(reports6)}/{len(FAMILY6)} PDFs available.",
        f"",
    ]
    if missing5:
        lines.append(f"**Missing F5 PDFs:** {', '.join(missing5)}")
    if missing6:
        lines.append(f"**Missing F6 PDFs:** {', '.join(missing6)}")
    unreadable = [s for s, r in reports6.items() if r.get("imageOnly")]
    if unreadable:
        lines.append(f"**Image-only F6 PDFs (no text layer):** {', '.join(unreadable)}")
    lines.append("")

    lines.append(ref_confirmation(reports5.get("619-318"), reports6.get("619-307")))
    lines.append("")

    lines.append("## Schema classification summary\n")
    lines.append("| Sheet | Family | Schema | Closest |")
    lines.append("|---|---|---|---|")
    for s in FAMILY5:
        if s in reports5:
            r = reports5[s]
            lines.append(f"| {s} | F5 | {r['schema']} | {r['closest']} |")
        else:
            lines.append(f"| {s} | F5 | **PDF missing** | — |")
    for s in FAMILY6:
        if s in reports6:
            r = reports6[s]
            sch = r["schema"]
            if r.get("imageOnly"):
                sch = "NOVEL: crosswalk (image-only PDF)"
            lines.append(f"| {s} | F6 | {sch} | {r['closest']} |")
        else:
            lines.append(f"| {s} | F6 | **PDF missing** | — |")
    lines.append("")

    lines.append("## Family 5 — per-sheet briefs\n")
    for s in FAMILY5:
        lines.append(f"### {s}")
        if s in reports5:
            lines.append(sheet_brief(reports5[s], 5))
        else:
            lines.append("- **Status:** PDF not in captures/\n")
        lines.append("")

    lines.append("## Family 6 — per-sheet briefs\n")
    for s in FAMILY6:
        lines.append(f"### {s}")
        if s in reports6:
            lines.append(sheet_brief(reports6[s], 6))
        else:
            lines.append("- **Status:** PDF not in captures/\n")
        lines.append("")

    lines.append("## Recommendations\n")
    lines.append(
        "1. **619-318** — Use as F5 structural reference if tables 318-01..06 + merging/downstream taper "
        "corridor zones match 619-302 shape with ramp callouts; clone 302 corridor schema then add ramp labels."
    )
    lines.append(
        "2. **619-307** — Use as F6 base for flagger tables/sign spacing; do **not** reuse 302 merging taper "
        "schema — flagger sheets need `Flagger`/`CenterlineCones` zones and optional PV."
    )
    lines.append("3. **Novel schema required (F6):**")
    novel = [(s, reports6[s]["schema"]) for s in FAMILY6 if s in reports6 and "NOVEL" in reports6[s]["schema"]]
    missing_novel = [s for s in ["619-419", "619-420", "619-519", "619-520", "619-524", "619-090", "619-091"] if s not in reports6]
    for s, sch in novel:
        lines.append(f"   - {s}: {sch}")
    for s in missing_novel:
        reg_note = {
            "619-419": "intermediate sidewalk",
            "619-420": "intermediate crosswalk",
            "619-519": "long-term sidewalk",
            "619-520": "long-term crosswalk",
            "619-524": "long-term signal",
            "619-090": "temporary road closure",
            "619-091": "temporary intersection closure",
        }.get(s, "")
        lines.append(f"   - {s}: **PDF missing** — expect NOVEL ({reg_note})")
    lines.append("4. **Corridor clones (F5):** 316, 319, 416–418, 517–518 likely share 318 table stack with ramp/shoulder variants.")
    lines.append("5. **Corridor clones (F6):** 308, 407, 421, 422 may share 307 flagger base; 314 moving flaggers may need mobile variant.")
    if missing6:
        lines.append("6. **Download still needed:** " + ", ".join(missing6))
    if unreadable:
        lines.append(f"7. **OCR/visual review needed:** {', '.join(unreadable)} (large image PDFs with no extractable text).")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT} ({len(lines)} lines)")


if __name__ == "__main__":
    main()
