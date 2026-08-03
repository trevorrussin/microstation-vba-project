"""One-off Family 4 Parkway sheet recon via PyMuPDF."""
from __future__ import annotations

import json
import re
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
CAP = ROOT / "Bridge" / "captures"
SHEETS = ["619-306", "619-212", "619-114", "619-041"]

SIGN_PAT = re.compile(
    r"\b(W20[-\dA-Za-z]*|W21[-\dA-Za-z]*|R\d[-\dA-Za-z]*|NY[A-Z0-9][-\dA-Za-z]*|"
    r"G20[-\dA-Za-z]*|W4[-\dA-Za-z]*|W8[-\dA-Za-z]*|W7[-\dA-Za-z]*|W3[-\dA-Za-z]*)\b",
    re.I,
)
TABLE_PAT = re.compile(r"TABLE\s+(\d{3}-\d{2})", re.I)
SPACING_PAT = re.compile(r"(?:^|[^\d])(1000|1320|1500|2640|500)(?:'|FT| ft)?(?:[^\d]|$)", re.I)

KEYWORDS = {
    "merging_taper": re.compile(r"MERGING\s+TAPER", re.I),
    "shoulder_taper": re.compile(r"SHOULDER\s+TAPER", re.I),
    "downstream_taper": re.compile(r"DOWNSTREAM\s+TAPER", re.I),
    "flagger": re.compile(r"FLAGGER|FLAGGING|AFAD", re.I),
    "parkway": re.compile(r"PARKWAY", re.I),
    "short_term": re.compile(r"SHORT\s+TERM", re.I),
    "short_duration": re.compile(r"SHORT\s+DURATION", re.I),
    "mobile": re.compile(r"\bMOBILE\b", re.I),
    "mowing": re.compile(r"MOWING|MULCH", re.I),
    "shoulder_lt8": re.compile(r"SHOULDER\s*<\s*8|SHOULDER\s+LESS\s+THAN\s+8", re.I),
    "channelizing": re.compile(r"CHANNELIZ", re.I),
    "roll_ahead": re.compile(r"ROLL\s+AHEAD|ROLL-AHEAD", re.I),
    "buffer": re.compile(r"BUFFER", re.I),
    "workspace": re.compile(r"WORK\s*SPACE|WORKSPACE", re.I),
}


def squash(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip())


def role_from_region(rt: str) -> str:
    if "ORDER OF PLACEMENT" in rt or "PLACEMENT ORDER" in rt:
        return "orderOfPlacement"
    if "ROLL AHEAD" in rt or "ROLL-AHEAD" in rt:
        return "rollAhead"
    if "TAPER" in rt and "BUFFER" in rt and ("LANE WIDTH" in rt or "SHOULDER WIDTH" in rt):
        return "taperAndBuffer"
    if "SIGN SPACING" in rt or ("ADVANCE WARNING" in rt and "DISTANCE" in rt):
        return "signSpacing"
    if "SIGN SIZE" in rt or "619-012" in rt:
        return "signSize"
    if "CHANNELIZ" in rt or "DEVICE SPACING" in rt:
        return "channelizing"
    if "PROTECTIVE VEHICLE" in rt and "TAPER" not in rt and "BUFFER" not in rt:
        return "protectiveVehicle"
    if "MOWING" in rt or ("OPERATION" in rt and "SPEED" in rt):
        return "operation/mowing"
    if "SHOULDER" in rt and "TAPER" in rt:
        return "shoulderTaperOnly"
    return "unknown"


def extract_table_regions(page: fitz.Page, pi: int) -> list[dict]:
    words = page.get_text("words")
    table_words = [w for w in words if w[4].upper() == "TABLE"]
    out = []
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
        region_text = " ".join(
            w[4] for w in sorted(region_words, key=lambda w: (w[1], w[0]))
        )
        rt = region_text.upper()
        header = " ".join(
            w[4]
            for w in sorted(region_words, key=lambda w: (w[1], w[0]))[:50]
        )
        out.append(
            {
                "id": tid,
                "page": pi,
                "roleGuess": role_from_region(rt),
                "headerSnippet": header[:280],
                "yBand": [round(y0, 1), round(y1, 1)],
            }
        )
    return out


def analyze(sheet: str) -> dict:
    pdf_path = CAP / f"{sheet}.pdf"
    doc = fitz.open(pdf_path)
    result = {
        "sheet": sheet,
        "path": str(pdf_path),
        "bytes": pdf_path.stat().st_size,
        "pageCount": len(doc),
        "pages": [],
        "tablesListed": [],
        "signs": set(),
        "keywords": {},
        "spacingTokens": set(),
        "titleLines": [],
        "durationHints": [],
        "roadTypeHints": [],
        "notesCount": 0,
    }
    all_text = []
    table_roles = []
    for pi, page in enumerate(doc):
        rot = page.rotation
        mb = page.mediabox
        text = page.get_text("text")
        words = page.get_text("words")
        all_text.append(text)
        top_words = sorted([w for w in words if w[1] < 120], key=lambda w: (w[1], w[0]))
        top_lines = []
        cur_y = None
        cur: list[str] = []
        for w in top_words:
            if cur_y is None or abs(w[1] - cur_y) < 4:
                cur.append(w[4])
                cur_y = w[1] if cur_y is None else cur_y
            else:
                if cur:
                    top_lines.append(" ".join(cur))
                cur = [w[4]]
                cur_y = w[1]
        if cur:
            top_lines.append(" ".join(cur))
        result["pages"].append(
            {
                "index": pi,
                "rotation": rot,
                "mediabox": [round(mb.x0, 1), round(mb.y0, 1), round(mb.x1, 1), round(mb.y1, 1)],
                "wordCount": len(words),
                "charCount": len(text),
                "topTitleBlock": top_lines[:15],
            }
        )
        for m in TABLE_PAT.finditer(text):
            result["tablesListed"].append({"id": m.group(1).upper(), "page": pi})
        for m in SIGN_PAT.finditer(text):
            result["signs"].add(m.group(1).upper().replace(" ", ""))
        for m in SPACING_PAT.finditer(text):
            result["spacingTokens"].add(m.group(1))
        table_roles.extend(extract_table_regions(page, pi))

    full = "\n".join(all_text)
    for k, pat in KEYWORDS.items():
        hits = pat.findall(full)
        result["keywords"][k] = bool(hits)
        if hits and k in (
            "merging_taper",
            "shoulder_taper",
            "downstream_taper",
            "flagger",
            "parkway",
        ):
            result["keywords"][k + "_count"] = len(hits)

    for line in full.splitlines():
        u = line.upper().strip()
        if not u:
            continue
        if any(
            x in u
            for x in [
                "SHORT TERM",
                "SHORT DURATION",
                "MOBILE",
                "MOWING",
                "PARKWAY",
                "FREEWAY",
                "NON-FREEWAY",
                "MULTILANE",
                "UNDIVIDED",
                "DIVIDED",
            ]
        ):
            if len(u) < 120:
                if any(x in u for x in ["SHORT TERM", "SHORT DURATION", "MOBILE", "MOWING"]):
                    result["durationHints"].append(squash(line))
                if any(
                    x in u
                    for x in [
                        "PARKWAY",
                        "FREEWAY",
                        "NON-FREEWAY",
                        "MULTILANE",
                        "UNDIVIDED",
                        "DIVIDED",
                        "ROADWAY",
                    ]
                ):
                    result["roadTypeHints"].append(squash(line))
        if u.startswith("NOTE ") or u.startswith("NOTES"):
            result["notesCount"] += 1

    lines = [squash(l) for l in all_text[0].splitlines() if squash(l)]
    for l in lines[:30]:
        if any(
            x in l.upper()
            for x in [
                "WORK ZONE",
                "LANE CLOSURE",
                "PARKWAY",
                "MOWING",
                "MOBILE",
                "SHOULDER",
                "RIGHT",
                "LEFT",
            ]
        ):
            result["titleLines"].append(l)

    plan_labels = []
    for pat in [
        r"MERGING TAPER",
        r"SHOULDER TAPER",
        r"DOWNSTREAM TAPER",
        r"WORK SPACE",
        r"VEHICLE SPACE",
        r"BUFFER",
        r"ROLL AHEAD",
        r"CHANNELIZ",
        r"PROTECTIVE VEHICLE",
        r"ARROW PANEL",
        r"\b[ABCD]\s*=\s*\d+",
        r"\b\d{3,4}'\b",
        r"L/\d",
    ]:
        for m in re.finditer(pat, full, re.I):
            start = max(0, m.start() - 35)
            plan_labels.append(squash(full[start : m.end() + 35]))

    result["signs"] = sorted(result["signs"])
    result["spacingTokens"] = sorted(result["spacingTokens"], key=int)
    result["tableRoles"] = table_roles
    result["planLabels"] = list(dict.fromkeys(plan_labels))[:30]
    result["operationLine"] = next(
        (squash(l) for l in full.splitlines() if "OPERATION" in l.upper() and len(l) < 80),
        "",
    )
    doc.close()
    return result


def compare_families(r: dict) -> dict:
    """Rough similarity scoring vs Family 1/2/3 reference sheets."""
    k = r["keywords"]
    signs = set(r["signs"])
    tables = [t["roleGuess"] for t in r["tableRoles"]]
    score = {"family1_311": 0, "family2_302": 0, "family3_301": 0, "notes": []}

    if k.get("merging_taper"):
        score["family2_302"] += 3
        score["family1_311"] += 2
        score["notes"].append("Has MERGING TAPER -> lane-closure family shape")
    if k.get("shoulder_taper") and not k.get("merging_taper"):
        score["family3_301"] += 2
        score["family1_311"] += 1
    if k.get("downstream_taper"):
        score["family2_302"] += 1
        score["family1_311"] += 1
    if "signSpacing" in tables:
        score["family2_302"] += 2
        score["family1_311"] += 2
    if "rollAhead" in tables:
        score["family3_301"] += 3
    if k.get("mobile"):
        score["family3_301"] += 2
        score["notes"].append("Mobile operation -> closer to 619-205 mobile shoulder pattern")
    if k.get("mowing"):
        score["notes"].append("Mowing sheet -> minimal taper/table stack")
    if any(s.startswith("W21-") for s in signs):
        score["family3_301"] += 2
    if any(s.startswith("NYW8") for s in signs):
        score["family2_302"] += 1
        score["family1_311"] += 1
    if "1320" in r["spacingTokens"] or "1500" in r["spacingTokens"]:
        score["family3_301"] += 1
    if "1000" in r["spacingTokens"]:
        score["family3_301"] += 1
        score["notes"].append("1000' spacing token present")
    if "2640" in r["spacingTokens"]:
        score["family3_301"] += 1

    best = max(score, key=lambda x: score[x] if x != "notes" else -1)
    score["closestFamily"] = best.replace("_", " ").upper()
    return score


def table_titles_from_text(text: str) -> list[dict]:
    out = []
    for m in re.finditer(r"TABLE\s+(\d{3}-\d{2})\s*:\s*([^\n]+)", text, re.I):
        out.append({"id": m.group(1).upper(), "title": squash(m.group(2))})
    # dedupe preserving order
    seen = set()
    deduped = []
    for t in out:
        if t["id"] not in seen:
            seen.add(t["id"])
            deduped.append(t)
    return deduped


def deep_extract(sheet: str) -> dict:
    doc = fitz.open(CAP / f"{sheet}.pdf")
    text = "\n".join(page.get_text("text") for page in doc)
    doc.close()
    signs = sorted(set(re.findall(r"\b(?:NY[A-Z0-9][-\w]*|W\d{1,2}[-\w]*|R\d[-\w]*|G20[-\w]*)\b", text, re.I)))
    spacing = sorted(set(re.findall(r"(?<!\d)(1000|1320|1500|2640|500)(?:'| FT)?", text)), key=int)
    phrases = {}
    for pat in [
        "MERGING TAPER",
        "SHOULDER TAPER",
        "DOWNSTREAM TAPER",
        "FLAGGER",
        "PARKWAY",
        "MOBILE",
        "MOWING",
        "SHORT TERM",
        "SHORT DURATION",
        "MOVING OPERATION",
        "STATIONARY OPERATION",
        "ADVANCE WARNING",
        "SIGN SPACING",
        "ORDER OF PLACEMENT",
        "SHOULDER < 8",
        "SHOULDER LESS THAN 8",
    ]:
        phrases[pat] = pat in text.upper()
    title_lines = [
        squash(l)
        for l in text.splitlines()
        if any(x in l.upper() for x in ["WORK ZONE TRAFFIC", "LANE CLOSURE", "PARKWAY", "MOWING", "ENCROACHMENT"])
        and len(l.strip()) < 100
    ]
    return {
        "tableTitles": table_titles_from_text(text),
        "signsAll": signs,
        "spacing": spacing,
        "phrases": phrases,
        "titleLines": list(dict.fromkeys(title_lines))[:8],
    }


if __name__ == "__main__":
    reports = {}
    for s in SHEETS:
        r = analyze(s)
        r["familySimilarity"] = compare_families(r)
        r["deep"] = deep_extract(s)
        reports[s] = r

    # quick reference comparison for freeway siblings not in Family 4
    ref_compare = {}
    for ref in ["619-302", "619-304", "619-206", "619-111", "619-205", "619-031"]:
        p = CAP / f"{ref}.pdf"
        if not p.exists():
            ref_compare[ref] = {"missing": True}
            continue
        doc = fitz.open(p)
        text = doc[0].get_text("text")
        pages = len(doc)
        doc.close()
        ref_compare[ref] = {
            "pages": pages,
            "tables": table_titles_from_text(text),
            "phrases": {k: k in text.upper() for k in ["MERGING TAPER", "SHOULDER TAPER", "PARKWAY", "MOBILE"]},
        }

    print(json.dumps({"family4": reports, "referenceSheets": ref_compare}, indent=2))
