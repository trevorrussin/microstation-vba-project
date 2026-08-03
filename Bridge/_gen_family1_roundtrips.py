"""Generate Bridge/roundtrip/619-{n}.py for Family 1 sheets.

Each script re-extracts key cells from the PDF and diffs against the JSON.
Shared tables that match 311 are compared via slash-cell identity; sheet-specific
bits (titles, sign codes, note phrases) are phrase-checked.
"""
from __future__ import annotations

import pathlib
import textwrap

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "Bridge/roundtrip"

# sheet -> (tables_page_idx, taper_box, roll_box, aw_box, size_codes, note_phrases, extras)
# boxes are (x0,y0,x1,y1) on the tables page; None = skip that check
CONFIG = {
    "202": {
        "page": 0, "rotation": 0,
        "taper": None,
        "roll": (1040, 40, 1224, 160),
        "aw": (990, 240, 1224, 340),
        "aw_has_c": False,
        "size_codes": ["NYW8-33", "W4-2L", "W20-1", "WARNING FLAG"],
        "phrases": ["SHORT DURATION", "OPERATOR", "ROLL AHEAD"],
        "pv_phrase": "P, TMIA",
    },
    "203": {
        "page": 0, "rotation": 0,
        "taper": None,
        "roll": (1040, 40, 1224, 160),
        "aw": (990, 240, 1224, 340),
        "aw_has_c": False,
        "size_codes": ["NYW8-33", "W4-2R", "W20-1", "WARNING FLAG"],
        "phrases": ["SHORT DURATION", "OPERATOR", "ROLL AHEAD"],
        "pv_phrase": "P, TMIA",
    },
    "312": {
        "page": 1, "rotation": 0,
        "taper": (90, 590, 400, 720),
        "taper_has_shoulder": False,
        "roll": (90, 150, 400, 280),
        "aw": (90, 40, 450, 150),
        "aw_has_c": False,
        "size_codes": ["G20-2", "NYW8-33", "R4-7", "W4-2L", "W9-3", "W20-1", "W20-5"],
        "phrases": ["TWO-WAY LEFT TURN", "L/2", "SHORT-TERM"],
        "pv_phrase": "P, TMIA",
    },
    "317": {
        "page": 1, "rotation": 0,
        "taper": (100, 160, 780, 350),
        "taper_ytol": 8,
        "taper_has_shoulder": True,
        "roll": (90, 350, 450, 470),
        "aw": (100, 40, 500, 150),
        "aw_has_c": True,
        "size_codes": ["G20-2", "NYW8-33", "W4-2R", "W20-1", "W20-5"],
        "phrases": ["SINGLE LANE", "CHANNELIZING DEVICE APPLICATION", "MERGING"],
        "pv_phrase": "P, TMIA",
    },
    "325": {
        "page": 1, "rotation": 0,
        "taper": (100, 450, 780, 630),
        "taper_has_shoulder": True,
        "roll": (90, 320, 450, 440),
        "aw": (90, 630, 500, 760),
        "aw_has_c": True,
        "size_codes": ["G20-2", "NYW8-33", "W4-2R", "W20-1", "W20-5"],
        "phrases": ["DOUBLE INTERIOR", "CHANNELIZING DEVICE APPLICATION"],
        "pv_phrase": "P, TMIA",
    },
    "412": {
        "page": 1, "rotation": 270,
        "taper": None,  # rotation makes windows awkward — phrase + sibling identity
        "roll": None,
        "aw": None,
        "aw_has_c": False,
        "size_codes": ["G20-2", "NYR9-11", "W20-1", "W20-5", "W4-2L", "W9-3", "R4-7"],
        "phrases": ["INTERMEDIATE", "TWO-WAY LEFT TURN", "NOTE 2"],
        "pv_phrase": "NOTE 2",
        "sibling_taper_of": "311",
    },
    "414": {
        "page": 1, "rotation": 0,
        "taper": (100, 148, 780, 340),
        "taper_has_shoulder": True,
        "roll": (90, 340, 450, 460),
        "aw": (90, 40, 500, 140),
        "aw_has_c": True,
        "size_codes": ["G20-2", "NYR9-11", "NYW8-33", "W4-2R", "W20-1", "W20-5R"],
        "phrases": ["INTERMEDIATE", "CHANNELIZING DEVICE APPLICATION", "20"],
        "pv_phrase": "SEE NOTE 2",
    },
    "423": {
        "page": 1, "rotation": 0,
        "taper": (100, 143, 780, 330),
        "taper_has_shoulder": True,
        "roll": (90, 330, 450, 460),
        "aw": (90, 40, 500, 140),
        "aw_has_c": True,
        "size_codes": ["G20-2", "NYW8-33", "W4-2L", "W20-1", "W20-5"],
        "phrases": ["INTERMEDIATE", "DOUBLE INTERIOR", "CHANNELIZING"],
        "pv_phrase": "SEE NOTE 2",
    },
    "523": {
        "page": 1, "rotation": 0,
        "taper": (100, 143, 780, 330),
        "taper_has_shoulder": True,
        "roll": (90, 330, 450, 460),
        "aw": (90, 40, 500, 140),
        "aw_has_c": True,
        "size_codes": ["G20-2", "NYR9-11", "NYW8-33", "W4-2L", "W20-1", "W20-5"],
        "phrases": ["LONG TERM", "DOUBLE INTERIOR", "CHANNELIZING"],
        "pv_phrase": "SEE NOTE 2",
    },
}

TEMPLATE = r'''"""Round-trip check for Data/sheet-specs/619-{num}.json (Family 1).

Run: python Bridge/roundtrip/619-{num}.py
"""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, squash, assert_row_count, row_text

spec = json.loads((ROOT / "Data/sheet-specs/619-{num}.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
page = doc[{page}]
W = page.get_text("words")
text = page.get_text()
# Also join all pages for phrase checks
all_text = " ".join(doc[i].get_text() for i in range(len(doc)))
fails = []
roles = spec["tableRoles"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
lw = ["10", "11", "12"]


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{{label}}: PDF={{pdf!r}} JSON={{js!r}}")


def extract_taper(box, ytol=10, has_shoulder=True):
    raw = group_rows(words_in_window(W, *box), y_tol=3.0)
    merged = defaultdict(list)
    for r in raw:
        merged[round(r[0][1] / ytol)].extend(r)
    rows = []
    for k in sorted(merged):
        toks = [w[4] for w in sorted(merged[k], key=lambda w: w[0])]
        if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55):
            cells = [t for t in toks if "/" in t]
            rows.append((int(toks[0]), cells))
    return rows


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()

{body}

print("fails:", len(fails))
for f in fails:
    print(" ", f)
sys.exit(1 if fails else 0)
'''


def body_for(num: str, cfg: dict) -> str:
    lines = []
    # phrases
    lines.append("# ---- phrase checks")
    for ph in cfg["phrases"]:
        lines.append(f"if squash({ph!r}) not in squash(all_text):")
        lines.append(f"    fails.append('phrase missing: ' + {ph!r})")
    lines.append(f"if squash({cfg['pv_phrase']!r}) not in squash(all_text):")
    lines.append(f"    fails.append('PV phrase missing: ' + {cfg['pv_phrase']!r})")
    lines.append("print('phrases ok')")

    # size codes
    lines.append("# ---- sign size codes present in PDF text")
    lines.append(f"need = {cfg['size_codes']!r}")
    lines.append("for code in need:")
    lines.append("    if squash(code) not in squash(all_text) and squash(code.replace('-', '')) not in squash(all_text):")
    lines.append("        # WARNING FLAG often split across lines")
    lines.append("        if code == 'WARNING FLAG' and 'FLAG' in all_text.upper():")
    lines.append("            continue")
    lines.append("        fails.append(f'size code missing in PDF: {code!r}')")
    lines.append("print('size codes checked', need)")

    # JSON size table sync
    lines.append("sz = spec['tables'][roles['signSizes']]")
    lines.append("js_codes = [r['signCode'] for r in sz['rows']]")
    lines.append("for code in need:")
    lines.append("    if code not in js_codes and code != 'WARNING FLAG':")
    lines.append("        fails.append(f'JSON size table missing {code!r}')")
    lines.append("print('JSON size codes', js_codes)")

    if cfg.get("taper"):
        has_sh = cfg.get("taper_has_shoulder", True)
        ytol = cfg.get("taper_ytol", 10)
        lines.append("# ---- taper table")
        lines.append(f"taper = spec['tables'][roles['taperAndBuffer']]")
        lines.append(f"pdf_rows = extract_taper({cfg['taper']!r}, ytol={ytol}, has_shoulder={has_sh})")
        lines.append("assert_row_count(pdf_rows, 7, 'taper')")
        lines.append("for (spd, cells), js in zip(pdf_rows, taper['rows']):")
        lines.append("    eq('speed', spd, js['speedMph'])")
        lines.append("    exp = [f\"{js['longitudinalBufferSpace']['ft']}/{js['longitudinalBufferSpace']['skipLines']}\"]")
        lines.append("    for w_ in lw:")
        lines.append("        e = js['laneTaper'][w_]")
        lines.append("        exp.append(f\"{e['ft']}/{e['skipLines']}/{e['devices']}\")")
        if has_sh:
            lines.append("    for b in bands:")
            lines.append("        e = js['shoulderTaper'][b]")
            lines.append("        exp.append(f\"{e['ft']}/{e['skipLines']}/{e['devices']}\")")
        lines.append("    if cells != exp:")
        lines.append("        fails.append(f'taper {spd}: PDF={cells} JSON={exp}')")
        lines.append("print('taper rows', len(pdf_rows))")
    elif cfg.get("sibling_taper_of") == "311":
        lines.append("# ---- sibling identity: taper JSON == 311-02 (no PDF window due to rotation)")
        lines.append("ref = json.loads((ROOT / 'Data/sheet-specs/619-311.json').read_text(encoding='utf-8'))")
        lines.append("if 'taperAndBuffer' in roles:")
        lines.append("    t = spec['tables'][roles['taperAndBuffer']]")
        lines.append("    r311 = ref['tables']['311-02']")
        lines.append("    for a, b in zip(t['rows'], r311['rows']):")
        lines.append("        if a.get('laneTaper') != b.get('laneTaper') or a.get('longitudinalBufferSpace') != b.get('longitudinalBufferSpace'):")
        lines.append("            fails.append(f'taper identity vs 311 failed at {a.get(\"speedMph\")}')")
        lines.append("    print('taper identity vs 311 ok')")

    if cfg.get("roll"):
        lines.append("# ---- roll ahead")
        lines.append("roll = spec['tables'][roles['rollAheadDistance']]")
        lines.append(f"raw = group_rows(words_in_window(W, *{cfg['roll']!r}), y_tol=8.0)")
        lines.append("data = [r for r in raw if any(is_ratio(w[4]) for w in r)]")
        lines.append("assert_row_count(data, 3, 'roll')")
        lines.append("for r, js in zip(data, roll['rows']):")
        lines.append("    ratios = [w[4] for w in r if is_ratio(w[4])]")
        lines.append("    eq(f\"roll {js.get('speedBand')} min\", ratios[0], f\"{js['min']['ft']}/{js['min']['skipLines']}\")")
        lines.append("    eq(f\"roll {js.get('speedBand')} max\", ratios[1], f\"{js['max']['ft']}/{js['max']['skipLines']}\")")
        lines.append("print('roll rows', len(data))")

    if cfg.get("aw"):
        lines.append("# ---- advance warning")
        lines.append("aw = spec['tables'][roles['advanceWarningSpacing']]")
        lines.append(f"raw = group_rows(words_in_window(W, *{cfg['aw']!r}), y_tol=5.0)")
        lines.append("data = [r for r in raw if any(w[4].isdigit() and len(w[4]) == 3 for w in r)]")
        lines.append("assert_row_count(data, 4, 'aw')")
        lines.append("for r, js in zip(data, aw['rows']):")
        lines.append("    nums = [w[4] for w in r if w[4].isdigit() and len(w[4]) == 3]")
        lines.append("    eq('aw A', nums[0], js['A'])")
        lines.append("    eq('aw B', nums[1], js['B'])")
        if cfg.get("aw_has_c"):
            lines.append("    eq('aw C', nums[2], js['C'])")
        else:
            lines.append("    if 'C' in js:")
            lines.append("        fails.append('JSON has C but sheet is A/B only')")
        lines.append("print('aw rows', len(data))")

    return "\n".join(lines)


def main():
    for num, cfg in CONFIG.items():
        body = body_for(num, cfg)
        src = TEMPLATE.format(num=num, page=cfg["page"], body=body)
        path = OUT / f"619-{num}.py"
        path.write_text(src, encoding="utf-8")
        print("wrote", path.name)


if __name__ == "__main__":
    main()
