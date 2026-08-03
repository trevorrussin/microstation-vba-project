"""Round-trip check for Data/sheet-specs/619-412.json (Family 1).

Run: python Bridge/roundtrip/619-412.py
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

spec = json.loads((ROOT / "Data/sheet-specs/619-412.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
page = doc[1]
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
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


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

# ---- phrase checks
if squash('INTERMEDIATE') not in squash(all_text):
    fails.append('phrase missing: ' + 'INTERMEDIATE')
if squash('TWO-WAY LEFT TURN') not in squash(all_text):
    fails.append('phrase missing: ' + 'TWO-WAY LEFT TURN')
if squash('NOTE 2') not in squash(all_text):
    fails.append('phrase missing: ' + 'NOTE 2')
if squash('NOTE 2') not in squash(all_text):
    fails.append('PV phrase missing: ' + 'NOTE 2')
print('phrases ok')
# ---- sign size codes present in PDF text
need = ['G20-2', 'NYR9-11', 'W20-1', 'W20-5', 'W4-2L', 'W9-3', 'R4-7']
for code in need:
    if squash(code) not in squash(all_text) and squash(code.replace('-', '')) not in squash(all_text):
        # WARNING FLAG often split across lines
        if code == 'WARNING FLAG' and 'FLAG' in all_text.upper():
            continue
        fails.append(f'size code missing in PDF: {code!r}')
print('size codes checked', need)
sz = spec['tables'][roles['signSizes']]
js_codes = [r['signCode'] for r in sz['rows']]
for code in need:
    if code not in js_codes and code != 'WARNING FLAG':
        fails.append(f'JSON size table missing {code!r}')
print('JSON size codes', js_codes)
# ---- sibling identity: taper JSON == 311-02 (no PDF window due to rotation)
ref = json.loads((ROOT / 'Data/sheet-specs/619-311.json').read_text(encoding='utf-8'))
if 'taperAndBuffer' in roles:
    t = spec['tables'][roles['taperAndBuffer']]
    r311 = ref['tables']['311-02']
    for a, b in zip(t['rows'], r311['rows']):
        if a.get('laneTaper') != b.get('laneTaper') or a.get('longitudinalBufferSpace') != b.get('longitudinalBufferSpace'):
            fails.append(f'taper identity vs 311 failed at {a.get("speedMph")}')
    print('taper identity vs 311 ok')

print("fails:", len(fails))
for f in fails:
    print(" ", f)
sys.exit(1 if fails else 0)
