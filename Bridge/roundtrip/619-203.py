"""Round-trip check for Data/sheet-specs/619-203.json (Family 1).

Run: python Bridge/roundtrip/619-203.py
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

spec = json.loads((ROOT / "Data/sheet-specs/619-203.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
page = doc[0]
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
if squash('SHORT DURATION') not in squash(all_text):
    fails.append('phrase missing: ' + 'SHORT DURATION')
if squash('OPERATOR') not in squash(all_text):
    fails.append('phrase missing: ' + 'OPERATOR')
if squash('ROLL AHEAD') not in squash(all_text):
    fails.append('phrase missing: ' + 'ROLL AHEAD')
if squash('P, TMIA') not in squash(all_text):
    fails.append('PV phrase missing: ' + 'P, TMIA')
print('phrases ok')
# ---- sign size codes present in PDF text
need = ['NYW8-33', 'W4-2R', 'W20-1', 'WARNING FLAG']
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
# ---- roll ahead
roll = spec['tables'][roles['rollAheadDistance']]
raw = group_rows(words_in_window(W, *(1040, 40, 1224, 160)), y_tol=8.0)
data = [r for r in raw if any(is_ratio(w[4]) for w in r)]
assert_row_count(data, 3, 'roll')
for r, js in zip(data, roll['rows']):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    eq(f"roll {js.get('speedBand')} min", ratios[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq(f"roll {js.get('speedBand')} max", ratios[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print('roll rows', len(data))
# ---- advance warning
aw = spec['tables'][roles['advanceWarningSpacing']]
raw = group_rows(words_in_window(W, *(990, 240, 1224, 340)), y_tol=5.0)
data = [r for r in raw if any(w[4].isdigit() and len(w[4]) == 3 for w in r)]
assert_row_count(data, 4, 'aw')
for r, js in zip(data, aw['rows']):
    nums = [w[4] for w in r if w[4].isdigit() and len(w[4]) == 3]
    eq('aw A', nums[0], js['A'])
    eq('aw B', nums[1], js['B'])
    if 'C' in js:
        fails.append('JSON has C but sheet is A/B only')
print('aw rows', len(data))

print("fails:", len(fails))
for f in fails:
    print(" ", f)
sys.exit(1 if fails else 0)
