"""Round-trip check for Data/sheet-specs/619-414.json (Family 1).

Run: python Bridge/roundtrip/619-414.py
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

spec = json.loads((ROOT / "Data/sheet-specs/619-414.json").read_text(encoding="utf-8"))
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
if squash('CHANNELIZING DEVICE APPLICATION') not in squash(all_text):
    fails.append('phrase missing: ' + 'CHANNELIZING DEVICE APPLICATION')
if squash('20') not in squash(all_text):
    fails.append('phrase missing: ' + '20')
if squash('SEE NOTE 2') not in squash(all_text):
    fails.append('PV phrase missing: ' + 'SEE NOTE 2')
print('phrases ok')
# ---- sign size codes present in PDF text
need = ['G20-2', 'NYR9-11', 'NYW8-33', 'W4-2R', 'W20-1', 'W20-5R']
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
# ---- taper table
taper = spec['tables'][roles['taperAndBuffer']]
pdf_rows = extract_taper((100, 148, 780, 340), ytol=10, has_shoulder=True)
assert_row_count(pdf_rows, 7, 'taper')
for (spd, cells), js in zip(pdf_rows, taper['rows']):
    eq('speed', spd, js['speedMph'])
    exp = [f"{js['longitudinalBufferSpace']['ft']}/{js['longitudinalBufferSpace']['skipLines']}"]
    for w_ in lw:
        e = js['laneTaper'][w_]
        exp.append(f"{e['ft']}/{e['skipLines']}/{e['devices']}")
    for b in bands:
        e = js['shoulderTaper'][b]
        exp.append(f"{e['ft']}/{e['skipLines']}/{e['devices']}")
    if cells != exp:
        fails.append(f'taper {spd}: PDF={cells} JSON={exp}')
print('taper rows', len(pdf_rows))
# ---- roll ahead
roll = spec['tables'][roles['rollAheadDistance']]
raw = group_rows(words_in_window(W, *(90, 340, 450, 460)), y_tol=8.0)
data = [r for r in raw if any(is_ratio(w[4]) for w in r)]
assert_row_count(data, 3, 'roll')
for r, js in zip(data, roll['rows']):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    eq(f"roll {js.get('speedBand')} min", ratios[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq(f"roll {js.get('speedBand')} max", ratios[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print('roll rows', len(data))
# ---- advance warning
aw = spec['tables'][roles['advanceWarningSpacing']]
raw = group_rows(words_in_window(W, *(90, 40, 500, 140)), y_tol=5.0)
data = [r for r in raw if any(w[4].isdigit() and len(w[4]) == 3 for w in r)]
assert_row_count(data, 4, 'aw')
for r, js in zip(data, aw['rows']):
    nums = [w[4] for w in r if w[4].isdigit() and len(w[4]) == 3]
    eq('aw A', nums[0], js['A'])
    eq('aw B', nums[1], js['B'])
    eq('aw C', nums[2], js['C'])
print('aw rows', len(data))

print("fails:", len(fails))
for f in fails:
    print(" ", f)
sys.exit(1 if fails else 0)
