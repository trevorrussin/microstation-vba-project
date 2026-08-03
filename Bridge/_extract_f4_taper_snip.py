"""Extract taper table speed rows for Family 4 recon."""
import fitz, re
from collections import defaultdict
from pathlib import Path

CAP = Path(__file__).resolve().parents[1] / "Bridge" / "captures"


def group_rows(words, y_tol=4.0):
    rows = defaultdict(list)
    for w in words:
        rows[round(w[1] / y_tol)].append(w)
    return [sorted(rows[k], key=lambda w: w[0]) for k in sorted(rows)]


def extract_taper_table(sheet, table_id):
    doc = fitz.open(CAP / f"{sheet}.pdf")
    text = "\n".join(p.get_text("text") for p in doc)
    words = []
    for p in doc:
        words.extend(p.get_text("words"))
    doc.close()
    # locate TABLE title
    pat = re.compile(rf"TABLE\s+{re.escape(table_id)}", re.I)
    tw = next((w for w in words if pat.match(w[4]) or (w[4].upper() == "TABLE" and any(
        abs(w2[1]-w[1])<4 and table_id.split("-")[1] in w2[4] for w2 in words
    ))), None)
    if not tw:
        # search all pages by text position
        m = pat.search(text)
        if not m:
            return {"error": "title not found"}
    # find y of table title from words containing table id
    title_ws = [w for w in words if table_id.replace("-", "") in w[4].replace("-", "") or w[4].upper() == table_id.split("-")[1]]
    if not title_ws:
        title_ws = [w for w in words if w[4].upper() == "TABLE"]
    y0 = min(w[1] for w in title_ws) if title_ws else 300
    region = [w for w in words if y0 <= w[1] <= y0 + 220]
    rows = group_rows(region, 6.0)
    row_texts = [" ".join(w[4] for w in r) for r in rows if len(" ".join(w[4] for w in r)) > 3]
    speeds = sorted(set(re.findall(r"\b(25|30|35|40|45|50|55|60|65)\b", "\n".join(row_texts))))
    return {"table": table_id, "rowCount": len(row_texts), "speeds": speeds, "firstRows": row_texts[:8], "lastRows": row_texts[-3:]}


def all_tables(sheet):
    doc = fitz.open(CAP / f"{sheet}.pdf")
    text = "\n".join(p.get_text("text") for p in doc)
    doc.close()
    return [(m.group(1), m.group(2).strip()) for m in re.finditer(r"TABLE\s+(\d{3}-\d{2})\s*:\s*([^\n]+)", text, re.I)]


if __name__ == "__main__":
    for sheet in ["619-306", "619-212", "619-304", "619-206"]:
        print("===", sheet, "tables:", all_tables(sheet))
        tid = {"619-306": "306-03", "619-212": "212-03", "619-304": "304-03", "619-206": "206-04"}.get(sheet)
        if tid and any(t[0] == tid for t in all_tables(sheet)):
            print(extract_taper_table(sheet, tid))
        print()
