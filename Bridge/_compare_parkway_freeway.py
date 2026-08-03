"""Compare 619-306 Parkway vs freeway siblings."""
import fitz, re, json
from pathlib import Path

CAP = Path(__file__).resolve().parents[1] / "Bridge" / "captures"

def snap(sheet):
    doc = fitz.open(CAP / f"{sheet}.pdf")
    text = doc[0].get_text("text")
    doc.close()
    tables = [(m.group(1), m.group(2).strip()) for m in re.finditer(r"TABLE\s+(\d{3}-\d{2})\s*:\s*([^\n]+)", text, re.I)]
    seen = set(); tdedup = []
    for a,b in tables:
        if a not in seen:
            seen.add(a); tdedup.append((a,b))
    return {
        "sheet": sheet,
        "tables": tdedup,
        "signs": sorted(set(re.findall(r"\b(?:NY[A-Z0-9][-\w]*|W\d{1,2}[-\w]*|R\d[-\w]*|G20[-\w]*)\b", text, re.I))),
        "spacing": sorted(set(re.findall(r"(?<!\d)(1000|1320|1500|2640|500)(?:'| FT)?", text)), key=int),
        "features": {k: k in text.upper() for k in [
            "MERGING TAPER","SHOULDER TAPER","DOWNSTREAM TAPER","ADVANCE WARNING","SIGN SPACING",
            "PARKWAY","FREEWAY","SHORT TERM","SHORT DURATION","MOBILE","MOVING OPERATION","SHOULDER < 8"
        ]},
        "operation": next((l.strip() for l in text.splitlines() if "OPERATION" in l.upper() and len(l)<60), ""),
    }

pairs = ["619-306","619-304","619-212","619-206","619-114","619-111","619-041","619-031"]
out = {s: snap(s) for s in pairs if (CAP/f"{s}.pdf").exists()}
print(json.dumps(out, indent=2))
