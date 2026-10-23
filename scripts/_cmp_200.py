import csv
from pathlib import Path

old_p = Path(r"c:\repos\microstation-vba-project\outputs\727-alignment-offset\dashed_sandwich_offset_20260902_212409.csv")
new_p = Path(r"c:\repos\microstation-vba-project\outputs\727-alignment-offset\dashed_sandwich_offset_20260903_061451.csv")


def load(p):
    with p.open(newline="", encoding="utf-8") as f:
        return {r["station"]: r for r in csv.DictReader(f)}


def sta_ft(sta: str) -> float:
    a, b = sta.replace(" ", "").split("+")
    return int(a) * 100 + float(b)


old, new = load(old_p), load(new_p)
rows = []
for sta, r in new.items():
    if sta_ft(sta) >= 20000:
        rows.append((sta_ft(sta), sta, r, old.get(sta)))
rows.sort()

print("station | old N/S | new N/S")
for _ft, sta, r, o in rows:
    on = (o or {}).get("north") or "-"
    os_ = (o or {}).get("south") or "-"
    nn = r["north"] or "-"
    ns = r["south"] or "-"
    flag = ""
    if on != nn or os_ != ns:
        flag = " *"
    print(f"{sta:8} old={on:>6}/{os_:>6}  new={nn:>6}/{ns:>6}{flag}")

nvals = [float(r["north"]) for _f, _s, r, _o in rows if r["north"]]
svals = [float(r["south"]) for _f, _s, r, _o in rows if r["south"]]
both = sum(1 for _f, _s, r, _o in rows if r["north"] and r["south"])
print(f"200+: both={both}/{len(rows)}")
print(f"200+: north n={len(nvals)} max={max(nvals) if nvals else None} mean={sum(nvals)/len(nvals):.3f}" if nvals else "no north")
print(f"200+: south n={len(svals)} max={max(svals) if svals else None} mean={sum(svals)/len(svals):.3f}" if svals else "no south")
