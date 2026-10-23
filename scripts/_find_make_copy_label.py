from pathlib import Path
import re

base = Path(r"C:\Program Files\Bentley\MicroStation 2023\MicroStation")
# Search for UTF-16 "Make Copy" and dump surrounding ascii/utf16 identifiers
needle = "Make Copy".encode("utf-16le")
count = 0
for f in base.rglob("*"):
    if f.suffix.lower() not in {".dll", ".ma", ".rsc", ".exe", ".xml"}:
        continue
    try:
        data = f.read_bytes()
    except Exception:
        continue
    if needle not in data and b"Make Copy" not in data:
        continue
    count += 1
    print("===", f)
    # ascii contexts
    for m in re.finditer(rb".{0,40}Make Copy.{0,40}", data):
        print(" A:", m.group().decode("latin1", errors="ignore"))
    # utf16 contexts decoded
    idx = 0
    while True:
        i = data.find(needle, idx)
        if i < 0:
            break
        chunk = data[max(0, i - 80) : i + 80]
        # extract utf16 strings nearby
        parts = re.findall(rb"(?:[\x20-\x7e]\x00){4,40}", chunk)
        for p in parts:
            print(" U:", p.decode("utf-16le", errors="ignore"))
        idx = i + 2
        if idx > i + 100000:
            break
    if count >= 25:
        break
print("files", count)
