from pathlib import Path
import re

base = Path(r"C:\Program Files\Bentley\MicroStation 2023\MicroStation")
files = (
    list(base.rglob("Bentley.DgnPlatformNET.dll"))
    + list(base.rglob("Bentley.DgnDisplayNet.dll"))
    + list(base.rglob("DgnToolsNet2.dll"))
    + list(base.rglob("ParametricCellManipulator.dll"))
    + list(base.rglob("*Manip*.dll"))
)

seen = set()
for f in files:
    if f in seen:
        continue
    seen.add(f)
    data = f.read_bytes()
    # ASCII strings containing MakeCopy / make copy
    ascii_hits = set(re.findall(rb"[\x20-\x7e]{0,60}[Mm]ake[Cc]opy[\x20-\x7e]{0,60}", data))
    ascii_hits |= set(re.findall(rb"[\x20-\x7e]{0,40}MAKECOPY[\x20-\x7e]{0,40}", data))
    ascii_hits |= set(re.findall(rb"[\x20-\x7e]{0,40}makeCopy[\x20-\x7e]{0,40}", data))
    # UTF-16LE
    u16 = set()
    for m in re.finditer(rb"(?:[\x20-\x7e]\x00){3,50}[Mm]\x00a\x00k\x00e\x00[Cc]\x00o\x00p\x00y(?:[\x20-\x7e]\x00){0,40}", data):
        u16.add(m.group().decode("utf-16le", errors="ignore"))
    if not ascii_hits and not u16:
        continue
    print("===", f)
    for h in sorted(ascii_hits)[:40]:
        print(" A:", h.decode("ascii", errors="ignore"))
    for h in sorted(u16)[:40]:
        print(" U:", h)
