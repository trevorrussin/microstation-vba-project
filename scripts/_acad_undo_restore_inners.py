"""Undo until Loc01 INNER feature lines return. Never opens another DWG."""
from __future__ import annotations

import time
import pythoncom
from win32com.client import GetObject

TARGETS = [
    "13D822", "13D824", "13D425", "13D42D", "13D2EE",
    "13D7C4", "13D7CC", "13D7BF", "13CED4", "13CED2",
    "13D263", "13D265", "13D25E", "13D282",
    "13D7DB", "13D7E0", "13D7E8",
]
OUTERS = ["13D80A", "13D25C", "13D7DC", "13CEC7", "13D2EA", "13D7BC"]


def alive(doc, hs):
    ok = []
    for h in hs:
        try:
            doc.HandleToObject(h)
            ok.append(h)
        except Exception:
            pass
    return ok


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("file", doc.Name, "Saved", doc.Saved)
    try:
        doc.SendCommand(chr(27) + chr(27))
        time.sleep(0.2)
    except Exception:
        pass
    before = alive(doc, TARGETS)
    print("inners before", len(before), before)
    print("outers before", alive(doc, OUTERS))
    # Prefer AutoCAD UNDO via SendCommand — one step at a time
    for i in range(40):
        have = alive(doc, TARGETS)
        if len(have) >= 12:
            print(f"STOP after {i} undos — inners {len(have)}")
            break
        try:
            doc.SendCommand("_U ")
        except Exception as ex:
            print("undo fail", ex)
            break
        time.sleep(0.35)
        have2 = alive(doc, TARGETS)
        print(f"undo {i+1}: inners {len(have2)} outers {len(alive(doc, OUTERS))} ms={doc.ModelSpace.Count}")
        if len(have2) < len(have) and i > 5:
            # undoing past restore — stop
            print("inners shrinking — stop")
            break
    final = alive(doc, TARGETS)
    print("FINAL inners", len(final), final)
    print("FINAL outers", alive(doc, OUTERS))
    try:
        e = doc.HandleToObject("13D986")
        print("13D986 still", e.ObjectName)
    except Exception:
        print("13D986 MISSING")


if __name__ == "__main__":
    main()
