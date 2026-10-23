import pythoncom
from win32com.client import GetObject

OUTERS = {"13D80A", "13D25C", "13D7DC", "13CEC7", "13D2EA", "13D7BC"}


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    ms = doc.ModelSpace
    n = ms.Count
    print("ModelSpace", n)
    hits = []
    for i in range(n):
        try:
            e = ms.Item(i)
            if e.Layer != "C-TOPO-FEAT":
                continue
            h = e.Handle
            if h in OUTERS:
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if 1028440 < cx < 1028620 and 187390 < cy < 187530:
                hits.append((h, e.ObjectName, cx, cy))
        except Exception:
            continue
    print("inner-ish C-TOPO-FEAT", len(hits))
    for row in sorted(hits, key=lambda r: r[2]):
        print(row)

    steel = []
    for lyr in (
        "Straight Steel Faced Concrete Curb (18'' Deep)",
        "Depressed Steel Faced Concrete Curb (18'' Deep)",
        "Corner Steel Faced Concrete Curb (18'' Deep)",
    ):
        for i in range(n):
            try:
                e = ms.Item(i)
                if e.Layer != lyr:
                    continue
                mn, mx = e.GetBoundingBox()
                cx = (float(mn[0]) + float(mx[0])) / 2
                cy = (float(mn[1]) + float(mx[1])) / 2
                if 1028440 < cx < 1028620 and 187390 < cy < 187530:
                    steel.append((lyr, e.Handle, e.ObjectName))
            except Exception:
                continue
    print("steel segs in box", len(steel))


if __name__ == "__main__":
    main()
