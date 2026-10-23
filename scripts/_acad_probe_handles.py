import pythoncom
from win32com.client import GetObject

INNERS = [
    "13D822", "13D824", "13D425", "13D42D", "13D2EE", "13D7C4", "13D7CC", "13D7BF",
    "13CED4", "13CED2", "13D263", "13D265", "13D25E", "13D282", "13D7DB", "13D7E0", "13D7E8",
]

def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    for h in INNERS:
        try:
            e = doc.HandleToObject(h)
            print(h, e.ObjectName, e.Layer)
        except Exception as ex:
            print(h, "MISSING", ex)

if __name__ == "__main__":
    main()
