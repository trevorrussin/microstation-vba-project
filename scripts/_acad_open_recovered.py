import time
import pythoncom
from win32com.client import GetObject

PATH = r"C:\Users\RussinT\OneDrive - AECOM\Documents\AutoCAD Sheet Sets\HWP24QC1_Location_01-02_50-52_SGD_RECOVERED_20260929_1705.dwg"


def main():
    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    acad.Visible = True
    try:
        print("docs", acad.Documents.Count)
        for i in range(acad.Documents.Count):
            print(" ", i, acad.Documents.Item(i).FullName)
    except Exception as e:
        print("list", e)

    # Already open?
    try:
        for i in range(acad.Documents.Count):
            d = acad.Documents.Item(i)
            if "RECOVERED_20260929_1705" in d.Name:
                d.Activate()
                print("ALREADY_OPEN", d.FullName)
                return
    except Exception:
        pass

    try:
        doc = acad.Documents.Open(PATH, False)  # ReadOnly=False
        print("OPENED", doc.FullName, "Saved", doc.Saved)
        return
    except Exception as e:
        print("Documents.Open:", e)

    try:
        ad = acad.ActiveDocument
        ad.SendCommand("FILEDIA 0 ")
        time.sleep(0.2)
        # double-backslash for AutoCAD command line? use forward slashes
        p = PATH.replace("\\", "/")
        ad.SendCommand(f'_.OPEN "{p}" ')
        time.sleep(3)
        print("after SendCommand active=", acad.ActiveDocument.FullName)
    except Exception as e:
        print("SendCommand open fail", e)


if __name__ == "__main__":
    main()
