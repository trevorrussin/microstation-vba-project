import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
adf = app.ActiveDesignFile
xs = None
for m in adf.Models:
    if m.Name == "XS Stacked Views":
        xs = m
        break
print("XS", xs.Name)
try:
    print("XS atts", xs.Attachments.Count)
    for i in range(1, xs.Attachments.Count + 1):
        a = xs.Attachments.Item(i)
        print(" ", i, a.Name, getattr(a, "LogicalName", None), getattr(a, "Description", None))
        try:
            print("   file", a.DesignFile.FullName)
        except Exception as e:
            print("   file err", e)
except Exception as e:
    print("atts err", e)

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = xs.Scan(o)
conduit = []
cells = []
dims = []
other = []
while ee.MoveNext():
    e = ee.Current
    try:
        if e.Type == 2:
            try:
                cells.append((int(e.ID64), e.AsCellElement.Name))
            except Exception:
                pass
        if e.Type == 33:
            dims.append(int(e.ID64))
        txt = ""
        if e.IsTextElement:
            txt = e.AsTextElement.Text
        elif e.IsTextNodeElement:
            tn = e.AsTextNodeElement
            txt = " | ".join(tn.TextLine(i) for i in range(1, int(tn.TextLinesCount) + 1))
        u = (txt or "").upper()
        if "CONDUIT" in u or "SIGNAL CABLE" in u or "PULL" in u:
            r = e.Range
            conduit.append(
                (
                    int(e.ID64),
                    int(e.Type),
                    e.Level.Name if e.Level else "?",
                    round(float(r.Low.X), 2),
                    round(float(r.Low.Y), 2),
                    txt[:140],
                )
            )
        if any(
            k in u
            for k in (
                "680.818",
                "680.5001",
                "680.822",
                "680.8308",
                "VEHICULAR",
                "APS",
                "STREETLIGHT",
                "LUMINAIRE",
            )
        ):
            other.append((int(e.ID64), int(e.Type), e.Level.Name if e.Level else "?", txt[:120]))
    except Exception:
        continue

print("conduit/cable texts", len(conduit))
for c in conduit:
    print(c)
print("cells n", len(cells), "sample", cells[:15])
print("dims n", len(dims), "sample", dims[:20])
print("other related", len(other))
for row in other[:50]:
    print(row)
