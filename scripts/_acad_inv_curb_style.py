"""Read-only: per SGD, count C-TOPO-FEAT feature lines vs PRO-S curb polylines vs ramps.

Classifies sheet curb style: 'islandFL' (Loc01-like C-TOPO-FEAT) vs 'proS_only'.
Opens read-only, never saves; skips docs already open. Output: outputs/curb-style-inventory.json
"""
import json, os, sys, time
import pythoncom
from pywintypes import com_error
from win32com.client import GetObject, VARIANT

SHOTS = r"C:\repos\microstation-vba-project\Screenshots"
OUT = r"C:\repos\microstation-vba-project\outputs\curb-style-inventory.json"


def R(fn):
    for _ in range(150):
        try:
            return fn()
        except com_error as e:
            if e.args and e.args[0] == -2147418111:
                time.sleep(0.4); continue
            raise
    return fn()


def count(doc, layer, types):
    try: R(lambda: doc.SelectionSets.Item("INVS").Delete())
    except Exception: pass
    ss = R(lambda: doc.SelectionSets.Add("INVS"))
    R(lambda: ss.Select(5, None, None, VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8]),
                        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [layer])))
    c = {}
    for i in range(R(lambda: ss.Count)):
        on = R(lambda i=i: ss.Item(i).ObjectName)
        if on in types: c[on] = c.get(on, 0) + 1
    R(lambda: ss.Delete())
    return c


pythoncom.CoInitialize()
a = GetObject(None, "AutoCAD.Application")
names = sorted(n for n in os.listdir(SHOTS) if n.upper().endswith("_SGD.DWG"))
res = {}
for n in names:
    p = os.path.join(SHOTS, n)
    if any(R(lambda i=i: a.Documents.Item(i).Name).lower() == n.lower() for i in range(R(lambda: a.Documents.Count))):
        res[n] = "open-skipped"; continue
    doc = R(lambda: a.Documents.Open(p, True))
    try:
        res[n] = {
            "topoFL": count(doc, "C-TOPO-FEAT", ("AeccDbFeatureLine", "AcDb3dPolyline")),
            "proS": count(doc, "PRO-S-CURB-LINE", ("AcDbPolyline", "AcDbBlockReference")),
            "propD": count(doc, "PROP-D-ROAD-RAMPS", ("AcDbBlockReference",)),
        }
    finally:
        R(lambda: doc.Close(False))
    print(n, res[n], flush=True)
    json.dump(res, open(OUT, "w"), indent=1)
