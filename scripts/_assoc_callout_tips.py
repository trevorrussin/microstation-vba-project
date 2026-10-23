"""Associate rebuilt callout note-dim tips to nearest drawing elements via VBA."""
from __future__ import annotations

import math
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

# Precomputed from earlier scan (dim -> tip nearest keypoint). Re-validate quickly.
# (dimId, targetId, vertex, numer, denom)
SEED = [
    (5917, 4889, 1, 0, 1),
    (5922, 5840, 0, 0, 1),
    (5927, 3997, 1, 5, 8),
    (5932, 3399, 1, 3, 4),
    (5937, 5562, 1, 1, 2),
    (5942, 3997, 2, 3, 4),
    (5947, 4885, 1, 1, 2),
]


def refine(mr, tip, seed_tid: int):
    """Try keypoints on seed target + nearby lines; return best tuple."""
    candidates = []
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    while ee.MoveNext():
        e = ee.Current
        try:
            rng = e.Range
            dx = max(float(rng.Low.X) - tip.X, 0, tip.X - float(rng.High.X))
            dy = max(float(rng.Low.Y) - tip.Y, 0, tip.Y - float(rng.High.Y))
            bd = math.hypot(dx, dy)
            if bd > 25:
                continue
            if e.Type not in (3, 4, 14, 12, 16):
                continue
            # prioritize seed + close lines
            pri = 0 if e.ID64 == seed_tid else (1 if e.Type == 3 else 2)
            candidates.append((pri, bd, e.ID64, e))
        except Exception:
            pass
    candidates.sort(key=lambda t: (t[0], t[1], t[2]))
    best = None
    for pri, bd, eid, e in candidates[:12]:
        try:
            dp = e.GetDisplayPath()
            nvert = 2
            try:
                nvert = max(2, len(e.AsVertexList.GetVertices()))
            except Exception:
                pass
            for v in range(min(nvert, 6)):
                for den in (1, 2, 4, 8):
                    for num in range(0, den + 1):
                        try:
                            ap = dp.GetKeypointAssociation(v, num, den)
                            pt = mr.Point3dFromAssociationPoint(ap)
                            dist = math.hypot(pt.X - tip.X, pt.Y - tip.Y)
                            if best is None or dist < best[0]:
                                best = (dist, e.ID64, v, num, den)
                        except Exception:
                            pass
        except Exception:
            pass
        if best and best[0] < 0.05:
            break
    return best


def main() -> int:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference

    assocs = []
    for did, tid, v, num, den in SEED:
        tip = mr.GetElementByID64(did).AsDimensionElement.Point(1)
        b = refine(mr, tip, tid)
        if not b:
            print(f"dim {did}: FAIL no target")
            continue
        print(f"dim {did}: dist={b[0]:.3f} tgt={b[1]} key={b[2]},{b[3]}/{b[4]}")
        if b[0] < 30:
            assocs.append((did, b[1], b[2], b[3], b[4], b[0]))

    if not assocs:
        return 1

    code = "Option Explicit\r\n\r\nPublic Sub RunAssociateCalloutTips()\r\n    Dim r As String\r\n    Dim f As Integer\r\n    f = FreeFile\r\n    Open \"c:\\repos\\microstation-vba-project\\scripts\\_assoc_results.txt\" For Output As #f\r\n"
    for did, tid, v, num, den, dist in assocs:
        code += f"    r = AssociateOne({did}#, {tid}#, {v}, {num}, {den})\r\n"
        code += f'    Print #f, "dim {did}: " & r\r\n'
    code += """    Close #f
End Sub

Private Function AssociateOne(dimId As Double, targetId As Double, vertex As Long, numer As Long, denom As Long) As String
    On Error GoTo Fail
    Dim dimEl As Element, tgt As Element
    Set dimEl = FindById(dimId)
    Set tgt = FindById(targetId)
    If dimEl Is Nothing Then AssociateOne = "ERROR dim missing": Exit Function
    If tgt Is Nothing Then AssociateOne = "ERROR tgt missing": Exit Function
    If Not dimEl.IsDimensionElement Then AssociateOne = "ERROR not dim": Exit Function
    Dim oDim As DimensionElement
    Set oDim = dimEl.AsDimensionElement
    Dim stub As Point3d
    stub = oDim.Point(2)
    Dim oPath As DisplayPath
    Set oPath = tgt.GetDisplayPath
    Dim oAssoc As AssociationPoint
    oAssoc = oPath.GetKeypointAssociation(vertex, numer, denom)
    Do While oDim.ReferencePointsCount > 0
        oDim.DeletePoint 1
    Loop
    oDim.AddReferencePoint ActiveModelReference, oAssoc
    oDim.AddReferencePoint ActiveModelReference, stub
    oDim.Rewrite
    AssociateOne = "OK pts=" & oDim.ReferencePointsCount & " tgt=" & CStr(targetId)
    Exit Function
Fail:
    AssociateOne = "ERROR " & Err.Description
End Function

Private Function FindById(elementId As Double) As Element
    Dim oScan As New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        On Error Resume Next
        If CDbl(el.ID.Low) = elementId And el.ID.High = 0 Then
            Set FindById = el
            Exit Function
        End If
        On Error GoTo 0
    Loop
    Set FindById = Nothing
End Function
"""

    vbe = app.VBE
    proj = None
    for i in range(1, vbe.VBProjects.Count + 1):
        p = vbe.VBProjects.Item(i)
        if p.Name == "OpenCloseEventHandler" and p.Protection == 0:
            proj = p
            break
    if proj is None:
        print("ERROR: no writable VBA project")
        return 1

    try:
        proj.VBComponents.Remove(proj.VBComponents.Item("CursorAssocTip"))
    except Exception:
        pass
    # Remove leftover ModuleN from prior failed rename
    for i in range(proj.VBComponents.Count, 0, -1):
        c = proj.VBComponents.Item(i)
        if c.Name.startswith("Module") and c.Type == 1:
            try:
                cm = c.CodeModule
                if cm.CountOfLines > 0 and "RunAssociateCalloutTips" in cm.Lines(1, min(cm.CountOfLines, 5)):
                    proj.VBComponents.Remove(c)
            except Exception:
                pass

    comp = proj.VBComponents.Add(1)  # std module; name assigned by VBE
    mod_name = comp.Name
    comp.CodeModule.AddFromString(code)
    print("injected", mod_name, comp.CodeModule.CountOfLines, "lines,", len(assocs), "dims")

    app.CadInputQueue.SendKeyin(
        f"VBA RUN [OpenCloseEventHandler]{mod_name}.RunAssociateCalloutTips"
    )
    print("VBA RUN sent", mod_name)
    time.sleep(4)

    for did, tid, v, num, den, dist in assocs:
        d = mr.GetElementByID64(did).AsDimensionElement
        print(
            f"verify {did} pts={d.ReferencePointsCount} "
            f"p1=({d.Point(1).X:.2f},{d.Point(1).Y:.2f}) "
            f"p2=({d.Point(2).X:.2f},{d.Point(2).Y:.2f})"
        )

    app.ActiveSettings.AssociationLockEnabled = True
    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    app.CadInputQueue.SendKeyin("REDRAW ALL")
    print("SAVE+REDRAW")

    try:
        proj.VBComponents.Remove(proj.VBComponents.Item(mod_name))
        print("cleaned", mod_name)
    except Exception as e:
        print("cleanup", e)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
