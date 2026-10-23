Attribute VB_Name = "AssociateCalloutTips"
' Snap rebuilt note-dim tips to nearest keypoint XYZ (Point3dFromAssociationPoint).
' NOTE: MicroStation CONNECT VBA AddReferencePoint only accepts Point3d — not
' AssociationPoint — so tips sit on the feature but are not true associations
' (whole callout still translates as a unit). F5 RunAssociateCalloutTips.

Option Explicit

Public Sub RunAssociateCalloutTips()
    Dim r As String
    Dim f As Integer
    f = FreeFile
    Open "c:\repos\microstation-vba-project\scripts\_assoc_results.txt" For Output As #f
    r = SnapTip(5917#, 4889#, 1, 0, 1)
    Print #f, "dim 5917: " & r
    r = SnapTip(5922#, 5840#, 0, 0, 1)
    Print #f, "dim 5922: " & r
    r = SnapTip(5927#, 3997#, 1, 5, 8)
    Print #f, "dim 5927: " & r
    r = SnapTip(5932#, 3399#, 1, 3, 4)
    Print #f, "dim 5932: " & r
    r = SnapTip(5937#, 5562#, 1, 1, 2)
    Print #f, "dim 5937: " & r
    r = SnapTip(5942#, 5844#, 0, 0, 1)
    Print #f, "dim 5942: " & r
    r = SnapTip(5947#, 4885#, 1, 1, 2)
    Print #f, "dim 5947: " & r
    Close #f
    ActiveDesignFile.Save
    Debug.Print "AssociateCalloutTips DONE"
End Sub

Private Function SnapTip(ByVal dimId As Double, ByVal targetId As Double, _
                         ByVal vertex As Long, ByVal numer As Long, ByVal denom As Long) As String
    On Error GoTo Fail
    Dim dimEl As Element, tgt As Element
    Set dimEl = FindById(dimId)
    Set tgt = FindById(targetId)
    If dimEl Is Nothing Then SnapTip = "ERROR dim missing": Exit Function
    If tgt Is Nothing Then SnapTip = "ERROR tgt missing": Exit Function
    If Not dimEl.IsDimensionElement Then SnapTip = "ERROR not dim": Exit Function

    Dim oDim As DimensionElement
    Set oDim = dimEl.AsDimensionElement
    Dim stub As Point3d
    stub = oDim.Point(2)

    Dim oPath As DisplayPath
    Set oPath = tgt.GetDisplayPath
    Dim oAssoc As AssociationPoint
    oAssoc = oPath.GetKeypointAssociation(vertex, numer, denom)

    ' CONNECT typelib: AddReferencePoint(Model, Point3d) only — AssociationPoint
    ' compiles as ByRef mismatch. Convert to free XYZ on the keypoint.
    Dim tip As Point3d
    tip = ActiveModelReference.Point3dFromAssociationPoint(oAssoc)

    Do While oDim.ReferencePointsCount > 0
        oDim.DeletePoint 1
    Loop
    oDim.AddReferencePoint ActiveModelReference, tip
    oDim.AddReferencePoint ActiveModelReference, stub
    oDim.Rewrite
    SnapTip = "OK pts=" & oDim.ReferencePointsCount & " tgt=" & CStr(targetId) & " (xyz snap, not assoc)"
    Exit Function
Fail:
    SnapTip = "ERROR " & Err.Description
End Function

Private Function FindById(ByVal elementId As Double) As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Dim el As Element
    Dim id As DLong
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        id = el.ID
        If ElIDAsDouble(id) = elementId Then
            Set FindById = el
            Exit Function
        End If
    Loop
    Set FindById = Nothing
End Function
