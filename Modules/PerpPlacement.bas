
Option Explicit

' ============================================================
' ALIGNMENT PLACEMENT MODULE
' ------------------------------------------------------------
' Walks along the drawn alignment and places perpendicular
' tick-lines at each WZTC order item location, separated by
' the spacings configured in the Workzone Designer form.
'
' Flow:
'   1. StartAlignmentPlacement() builds the path from the
'      elements drawn after the snapshot ID.
'   2. PlacePerp (modeless) calls
'      PlaceLineForCurrentItem() or SkipCurrentItem()
'      for each item in sequence.
'   3. Each perpendicular line is 40 ft long (20 ft each side).
' ============================================================

' ---- Internal path segment type ----
Private Type PathSeg
    IsArc   As Boolean
    SX As Double: SY As Double: SZ As Double   ' start point
    EX As Double: EY As Double: EZ As Double   ' end point
    SegLen  As Double
    ' Arc-only properties (all in design-file master units / radians)
    CX As Double: CY As Double: CZ As Double   ' centre
    Radius      As Double
    StartAngle  As Double   ' radians, standard maths convention
    SweepAngle  As Double   ' radians, positive=CCW, negative=CW
End Type

' ---- Module-level path state ----
Private pathSegs()     As PathSeg
Private pathSegCount   As Integer
Private totalPathLen   As Double

' ---- Placement progress (read by PlacePerp) ----
Public currentItemIdx  As Integer   ' 0-based index into wztcOrderLabels
Public currentPathPos  As Double    ' cumulative distance from alignment start (ft)

' Default half-length of each perpendicular tick line (master units = ft)
Private Const PERP_HALF_LEN As Double = 20

' ============================================================
' MAIN ENTRY POINT
' ============================================================
Public Sub StartAlignmentPlacement()
    If Not BuildAlignmentPath() Then
        MsgBox "Could not build alignment path." & vbCrLf & _
               "Make sure you drew the alignment before clicking Done.", _
               vbExclamation, "Alignment Placement"
        Exit Sub
    End If

    If wztcOrderLabelCount <= 0 Then
        MsgBox "No WZTC order items found." & vbCrLf & _
               "Please run Workzone Designer first.", _
               vbExclamation, "Alignment Placement"
        Exit Sub
    End If

    ' Reset progress
    currentItemIdx = 0
    currentPathPos = 0
    wztcPlacedSignCount = 0

    ' Show the placement form
    PlacePerp.Show vbModeless
End Sub

' ============================================================
' BUILD ALIGNMENT PATH
' Collects line/arc elements newer than wztcAlignmentStartMaxID,
' sorts them into a connected chain starting at the first
' recorded click point, and fills pathSegs().
' Returns True on success.
' ============================================================
Private Function BuildAlignmentPath() As Boolean
    On Error GoTo BuildErr

    pathSegCount = 0
    totalPathLen = 0
    ReDim pathSegs(0)

    ' --- collect candidate elements ---
    Dim elems()  As Element
    Dim nElems   As Integer
    nElems = 0

    Dim el As Element
    Dim oEnum As ElementEnumerator
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If ElIDAsDouble(el.ID) > wztcAlignmentStartMaxID Then
            If el.Type = msdElementTypeLine Or el.Type = msdElementTypeArc Then
                nElems = nElems + 1
                ReDim Preserve elems(1 To nElems)
                Set elems(nElems) = el
            End If
        End If
    Loop

    If nElems = 0 Then
        BuildAlignmentPath = False
        Exit Function
    End If

    ' --- sort elements by element ID ascending (= drawing order = path order) ---
    ' Elements are added to the drawing in the order they are drawn, so their
    ' IDs are monotonically increasing along the path.
    Dim elemIDs() As Double
    ReDim elemIDs(1 To nElems)
    Dim i As Integer, j As Integer
    For i = 1 To nElems
        elemIDs(i) = ElIDAsDouble(elems(i).ID)
    Next i
    Dim tmpEl As Element, tmpID As Double
    For i = 2 To nElems
        tmpID = elemIDs(i)
        Set tmpEl = elems(i)
        j = i - 1
        Do While j >= 1 And elemIDs(j) > tmpID
            elemIDs(j + 1) = elemIDs(j)
            Set elems(j + 1) = elems(j)
            j = j - 1
        Loop
        elemIDs(j + 1) = tmpID
        Set elems(j + 1) = tmpEl
    Next i

    ' --- build path in drawing order ---
    '
    ' wztcAlignmentFirstPoint* is the first mouse-click when drawing started.
    ' It is exactly the path-direction start of the first element.
    '
    ' Lines : le.StartPoint / EndPoint — pick the end closest to chainPt.
    ' Arcs  : The center is obtained via ae.CenterPoint (MicroStation 2023).
    '         If that property is not available, both possible centers are
    '         computed and validated against ae.Range (works in all versions).
    '         The chain direction is then determined by checking which
    '         geometric endpoint (sa or sa+sw) is closer to chainPt.
    Dim chainX As Double, chainY As Double, chainZ As Double
    chainX = wztcAlignmentFirstPointX
    chainY = wztcAlignmentFirstPointY
    chainZ = wztcAlignmentFirstPointZ

    ReDim pathSegs(1 To nElems)

    For i = 1 To nElems
        Dim e As Element
        Set e = elems(i)

        Dim seg As PathSeg

        If e.Type = msdElementTypeLine Then
            Dim le As LineElement
            Set le = e
            Dim sp As Point3d, ep As Point3d
            sp = le.startPoint
            ep = le.endPoint

            ' Pick orientation so the start matches the current chain end
            Dim dSP As Double, dEP As Double
            dSP = (sp.X - chainX) ^ 2 + (sp.Y - chainY) ^ 2
            dEP = (ep.X - chainX) ^ 2 + (ep.Y - chainY) ^ 2

            seg.IsArc = False
            If dSP <= dEP Then
                seg.SX = sp.X:  seg.SY = sp.Y:  seg.SZ = sp.Z
                seg.EX = ep.X:  seg.EY = ep.Y:  seg.EZ = ep.Z
            Else
                seg.SX = ep.X:  seg.SY = ep.Y:  seg.SZ = ep.Z
                seg.EX = sp.X:  seg.EY = sp.Y:  seg.EZ = sp.Z
            End If

            Dim ldx As Double, ldy As Double
            ldx = seg.EX - seg.SX:  ldy = seg.EY - seg.SY
            seg.SegLen = Sqr(ldx * ldx + ldy * ldy)

            chainX = seg.EX:  chainY = seg.EY:  chainZ = seg.EZ

        ElseIf e.Type = msdElementTypeArc Then
            Dim ae As ArcElement
            Set ae = e
            Dim R As Double, sa As Double, sw As Double
            R = ae.PrimaryRadius
            sa = ae.StartAngle
            sw = ae.SweepAngle

            ' --- Determine arc center ---
            ' Try ae.CenterPoint first (MicroStation 2023 / CONNECT edition).
            ' If not available, derive mathematically with Range validation.
            ' If angles appear to be in degrees, convert to radians and retry.
            Dim ctrX As Double, ctrY As Double
            Dim gotCenter As Boolean
            gotCenter = False

            Debug.Print "Arc seg " & i & ": R=" & R & " sa=" & sa & " sw=" & sw
            Debug.Print "  chainPt: " & chainX & ", " & chainY

            On Error Resume Next
            Dim ctrPt As Point3d
            ctrPt = ae.CenterPoint
            If Err.Number = 0 Then
                ctrX = ctrPt.X
                ctrY = ctrPt.Y
                gotCenter = True
                Debug.Print "  CenterPoint OK: " & ctrX & ", " & ctrY
            End If
            Err.Clear
            On Error GoTo BuildErr

            Dim rng As Range3d
            rng = ae.Range
            Dim tol As Double
            tol = R * 0.01 + 1

            If Not gotCenter Then
                ' Fallback 1: try angles as radians
                gotCenter = TryComputeCenter(chainX, chainY, R, sa, sw, rng, tol, ctrX, ctrY)
                If gotCenter Then
                    Debug.Print "  Center (radians): " & ctrX & ", " & ctrY
                End If
            End If

            If Not gotCenter Then
                ' Fallback 2: try angles as degrees (convert to radians)
                Dim PI As Double
                PI = 3.14159265358979
                Dim saRad As Double, swRad As Double
                saRad = sa * PI / 180#
                swRad = sw * PI / 180#
                gotCenter = TryComputeCenter(chainX, chainY, R, saRad, swRad, rng, tol, ctrX, ctrY)
                If gotCenter Then
                    ' Angles were in degrees — use converted values
                    sa = saRad
                    sw = swRad
                    Debug.Print "  Center (degrees->rad): " & ctrX & ", " & ctrY
                End If
            End If

            If Not gotCenter Then
                ' Last resort: use bounding box center, adjusted to be R from chainPt
                Dim bx As Double, by As Double
                bx = (rng.Low.X + rng.High.X) / 2
                by = (rng.Low.Y + rng.High.Y) / 2
                ' Project to be exactly R from chainPt
                Dim bd As Double
                bd = Sqr((bx - chainX) ^ 2 + (by - chainY) ^ 2)
                If bd > 0.001 Then
                    ctrX = chainX + (bx - chainX) * R / bd
                    ctrY = chainY + (by - chainY) * R / bd
                Else
                    ctrX = chainX + R
                    ctrY = chainY
                End If
                gotCenter = True
                Debug.Print "  Center (bbox fallback): " & ctrX & ", " & ctrY
            End If

            ' --- Determine path direction through the arc ---
            ' Compute both geometric endpoints from the center.
            Dim geoStartX As Double, geoStartY As Double
            Dim geoEndX As Double, geoEndY As Double
            geoStartX = ctrX + R * Cos(sa)
            geoStartY = ctrY + R * Sin(sa)
            geoEndX = ctrX + R * Cos(sa + sw)
            geoEndY = ctrY + R * Sin(sa + sw)

            ' Check which geometric endpoint is closer to the chain point
            Dim dGeoStart As Double, dGeoEnd As Double
            dGeoStart = (geoStartX - chainX) ^ 2 + (geoStartY - chainY) ^ 2
            dGeoEnd = (geoEndX - chainX) ^ 2 + (geoEndY - chainY) ^ 2

            seg.IsArc = True
            seg.CX = ctrX:  seg.CY = ctrY:  seg.CZ = chainZ
            seg.Radius = R

            If dGeoStart <= dGeoEnd Then
                ' Chain enters at geometric start — travel in sweep direction
                seg.StartAngle = sa
                seg.SweepAngle = sw
                seg.SX = geoStartX:  seg.SY = geoStartY:  seg.SZ = chainZ
                seg.EX = geoEndX:    seg.EY = geoEndY:    seg.EZ = chainZ
            Else
                ' Chain enters at geometric end — travel in reverse sweep direction
                seg.StartAngle = sa + sw
                seg.SweepAngle = -sw
                seg.SX = geoEndX:    seg.SY = geoEndY:    seg.SZ = chainZ
                seg.EX = geoStartX:  seg.EY = geoStartY:  seg.EZ = chainZ
            End If

            seg.SegLen = R * Abs(sw)

            chainX = seg.EX:  chainY = seg.EY:  chainZ = seg.EZ
        End If

        pathSegs(i) = seg
        totalPathLen = totalPathLen + seg.SegLen
    Next i

    pathSegCount = nElems
    BuildAlignmentPath = (pathSegCount > 0)
    Exit Function

BuildErr:
    Debug.Print "BuildAlignmentPath error: " & Err.Description
    BuildAlignmentPath = False
End Function

' ============================================================
' INTERPOLATE POINT AND TANGENT AT GIVEN ARC-LENGTH
' ptX/Y/Z  : output - point on alignment
' tanX/Y   : output - unit tangent in travel direction
' Returns False if path is empty.
' ============================================================
Public Function GetPointAndTangent(dist As Double, _
                                   ptX As Double, ptY As Double, ptZ As Double, _
                                   tanX As Double, tanY As Double) As Boolean
    On Error GoTo InterpErr

    If pathSegCount = 0 Then
        GetPointAndTangent = False
        Exit Function
    End If

    ' Clamp to valid range
    If dist < 0 Then dist = 0
    If dist > totalPathLen Then dist = totalPathLen

    Dim cumLen As Double
    cumLen = 0
    Dim i As Integer

    For i = 1 To pathSegCount
        Dim segEnd As Double
        segEnd = cumLen + pathSegs(i).SegLen

        If dist <= segEnd + 0.00001 Then
            Dim t As Double         ' distance into this segment
            t = dist - cumLen
            If t < 0 Then t = 0

            If Not pathSegs(i).IsArc Then
                ' ----- line segment -----
                Dim lLen As Double
                lLen = pathSegs(i).SegLen
                If lLen < 0.000001 Then lLen = 0.000001

                Dim tdx As Double, tdy As Double
                tdx = (pathSegs(i).EX - pathSegs(i).SX) / lLen
                tdy = (pathSegs(i).EY - pathSegs(i).SY) / lLen

                ptX = pathSegs(i).SX + t * tdx
                ptY = pathSegs(i).SY + t * tdy
                ptZ = pathSegs(i).SZ
                tanX = tdx
                tanY = tdy

            Else
                ' ----- arc segment -----
                Dim R   As Double
                Dim sa  As Double
                Dim sw  As Double
                R = pathSegs(i).Radius
                sa = pathSegs(i).StartAngle
                sw = pathSegs(i).SweepAngle

                ' Angle at distance t along arc
                Dim theta As Double
                If Abs(sw) > 0.000001 And R > 0.000001 Then
                    ' sign(sw) gives CW or CCW direction
                    theta = sa + (t / R) * (sw / Abs(sw))
                Else
                    theta = sa
                End If

                ptX = pathSegs(i).CX + R * Cos(theta)
                ptY = pathSegs(i).CY + R * Sin(theta)
                ptZ = pathSegs(i).SZ

                ' Tangent = radius-perpendicular in travel direction
                Dim swSign As Double
                swSign = IIf(sw >= 0, 1, -1)
                tanX = -Sin(theta) * swSign
                tanY = Cos(theta) * swSign
            End If

            ' Normalize tangent
            Dim mag As Double
            mag = Sqr(tanX * tanX + tanY * tanY)
            If mag > 0.000001 Then
                tanX = tanX / mag
                tanY = tanY / mag
            End If

            GetPointAndTangent = True
            Exit Function
        End If

        cumLen = segEnd
    Next i

    ' Fell through - clamp to end
    ptX = pathSegs(pathSegCount).EX
    ptY = pathSegs(pathSegCount).EY
    ptZ = pathSegs(pathSegCount).EZ
    tanX = 1:  tanY = 0
    GetPointAndTangent = True
    Exit Function

InterpErr:
    Debug.Print "GetPointAndTangent error: " & Err.Description
    GetPointAndTangent = False
End Function

' ============================================================
' PLACE A PERPENDICULAR LINE AT THE GIVEN POINT/TANGENT
' The line extends halfLen master units on each side.
' ============================================================
Public Sub PlacePerpendicularLine(ptX As Double, ptY As Double, ptZ As Double, _
                                   tanX As Double, tanY As Double, _
                                   halfLen As Double)
    On Error GoTo PlaceErr

    ' Perpendicular = rotate tangent 90 degrees in XY plane
    Dim perpX As Double, perpY As Double
    perpX = -tanY
    perpY = tanX

    ' Ensure unit length (should already be normalised from GetPointAndTangent)
    Dim mag As Double
    mag = Sqr(perpX * perpX + perpY * perpY)
    If mag > 0.000001 Then
        perpX = perpX / mag
        perpY = perpY / mag
    End If

    Dim pt1 As Point3d, pt2 As Point3d
    pt1.X = ptX - perpX * halfLen
    pt1.Y = ptY - perpY * halfLen
    pt1.Z = ptZ
    pt2.X = ptX + perpX * halfLen
    pt2.Y = ptY + perpY * halfLen
    pt2.Z = ptZ

    Dim lineEl As LineElement
    Set lineEl = CreateLineElement2(Nothing, pt1, pt2)
    ' Set element properties: Default level, color 0 (white), weight 0
    lineEl.Color = 0
    lineEl.LineWeight = 0
    lineEl.Level = ActiveDesignFile.Levels("Default")
    ActiveModelReference.AddElement lineEl
    lineEl.Rewrite

    Exit Sub
PlaceErr:
    MsgBox "Error placing perpendicular line: " & Err.Description, _
           vbCritical, "Placement Error"
End Sub

' ============================================================
' LOOK UP SPACING (master units / ft) FOR A WZTC LABEL
' ============================================================
Public Function GetSpacingForLabel(label As String) As Double
    Select Case Trim(label)
        Case "Downstream Taper":       GetSpacingForLabel = Val(wztcDownstreamTaper)
        Case "Roll Ahead Distance":    GetSpacingForLabel = Val(wztcRollAhead)
        Case "Vehicle Space":          GetSpacingForLabel = Val(wztcVehicleSpace)
        Case "Buffer Space":           GetSpacingForLabel = Val(wztcBufferSpace)
        Case "Merging/Shifting Taper": GetSpacingForLabel = Val(wztcMergingTaper)
        Case "Shoulder Taper":         GetSpacingForLabel = Val(wztcShoulderTapers)
        Case "Work Area":              GetSpacingForLabel = 0
        Case Else
            ' Sign number - look up in sign table
            Dim i As Integer
            For i = 1 To wztcSignCount
                If Trim(wztcSignNumbers(i)) = Trim(label) Then
                    GetSpacingForLabel = Val(wztcSignSpacings(i))
                    Exit Function
                End If
            Next i
            GetSpacingForLabel = 0
    End Select
End Function

' ============================================================
' PUBLIC STATE ACCESSORS (called by PlacePerp)
' ============================================================

Public Function GetCurrentItemLabel() As String
    If currentItemIdx >= 0 And currentItemIdx < wztcOrderLabelCount Then
        GetCurrentItemLabel = wztcOrderLabels(currentItemIdx)
    Else
        GetCurrentItemLabel = ""
    End If
End Function

Public Function GetCurrentItemSuggestedSpacing() As Double
    If currentItemIdx >= 0 And currentItemIdx < wztcOrderLabelCount Then
        GetCurrentItemSuggestedSpacing = GetSpacingForLabel(wztcOrderLabels(currentItemIdx))
    Else
        GetCurrentItemSuggestedSpacing = 0
    End If
End Function

' Place the perpendicular line for the current item using the given spacing,
' then advance to the next item.
' If the item is a sign number, records its geometry for the sign drawing step.
Public Sub PlaceLineForCurrentItem(spacing As Double)
    currentPathPos = currentPathPos + spacing

    Dim ptX As Double, ptY As Double, ptZ As Double
    Dim tanX As Double, tanY As Double
    Call GetPointAndTangent(currentPathPos, ptX, ptY, ptZ, tanX, tanY)
    Call PlacePerpendicularLine(ptX, ptY, ptZ, tanX, tanY, PERP_HALF_LEN)

    ' If this item is a sign number, store its perpendicular line geometry
    Dim lbl As String
    lbl = GetCurrentItemLabel()
    If IsSignLabel(lbl) Then
        Dim n As Integer
        n = wztcPlacedSignCount + 1
        wztcPlacedSignCount = n
        ReDim Preserve wztcPlacedSignNums(1 To n)
        ReDim Preserve wztcPlacedSignPtX(1 To n)
        ReDim Preserve wztcPlacedSignPtY(1 To n)
        ReDim Preserve wztcPlacedSignPtZ(1 To n)
        ReDim Preserve wztcPlacedSignPerpX(1 To n)
        ReDim Preserve wztcPlacedSignPerpY(1 To n)
        ReDim Preserve wztcPlacedSignSide(1 To n)
        ReDim Preserve wztcPlacedSignSize(1 To n)
        wztcPlacedSignNums(n) = lbl
        wztcPlacedSignPtX(n) = ptX
        wztcPlacedSignPtY(n) = ptY
        wztcPlacedSignPtZ(n) = ptZ
        ' Perpendicular vector = tangent rotated 90 degrees CCW
        wztcPlacedSignPerpX(n) = -tanY
        wztcPlacedSignPerpY(n) = tanX
        ' Look up side and size from the sign selection table
        Dim i As Integer
        wztcPlacedSignSide(n) = "Both Sides"
        wztcPlacedSignSize(n) = ""
        For i = 1 To wztcSignCount
            If Trim(wztcSignNumbers(i)) = Trim(lbl) Then
                wztcPlacedSignSide(n) = wztcSignSides(i)
                wztcPlacedSignSize(n) = wztcSignSizes(i)
                Exit For
            End If
        Next i
    End If

    currentItemIdx = currentItemIdx + 1
End Sub

' Advance past the current item without placing a line or advancing the path position.
' The next item will be placed at the same cumulative position as this item would have used.
Public Sub SkipCurrentItem()
    currentItemIdx = currentItemIdx + 1
End Sub

Public Function GetCurrentItemNumber() As Integer
    GetCurrentItemNumber = currentItemIdx + 1   ' 1-based for display
End Function

Public Function GetTotalItemCount() As Integer
    GetTotalItemCount = wztcOrderLabelCount
End Function

Public Function GetCurrentPosition() As Double
    GetCurrentPosition = currentPathPos
End Function

Public Function GetTotalPathLength() As Double
    GetTotalPathLength = totalPathLen
End Function

Public Function IsAllDone() As Boolean
    IsAllDone = (currentItemIdx >= wztcOrderLabelCount)
End Function

' ============================================================
' HELPER: Is the given label a sign number (not a spacing item)?
' Spacing items are the fixed parameter names; everything else
' is a sign number entered in the sign selection table.
' ============================================================
Private Function IsSignLabel(lbl As String) As Boolean
    Select Case Trim(lbl)
        Case "Downstream Taper", "Roll Ahead Distance", "Vehicle Space", _
             "Buffer Space", "Merging/Shifting Taper", "Shoulder Taper", "Work Area"
            IsSignLabel = False
        Case Else
            IsSignLabel = (Trim(lbl) <> "")
    End Select
End Function

' ============================================================
' TRY TO COMPUTE ARC CENTER FROM CHAIN POINT AND ANGLES
' Tests both candidate centers (chainPt at sa vs chainPt at sa+sw)
' and validates the arc midpoint against the bounding box.
' Returns True if a valid center was found.
' ============================================================
Private Function TryComputeCenter(chainX As Double, chainY As Double, _
                                   R As Double, sa As Double, sw As Double, _
                                   rng As Range3d, tol As Double, _
                                   ctrX As Double, ctrY As Double) As Boolean
    ' Candidate A: chainPt is at geometric start angle (sa)
    Dim ctrXA As Double, ctrYA As Double
    ctrXA = chainX - R * Cos(sa)
    ctrYA = chainY - R * Sin(sa)

    ' Candidate B: chainPt is at geometric end angle (sa + sw)
    Dim ctrXB As Double, ctrYB As Double
    ctrXB = chainX - R * Cos(sa + sw)
    ctrYB = chainY - R * Sin(sa + sw)

    ' Validate: the arc midpoint (at angle sa + sw/2) for the correct
    ' candidate must fall inside the element's bounding box.
    Dim midAngle As Double
    midAngle = sa + sw / 2

    ' Test candidate A
    Dim midXA As Double, midYA As Double
    midXA = ctrXA + R * Cos(midAngle)
    midYA = ctrYA + R * Sin(midAngle)

    If midXA >= rng.Low.X - tol And midXA <= rng.High.X + tol And _
       midYA >= rng.Low.Y - tol And midYA <= rng.High.Y + tol Then
        ctrX = ctrXA:  ctrY = ctrYA
        TryComputeCenter = True
        Exit Function
    End If

    ' Test candidate B
    Dim midXB As Double, midYB As Double
    midXB = ctrXB + R * Cos(midAngle)
    midYB = ctrYB + R * Sin(midAngle)

    If midXB >= rng.Low.X - tol And midXB <= rng.High.X + tol And _
       midYB >= rng.Low.Y - tol And midYB <= rng.High.Y + tol Then
        ctrX = ctrXB:  ctrY = ctrYB
        TryComputeCenter = True
        Exit Function
    End If

    TryComputeCenter = False
End Function


