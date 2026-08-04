
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
'   3. Each perpendicular line is 80 ft long (40 ft each side).
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
Public currentItemIdx  As Integer   ' 0-based index into current alignment's rows
Public currentPathPos  As Double    ' cumulative distance from alignment start (ft)

' Which alignment is currently being processed through PlacePerp
Private currentProcessingAlignIdx As Integer

' Default half-length of each perpendicular tick line (master units = ft)
' Total line length = 2 * PERP_HALF_LEN = 80 ft (40 ft each side of alignment)
Private Const PERP_HALF_LEN As Double = 40

' ============================================================
' MAIN ENTRY POINT
' ============================================================
Public Sub StartAlignmentPlacement()
    ' Reset perp line tracking (critical: prevents stale IDs from previous sessions)
    wztcPerpLineIDCount = 0
    wztcPlacedSignCount = 0

    ' Find the first committed alignment
    currentProcessingAlignIdx = 0
    Dim i As Integer
    For i = 1 To wztcAlignCount
        If wztcAlignDrawn(i) Then
            currentProcessingAlignIdx = i
            Exit For
        End If
    Next i

    ' Fallback: if new multi-alignment arrays are empty but old alignment exists,
    ' try alignment 1 using legacy wztcAlignmentStartMaxID path
    If currentProcessingAlignIdx = 0 Then
        If wztcAlignmentStartMaxID > 0 Then
            currentProcessingAlignIdx = 1
            wztcAlignDrawn(1) = True
            wztcAlignGraphicGroup(1) = -1   ' flag: use legacy ID-based scan
        Else
            MsgBox "No committed alignments found." & vbCrLf & _
                   "Please commit at least one alignment in AlignDraw first.", _
                   vbExclamation, "Alignment Placement"
            Exit Sub
        End If
    End If

    If Not BuildAlignmentPath(currentProcessingAlignIdx) Then
        MsgBox "Could not build alignment path for " & GetCurrentAlignmentName() & "." & vbCrLf & _
               "Make sure you committed the alignment after drawing it.", _
               vbExclamation, "Alignment Placement"
        Exit Sub
    End If

    If wztcAlignRowCounts(currentProcessingAlignIdx) <= 0 Then
        MsgBox "No rows found for " & GetCurrentAlignmentName() & "." & vbCrLf & _
               "Please add rows in WZTCDesigner first.", _
               vbExclamation, "Alignment Placement"
        Exit Sub
    End If

    currentItemIdx = 0
    currentPathPos = 0

    PlacePerp.Show vbModeless
End Sub

' ============================================================
' HEADLESS ALIGNMENT-PLACEMENT INIT (agent-driven-8-step-wizard plan,
' Component 3) -- same core logic as StartAlignmentPlacement above,
' but takes an EXPLICIT alignIdx (rather than auto-finding "the first
' committed alignment") and reports failure via errMsg instead of
' MsgBox (blocks headlessly -- established WZTCBridge rule, see that
' module's header). Deliberately does NOT reset wztcPerpLineIDCount/
' wztcPlacedSignCount or show PlacePerp.frm -- PlaceAllOrderTableStations
' below controls the reset explicitly so counts accumulate correctly
' across multiple alignments in one agent-driven session.
' ============================================================
Public Function InitAlignmentPlacementHeadless(aIdx As Integer, ByRef errMsg As String) As Boolean
    errMsg = ""
    If aIdx < 1 Or aIdx > wztcAlignCount Then
        errMsg = "alignIdx out of range: " & aIdx
        InitAlignmentPlacementHeadless = False
        Exit Function
    End If
    If Not wztcAlignDrawn(aIdx) Then
        errMsg = "ALIGNMENT_NOT_READY: alignment " & aIdx & " is not committed in the " & _
            "CURRENT session. If this alignment was already drawn earlier (visible on " & _
            "screen, e.g. from before a VBA hot-reload/IDE Reset wiped in-memory session " & _
            "state), do NOT redraw or recommit it -- call adopt_alignment(align_idx=" & _
            aIdx & ", element_id=<its LINE element id>) to re-bind it without redrawing. " & _
            "Only define_alignment_segment + commit_alignment if it genuinely does not " & _
            "exist yet."
        InitAlignmentPlacementHeadless = False
        Exit Function
    End If

    currentProcessingAlignIdx = aIdx

    If Not BuildAlignmentPath(currentProcessingAlignIdx) Then
        errMsg = "could not build alignment path for alignment " & aIdx
        InitAlignmentPlacementHeadless = False
        Exit Function
    End If

    If wztcAlignRowCounts(currentProcessingAlignIdx) <= 0 Then
        errMsg = "no order-table rows for alignment " & aIdx & " -- call BUILD_WZTC_ORDER_TABLE first"
        InitAlignmentPlacementHeadless = False
        Exit Function
    End If

    currentItemIdx = 0
    currentPathPos = 0
    InitAlignmentPlacementHeadless = True
End Function

' ============================================================
' BATCHED ORDER-TABLE STATION WALK (Component 3) -- walks EVERY row
' in aIdx's order table in one call instead of one PlacePerp.frm
' click per item, using the exact same math (PlaceLineForCurrentItem /
' GetPointAndTangent, unchanged) so results are identical to what a
' human clicking through PlacePerp would produce. This is the
' highest-leverage change in the whole plan: replaces what would
' otherwise be one place_perp_line-equivalent call per order item
' (potentially dozens across two alignments) with one call per
' alignment.
' resetSession=True clears wztcPerpLineIDCount/wztcPlacedSignCount
' before starting -- pass True for the first alignment in a fresh
' agent-driven session, False for subsequent alignments so placed-sign
' geometry accumulates correctly across alignments (mirrors
' StartAlignmentPlacement's one-time reset / AdvanceToNextAlignment's
' no-reset semantics in the interactive flow).
' Returns rows: itemNum, label, type, cumulativeStationFt, ptX, ptY,
' ptZ, tanX, tanY, isSign.
' ============================================================
Public Function PlaceAllOrderTableStations(aIdx As Integer, resetSession As Boolean) As String()
    Dim rows() As String
    Dim errMsg As String

    If resetSession Then
        wztcPerpLineIDCount = 0
        wztcPlacedSignCount = 0
    End If

    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        PlaceAllOrderTableStations = rows
        Exit Function
    End If

    Dim rCount As Integer: rCount = wztcAlignRowCounts(aIdx)
    ReDim rows(0 To rCount)
    rows(0) = "itemNum" & vbTab & "label" & vbTab & "type" & vbTab & "cumulativeStationFt" & vbTab & _
              "ptX" & vbTab & "ptY" & vbTab & "ptZ" & vbTab & "tanX" & vbTab & "tanY" & vbTab & "isSign"

    Do While Not IsAllDone()
        Dim itemNum As Integer: itemNum = GetCurrentItemNumber()
        Dim itemLabel As String: itemLabel = GetCurrentItemLabel()
        Dim rowType As String: rowType = wztcAlignRowTypes(aIdx, itemNum)
        Dim spacing As Double: spacing = GetCurrentItemSuggestedSpacing()

        Call PlaceLineForCurrentItem(spacing)

        Dim ptX As Double, ptY As Double, ptZ As Double, tanX As Double, tanY As Double
        Call GetPointAndTangent(currentPathPos, ptX, ptY, ptZ, tanX, tanY)

        rows(itemNum) = itemNum & vbTab & itemLabel & vbTab & rowType & vbTab & Format(currentPathPos, "0.0") & vbTab & _
                       Format(ptX, "0.0####") & vbTab & Format(ptY, "0.0####") & vbTab & Format(ptZ, "0.0####") & vbTab & _
                       Format(tanX, "0.0####") & vbTab & Format(tanY, "0.0####") & vbTab & _
                       IIf(rowType = "Sign", "Y", "N")
    Loop

    PlaceAllOrderTableStations = rows
End Function

' ============================================================
' Walk order-table stations WITHOUT placing perp ticks — same
' cumulative-station math as PlaceAllOrderTableStations.
' Returns rows: itemNum, label, type, cumulativeStationFt, ptX, ptY,
' ptZ, tanX, tanY, isSign, spacingFt
' ============================================================
Public Function EnumerateOrderTableStations(aIdx As Integer) As String()
    Dim rows() As String
    Dim errMsg As String

    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        EnumerateOrderTableStations = rows
        Exit Function
    End If

    Dim rCount As Integer: rCount = wztcAlignRowCounts(aIdx)
    ReDim rows(0 To rCount)
    rows(0) = "itemNum" & vbTab & "label" & vbTab & "type" & vbTab & "cumulativeStationFt" & vbTab & _
              "ptX" & vbTab & "ptY" & vbTab & "ptZ" & vbTab & "tanX" & vbTab & "tanY" & vbTab & _
              "isSign" & vbTab & "spacingFt"

    Do While Not IsAllDone()
        Dim itemNum As Integer: itemNum = GetCurrentItemNumber()
        Dim itemLabel As String: itemLabel = GetCurrentItemLabel()
        Dim rowType As String: rowType = wztcAlignRowTypes(aIdx, itemNum)
        Dim spacing As Double: spacing = GetCurrentItemSuggestedSpacing()

        currentPathPos = currentPathPos + spacing
        Dim ptX As Double, ptY As Double, ptZ As Double, tanX As Double, tanY As Double
        Call GetPointAndTangent(currentPathPos, ptX, ptY, ptZ, tanX, tanY)

        rows(itemNum) = itemNum & vbTab & itemLabel & vbTab & rowType & vbTab & Format(currentPathPos, "0.0") & vbTab & _
                       Format(ptX, "0.0####") & vbTab & Format(ptY, "0.0####") & vbTab & Format(ptZ, "0.0####") & vbTab & _
                       Format(tanX, "0.0####") & vbTab & Format(tanY, "0.0####") & vbTab & _
                       IIf(rowType = "Sign", "Y", "N") & vbTab & Format(spacing, "0.0")
        currentItemIdx = currentItemIdx + 1
    Loop

    EnumerateOrderTableStations = rows
End Function

' Which Non-Sign rows get a NAME label BELOW the dim (dim length stays
' above for every tick-to-tick span). Authority = standard sheet
' get_sheet_requirements elements — never engineer verbal shortcuts.
' Core layout names always labeled; tapers only when listed on the sheet.
' Labels may arrive ALL CAPS from sheet specs — match case-insensitively.
Private Function OrderLabelKind(label As String) As String
    Dim u As String: u = UCase$(Trim(label))
    If InStr(1, u, "ROLL AHEAD", vbBinaryCompare) > 0 Then
        OrderLabelKind = "RollAhead"
    ElseIf InStr(1, u, "VEHICLE SPACE", vbBinaryCompare) > 0 Then
        OrderLabelKind = "VehicleSpace"
    ElseIf InStr(1, u, "BUFFER", vbBinaryCompare) > 0 Then
        OrderLabelKind = "Buffer"
    ElseIf InStr(1, u, "SHOULDER TAPER", vbBinaryCompare) > 0 Then
        OrderLabelKind = "ShoulderTaper"
    ElseIf InStr(1, u, "DOWNSTREAM TAPER", vbBinaryCompare) > 0 Then
        OrderLabelKind = "DownstreamTaper"
    ElseIf InStr(1, u, "MERGING", vbBinaryCompare) > 0 _
        Or InStr(1, u, "SHIFTING TAPER", vbBinaryCompare) > 0 _
        Or InStr(1, u, "LANE TAPER", vbBinaryCompare) > 0 Then
        OrderLabelKind = "MergingTaper"
    ElseIf InStr(1, u, "WORK AREA", vbBinaryCompare) > 0 Then
        OrderLabelKind = "WorkArea"
    Else
        OrderLabelKind = ""
    End If
End Function

Private Function ShouldAnnotateNonSignLabel(label As String, sheetElements As String) As Boolean
    Dim kind As String: kind = OrderLabelKind(label)
    Dim elems As String: elems = sheetElements
    Select Case kind
        Case "RollAhead", "VehicleSpace", "Buffer"
            ShouldAnnotateNonSignLabel = True
        Case "MergingTaper"
            ShouldAnnotateNonSignLabel = (InStr(1, elems, "MergingTaper", vbTextCompare) > 0)
        Case "ShoulderTaper"
            ShouldAnnotateNonSignLabel = (InStr(1, elems, "ShoulderTaper", vbTextCompare) > 0)
        Case "DownstreamTaper"
            ShouldAnnotateNonSignLabel = (InStr(1, elems, "DownstreamTaper", vbTextCompare) > 0)
        Case Else
            ShouldAnnotateNonSignLabel = False
    End Select
End Function

' Outward unit perp from tangent + outwardSign.
Private Sub OutwardUnit(tanX As Double, tanY As Double, outwardSign As Double, _
                        ByRef outX As Double, ByRef outY As Double)
    If outwardSign >= 0 Then
        outX = -tanY: outY = tanX
    Else
        outX = tanY: outY = -tanX
    End If
End Sub

' Alignment station at path start (station 0) — first tick's "from" end.
Private Function PathStartPoint(ByRef x As Double, ByRef y As Double, _
                                ByRef z As Double, ByRef tanX As Double, _
                                ByRef tanY As Double) As Boolean
    On Error GoTo PSFail
    Call GetPointAndTangent(0#, x, y, z, tanX, tanY)
    PathStartPoint = True
    Exit Function
PSFail:
    PathStartPoint = False
End Function

' Place Non-Sign labels centered on the SAME tip-to-tip segment midpoint
' as the matching dimension. sheetElementsPipe gates optional tapers.
' textExtraAlong: feet past tip along outward (default 20; dims at tip).
Public Function PlaceOrderTableLabels(aIdx As Integer, outwardSign As Double, _
                                      Optional textExtraAlong As Double = 20#, _
                                      Optional sheetElementsPipe As String = "") As String()
    Dim rows() As String
    Dim enumRows() As String
    enumRows = EnumerateOrderTableStations(aIdx)
    If enumRows(0) = "error" Then
        PlaceOrderTableLabels = enumRows
        Exit Function
    End If

    Dim errMsg As String
    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        PlaceOrderTableLabels = rows
        Exit Function
    End If
    Dim sx As Double, sy As Double, sz As Double, stanX As Double, stanY As Double
    Dim haveStart As Boolean
    haveStart = PathStartPoint(sx, sy, sz, stanX, stanY)

    Dim nOut As Integer: nOut = 0
    ReDim rows(0 To 0)
    rows(0) = "itemNum" & vbTab & "label" & vbTab & "text" & vbTab & "x" & vbTab & "y" & vbTab & "status"

    Dim i As Integer
    For i = 1 To UBound(enumRows)
        Dim parts() As String
        parts = Split(enumRows(i), vbTab)
        If UBound(parts) < 10 Then GoTo NextLabel
        If parts(9) = "Y" Then GoTo NextLabel

        Dim label As String: label = parts(1)
        If Not ShouldAnnotateNonSignLabel(label, sheetElementsPipe) Then GoTo NextLabel

        Dim spacingFt As String: spacingFt = parts(10)
        If Val(spacingFt) <= 0 Then GoTo NextLabel

        Dim ptX As Double: ptX = CDbl(parts(4))
        Dim ptY As Double: ptY = CDbl(parts(5))
        Dim tanX As Double: tanX = CDbl(parts(7))
        Dim tanY As Double: tanY = CDbl(parts(8))

        Dim x1 As Double, y1 As Double
        If i >= 2 Then
            Dim prev() As String
            prev = Split(enumRows(i - 1), vbTab)
            If UBound(prev) < 5 Then GoTo NextLabel
            x1 = CDbl(prev(4)): y1 = CDbl(prev(5))
        ElseIf haveStart Then
            x1 = sx: y1 = sy
        Else
            GoTo NextLabel
        End If

        Dim outX As Double, outY As Double
        Call OutwardUnit(tanX, tanY, outwardSign, outX, outY)

        ' Tip-to-tip midpoint (same X as dim). Place BELOW the dim line:
        ' dim sits at tip+offsetDist (~15); label further outward so
        ' length stays above the dim and the name sits below (sheet pattern).
        Dim midX As Double, midY As Double
        midX = 0.5 * (x1 + ptX) + outX * PERP_HALF_LEN
        midY = 0.5 * (y1 + ptY) + outY * PERP_HALF_LEN
        Dim labelOut As Double
        labelOut = 15# + textExtraAlong   ' past dim line (default offsetDist=15)
        Dim tx As Double, ty As Double
        tx = midX + outX * labelOut
        ty = midY + outY * labelOut

        Dim txt As String
        txt = label
        If Len(Trim(spacingFt)) > 0 And Val(spacingFt) > 0 Then
            txt = label & " " & Format(Val(spacingFt), "0") & "'"
        End If

        Dim result As String
        result = WZTCExec.ExecPlaceTextLabel(txt, tx, ty, 0)

        nOut = nOut + 1
        ReDim Preserve rows(0 To nOut)
        rows(nOut) = parts(0) & vbTab & label & vbTab & txt & vbTab & _
                     Format(tx, "0.0####") & vbTab & Format(ty, "0.0####") & vbTab & result
NextLabel:
    Next i

    PlaceOrderTableLabels = rows
End Function

' Dimension EVERY consecutive tick pair (tip-to-tip), including Sign
' spacings and non-labeled Non-Sign rows. Length text sits above the
' dim line; optional name labels below come from PlaceOrderTableLabels.
' sheetElementsPipe retained for API compat but does NOT gate dims.
Public Function PlaceOrderTableDimensions(aIdx As Integer, outwardSign As Double, _
                                          Optional offsetDist As Double = 15#, _
                                          Optional sheetElementsPipe As String = "") As String()
    Dim rows() As String
    Dim enumRows() As String
    enumRows = EnumerateOrderTableStations(aIdx)
    If enumRows(0) = "error" Then
        PlaceOrderTableDimensions = enumRows
        Exit Function
    End If

    Dim errMsg As String
    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        PlaceOrderTableDimensions = rows
        Exit Function
    End If
    Dim sx As Double, sy As Double, sz As Double, stanX As Double, stanY As Double
    Dim haveStart As Boolean
    haveStart = PathStartPoint(sx, sy, sz, stanX, stanY)

    Dim nOut As Integer: nOut = 0
    ReDim rows(0 To 0)
    rows(0) = "fromItem" & vbTab & "toItem" & vbTab & "spacingFt" & vbTab & "elementId" & vbTab & "status"

    Dim i As Integer
    For i = 1 To UBound(enumRows)
        Dim b() As String
        b = Split(enumRows(i), vbTab)
        If UBound(b) < 10 Then GoTo NextDim

        Dim spacing As Double: spacing = CDbl(b(10))
        If spacing <= 0 Then GoTo NextDim

        Dim x2a As Double: x2a = CDbl(b(4))
        Dim y2a As Double: y2a = CDbl(b(5))
        Dim tanX As Double: tanX = CDbl(b(7))
        Dim tanY As Double: tanY = CDbl(b(8))

        Dim x1a As Double, y1a As Double
        Dim fromItem As String
        If i >= 2 Then
            Dim a() As String
            a = Split(enumRows(i - 1), vbTab)
            If UBound(a) < 5 Then GoTo NextDim
            x1a = CDbl(a(4)): y1a = CDbl(a(5))
            fromItem = a(0)
        ElseIf haveStart Then
            x1a = sx: y1a = sy
            fromItem = "0"
        Else
            GoTo NextDim
        End If

        Dim outX As Double, outY As Double
        Call OutwardUnit(tanX, tanY, outwardSign, outX, outY)

        ' Measure at tick tips so dims line up with the ticks
        Dim t1x As Double, t1y As Double, t2x As Double, t2y As Double
        t1x = x1a + outX * PERP_HALF_LEN: t1y = y1a + outY * PERP_HALF_LEN
        t2x = x2a + outX * PERP_HALF_LEN: t2y = y2a + outY * PERP_HALF_LEN

        Dim ox As Double, oy As Double
        ox = 0.5 * (t1x + t2x) + outX * offsetDist
        oy = 0.5 * (t1y + t2y) + outY * offsetDist

        Dim result As String
        result = WZTCExec.ExecPlaceDimension(t1x, t1y, t2x, t2y, ox, oy, 0)

        Dim eid As String: eid = ""
        Dim fields() As String
        fields = Split(result, vbTab)
        Dim f As Integer
        For f = 0 To UBound(fields)
            If Left(fields(f), 10) = "elementId=" Then eid = Mid(fields(f), 11)
        Next f

        nOut = nOut + 1
        ReDim Preserve rows(0 To nOut)
        rows(nOut) = fromItem & vbTab & b(0) & vbTab & Format(spacing, "0.0") & vbTab & eid & vbTab & result
NextDim:
    Next i

    PlaceOrderTableDimensions = rows
End Function

' ProtectiveVehicle centered between the two bay ticks (Vehicle Space when
' present; Buffer Space on shoulder sheets that omit VS — e.g. 619-301).
' ArrowPanel at Shoulder Taper tip; fallback Merging/Lane taper tip.
' sheetElementsPipe: pipe list from get_sheet_requirements elements field.
Public Function PlaceSheetSymbolCells(aIdx As Integer, outwardSign As Double, _
                                      sheetElementsPipe As String) As String()
    Dim rows() As String
    Dim enumRows() As String
    enumRows = EnumerateOrderTableStations(aIdx)
    If enumRows(0) = "error" Then
        PlaceSheetSymbolCells = enumRows
        Exit Function
    End If

    Dim nOut As Integer: nOut = 0
    ReDim rows(0 To 0)
    rows(0) = "sheetElement" & vbTab & "cellName" & vbTab & "x" & vbTab & "y" & vbTab & "angleDeg" & vbTab & "status"

    Dim wantPV As Boolean, wantAP As Boolean
    wantPV = (InStr(1, sheetElementsPipe, "ProtectiveVehicle", vbTextCompare) > 0)
    wantAP = (InStr(1, sheetElementsPipe, "ArrowPanel", vbTextCompare) > 0)
    If Not wantPV And Not wantAP Then
        nOut = 1
        ReDim Preserve rows(0 To 1)
        rows(1) = "-" & vbTab & "-" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & _
                  "OK note=no ProtectiveVehicle/ArrowPanel in sheet elements"
        PlaceSheetSymbolCells = rows
        Exit Function
    End If

    Dim vsIdx As Integer: vsIdx = 0
    Dim bufIdx As Integer: bufIdx = 0
    Dim vsLen As Double: vsLen = 50#
    Dim apIdx As Integer: apIdx = 0
    Dim i As Integer
    For i = 1 To UBound(enumRows)
        Dim parts() As String
        parts = Split(enumRows(i), vbTab)
        If UBound(parts) < 10 Then GoTo ContFind
        Dim kind As String: kind = OrderLabelKind(parts(1))
        If kind = "VehicleSpace" Then
            vsIdx = i
            If Val(parts(10)) > 0 Then vsLen = Val(parts(10))
        ElseIf kind = "Buffer" Then
            bufIdx = i
            ' Buffer Space is a roadway safety distance, not a vehicle-bay
            ' length -- do NOT scale the PV cell to it (was stretching
            ' TWZWVA_P to fill the entire buffer span, e.g. 360ft on
            ' 619-311, ~8x oversized). Keep the sane default vsLen=50
            ' when there's no real Vehicle Space row; only the midpoint
            ' position uses this bay, not its length.
        End If
        ' Prefer Shoulder Taper for AP (sheet callout); else Merging/Lane taper
        If kind = "ShoulderTaper" Then
            apIdx = i
        ElseIf apIdx = 0 And kind = "MergingTaper" Then
            apIdx = i
        End If
ContFind:
    Next i

    Dim bayIdx As Integer: bayIdx = vsIdx
    Dim bayNote As String: bayNote = "Vehicle Space"
    If bayIdx = 0 And bufIdx > 0 Then
        bayIdx = bufIdx
        bayNote = "Buffer Space (no Vehicle Space on this sheet)"
    End If

    If wantPV And bayIdx = 0 Then
        ReDim rows(0 To 1)
        rows(0) = "sheetElement" & vbTab & "cellName" & vbTab & "x" & vbTab & "y" & vbTab & "angleDeg" & vbTab & "status"
        rows(1) = "-" & vbTab & "-" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & _
                  "ERROR note=no Vehicle Space or Buffer Space row for ProtectiveVehicle"
        PlaceSheetSymbolCells = rows
        Exit Function
    End If

    Dim errMsg As String
    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        PlaceSheetSymbolCells = rows
        Exit Function
    End If

    Dim tanX As Double, tanY As Double, outX As Double, outY As Double
    Dim angDeg As Double

    If wantPV Then
        Dim b() As String
        b = Split(enumRows(bayIdx), vbTab)
        Dim x2 As Double: x2 = CDbl(b(4))
        Dim y2 As Double: y2 = CDbl(b(5))
        tanX = CDbl(b(7)): tanY = CDbl(b(8))
        Dim x1 As Double, y1 As Double
        If bayIdx >= 2 Then
            Dim a() As String
            a = Split(enumRows(bayIdx - 1), vbTab)
            x1 = CDbl(a(4)): y1 = CDbl(a(5))
        Else
            Dim sz As Double, stanX As Double, stanY As Double
            If Not PathStartPoint(x1, y1, sz, stanX, stanY) Then
                ReDim rows(0 To 1)
                rows(0) = "error"
                rows(1) = "cannot resolve " & bayNote & " start station"
                PlaceSheetSymbolCells = rows
                Exit Function
            End If
        End If
        Call OutwardUnit(tanX, tanY, outwardSign, outX, outY)
        Dim tip1x As Double, tip1y As Double, tip2x As Double, tip2y As Double
        tip1x = x1 + outX * PERP_HALF_LEN: tip1y = y1 + outY * PERP_HALF_LEN
        tip2x = x2 + outX * PERP_HALF_LEN: tip2y = y2 + outY * PERP_HALF_LEN
        Dim midX As Double, midY As Double
        midX = 0.5 * (tip1x + tip2x)
        midY = 0.5 * (tip1y + tip2y)
        angDeg = Atn2Deg(tanY, tanX)
        Dim cellNativeLen As Double: cellNativeLen = 45.7
        Dim sc As Double: sc = vsLen / cellNativeLen
        If sc < 0.1 Then sc = 1#
        nOut = nOut + 1
        ReDim Preserve rows(0 To nOut)
        rows(nOut) = PlaceOneSheetCell("ProtectiveVehicle", "TWZWVA_P", midX, midY, angDeg, sc)
        ' Append bay note into status when Buffer fallback was used
        If vsIdx = 0 Then
            rows(nOut) = rows(nOut) & " note=" & bayNote
        End If
    End If

    If wantAP Then
        If apIdx = 0 Then
            nOut = nOut + 1
            ReDim Preserve rows(0 To nOut)
            rows(nOut) = "ArrowPanel" & vbTab & "TWZAP_P" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & _
                         "ERROR note=no Shoulder/Merging Taper row for ArrowPanel"
        Else
            Dim apParts() As String
            apParts = Split(enumRows(apIdx), vbTab)
            Dim apPx As Double: apPx = CDbl(apParts(4))
            Dim apPy As Double: apPy = CDbl(apParts(5))
            Dim apTx As Double: apTx = CDbl(apParts(7))
            Dim apTy As Double: apTy = CDbl(apParts(8))
            Call OutwardUnit(apTx, apTy, outwardSign, outX, outY)
            ' Tip of the taper tick — sheet places AP on the closed-lane
            ' side at the shoulder/lane taper, not next to Vehicle Space.
            Dim apX As Double, apY As Double
            apX = apPx + outX * PERP_HALF_LEN
            apY = apPy + outY * PERP_HALF_LEN
            angDeg = Atn2Deg(apTy, apTx)
            nOut = nOut + 1
            ReDim Preserve rows(0 To nOut)
            rows(nOut) = PlaceOneSheetCell("ArrowPanel", "TWZAP_P", apX, apY, angDeg, 1#)
        End If
    End If

    PlaceSheetSymbolCells = rows
End Function

' Work-space box in the closed lane from path start through end of
' Vehicle Space (protected bay). On shoulder sheets with no VS row
' (619-301), fall back to Buffer Space end — same longitudinal bay the
' PV cell uses. laneWidthFt = closed-lane / closed-shoulder width.
Public Function PlaceOrderTableWorkspace(aIdx As Integer, outwardSign As Double, _
                                         Optional laneWidthFt As Double = 12#) As String
    On Error GoTo WSErr
    Dim enumRows() As String
    enumRows = EnumerateOrderTableStations(aIdx)
    If enumRows(0) = "error" Then
        PlaceOrderTableWorkspace = "ERROR" & vbTab & "note=" & enumRows(1)
        Exit Function
    End If

    Dim errMsg As String
    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        PlaceOrderTableWorkspace = "ERROR" & vbTab & "note=" & errMsg
        Exit Function
    End If

    Dim sx As Double, sy As Double, sz As Double, stanX As Double, stanY As Double
    If Not PathStartPoint(sx, sy, sz, stanX, stanY) Then
        PlaceOrderTableWorkspace = "ERROR" & vbTab & "note=path start unavailable"
        Exit Function
    End If

    Dim vsIdx As Integer: vsIdx = 0
    Dim bufIdx As Integer: bufIdx = 0
    Dim i As Integer
    For i = 1 To UBound(enumRows)
        Dim parts() As String
        parts = Split(enumRows(i), vbTab)
        If UBound(parts) >= 1 Then
            Dim kind As String: kind = OrderLabelKind(parts(1))
            If kind = "VehicleSpace" Then vsIdx = i: Exit For
            If kind = "Buffer" And bufIdx = 0 Then bufIdx = i
        End If
    Next i
    Dim endIdx As Integer: endIdx = vsIdx
    Dim endNote As String: endNote = "Vehicle Space"
    If endIdx = 0 And bufIdx > 0 Then
        endIdx = bufIdx
        endNote = "Buffer Space (no Vehicle Space)"
    End If
    If endIdx = 0 Then
        PlaceOrderTableWorkspace = "ERROR" & vbTab & "note=Vehicle Space/Buffer Space row not found"
        Exit Function
    End If

    Dim vs() As String
    vs = Split(enumRows(endIdx), vbTab)
    Dim ex As Double: ex = CDbl(vs(4))
    Dim ey As Double: ey = CDbl(vs(5))
    Dim tanX As Double: tanX = CDbl(vs(7))
    Dim tanY As Double: tanY = CDbl(vs(8))
    Dim outX As Double, outY As Double
    Call OutwardUnit(tanX, tanY, outwardSign, outX, outY)

    ' Rectangle: align edge → closed-lane edge (laneWidth), start→bay end
    Dim vtsv As String
    vtsv = Format(sx, "0.0####") & "," & Format(sy, "0.0####") & "|" & _
           Format(ex, "0.0####") & "," & Format(ey, "0.0####") & "|" & _
           Format(ex + outX * laneWidthFt, "0.0####") & "," & Format(ey + outY * laneWidthFt, "0.0####") & "|" & _
           Format(sx + outX * laneWidthFt, "0.0####") & "," & Format(sy + outY * laneWidthFt, "0.0####")

    Dim wsRes As String
    wsRes = WZTCExec.ExecPlaceWorkspace(vtsv)
    PlaceOrderTableWorkspace = wsRes & vbTab & "bay=" & endNote
    Exit Function
WSErr:
    PlaceOrderTableWorkspace = "ERROR" & vbTab & "note=" & Err.Description
End Function

' Channelizing: merging/lane-taper diagonal + longitudinal run along the
' closed lane from taper toe back to path start. Shoulder-only sheets
' (no Merging/Lane taper — e.g. 619-301) use Shoulder Taper as the
' primary diagonal instead. Bounded by order-table stations — never
' AccuDraw-length leftovers.
Public Function PlaceOrderTableChannelizing(aIdx As Integer, outwardSign As Double, _
                                            Optional laneWidthFt As Double = 12#) As String
    On Error GoTo ChErr
    Dim enumRows() As String
    enumRows = EnumerateOrderTableStations(aIdx)
    If enumRows(0) = "error" Then
        PlaceOrderTableChannelizing = "ERROR" & vbTab & "note=" & enumRows(1)
        Exit Function
    End If

    Dim errMsg As String
    If Not InitAlignmentPlacementHeadless(aIdx, errMsg) Then
        PlaceOrderTableChannelizing = "ERROR" & vbTab & "note=" & errMsg
        Exit Function
    End If

    Dim sx As Double, sy As Double, sz As Double, stanX As Double, stanY As Double
    If Not PathStartPoint(sx, sy, sz, stanX, stanY) Then
        PlaceOrderTableChannelizing = "ERROR" & vbTab & "note=path start unavailable"
        Exit Function
    End If

    Dim merIdx As Integer: merIdx = 0
    Dim shIdx As Integer: shIdx = 0
    Dim i As Integer
    For i = 1 To UBound(enumRows)
        Dim parts() As String
        parts = Split(enumRows(i), vbTab)
        If UBound(parts) < 1 Then GoTo ContCh
        Dim kind As String: kind = OrderLabelKind(parts(1))
        If kind = "MergingTaper" Then merIdx = i
        If kind = "ShoulderTaper" Then shIdx = i
ContCh:
    Next i

    ' Primary taper: merging/lane when present; else shoulder-only sheets
    Dim primIdx As Integer: primIdx = merIdx
    Dim primName As String: primName = "Merging/Lane Taper"
    If primIdx = 0 And shIdx > 0 Then
        primIdx = shIdx
        primName = "Shoulder Taper"
        shIdx = 0   ' don't double-draw shoulder as lead-in
    End If
    If primIdx = 0 Then
        PlaceOrderTableChannelizing = "ERROR" & vbTab & "note=no Merging/Lane or Shoulder Taper row"
        Exit Function
    End If

    ' Taper toe = start of primary segment (prev station); tip = primary end
    Dim mer() As String
    mer = Split(enumRows(primIdx), vbTab)
    Dim tipX As Double: tipX = CDbl(mer(4))
    Dim tipY As Double: tipY = CDbl(mer(5))
    Dim tanX As Double: tanX = CDbl(mer(7))
    Dim tanY As Double: tanY = CDbl(mer(8))
    Dim toeX As Double, toeY As Double
    If primIdx >= 2 Then
        Dim prev() As String
        prev = Split(enumRows(primIdx - 1), vbTab)
        toeX = CDbl(prev(4)): toeY = CDbl(prev(5))
    Else
        toeX = sx: toeY = sy
    End If

    Dim outX As Double, outY As Double
    Call OutwardUnit(tanX, tanY, outwardSign, outX, outY)

    ' Optional shoulder taper lead-in when both shoulder + merging exist
    Dim ids As String: ids = ""
    If shIdx > 0 And merIdx > 0 And shIdx > merIdx Then
        Dim sh() As String
        sh = Split(enumRows(shIdx), vbTab)
        Dim shX As Double: shX = CDbl(sh(4))
        Dim shY As Double: shY = CDbl(sh(5))
        Dim shVerts As String
        shVerts = Format(shX, "0.0####") & "," & Format(shY, "0.0####") & "|" & _
                  Format(tipX + outX * (laneWidthFt * 0.35), "0.0####") & "," & _
                  Format(tipY + outY * (laneWidthFt * 0.35), "0.0####")
        Dim shRes As String
        shRes = WZTCExec.ExecPlaceElementRun(2, shVerts)
        ids = ids & shRes & " || "
    End If

    ' Primary taper diagonal: align at upstream tip → full offset at toe
    Dim merVerts As String
    merVerts = Format(tipX, "0.0####") & "," & Format(tipY, "0.0####") & "|" & _
               Format(toeX + outX * laneWidthFt, "0.0####") & "," & _
               Format(toeY + outY * laneWidthFt, "0.0####")
    Dim merRes As String
    merRes = WZTCExec.ExecPlaceElementRun(2, merVerts)
    ids = ids & merRes & " || "

    ' Longitudinal closed-lane run: taper toe → path start (work end)
    Dim longVerts As String
    longVerts = Format(toeX + outX * laneWidthFt, "0.0####") & "," & _
                Format(toeY + outY * laneWidthFt, "0.0####") & "|" & _
                Format(sx + outX * laneWidthFt, "0.0####") & "," & _
                Format(sy + outY * laneWidthFt, "0.0####")
    Dim longRes As String
    longRes = WZTCExec.ExecPlaceElementRun(2, longVerts)
    ids = ids & longRes

    PlaceOrderTableChannelizing = "OK" & vbTab & "note=" & primName & "+longitudinal channelizing" & vbTab & "details=" & ids
    Exit Function
ChErr:
    PlaceOrderTableChannelizing = "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function PlaceOneSheetCell(sheetEl As String, cellName As String, _
                                   x As Double, y As Double, angleDeg As Double, _
                                   Optional scaleFactor As Double = 1#) As String
    On Error GoTo CellErr
    Dim beforeMax As Double
    beforeMax = 0
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Dim idv As Double: idv = ElIDAsDouble(oEnum.Current.ID)
        If idv > beforeMax Then beforeMax = idv
    Loop

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & angleDeg
    CadInputQueue.SendKeyin "ACTIVE SCALE " & Format(scaleFactor, "0.#####") & "," & Format(scaleFactor, "0.#####")
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    SetCExpressionValue "tcb->activeCellUtf16", cellName, ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    Dim pt As Point3d
    pt.X = x: pt.Y = y: pt.Z = 0
    CadInputQueue.SendDataPoint pt, 1
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    CadInputQueue.SendKeyin "ACTIVE SCALE 1,1"

    Dim newId As Double: newId = beforeMax
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        idv = ElIDAsDouble(oEnum.Current.ID)
        If idv > newId Then newId = idv
    Loop

    Dim st As String
    If newId > beforeMax Then
        st = "OK elementId=" & CStr(newId) & " scale=" & Format(scaleFactor, "0.###")
    Else
        st = "ERROR note=no new cell element"
    End If
    PlaceOneSheetCell = sheetEl & vbTab & cellName & vbTab & _
                        Format(x, "0.0####") & vbTab & Format(y, "0.0####") & vbTab & _
                        Format(angleDeg, "0.0##") & vbTab & st
    Exit Function
CellErr:
    On Error Resume Next
    CadInputQueue.SendKeyin "ACTIVE SCALE 1,1"
    PlaceOneSheetCell = sheetEl & vbTab & cellName & vbTab & _
                        Format(x, "0.0####") & vbTab & Format(y, "0.0####") & vbTab & _
                        Format(angleDeg, "0.0##") & vbTab & "ERROR note=" & Err.Description
End Function

Private Function Atn2Deg(y As Double, x As Double) As Double
    Const PI As Double = 3.14159265358979
    Dim a As Double
    If Abs(x) < 0.0000001 Then
        If y >= 0 Then Atn2Deg = 90 Else Atn2Deg = -90
        Exit Function
    End If
    a = Atn(y / x)
    If x < 0 Then
        If y >= 0 Then a = a + PI Else a = a - PI
    End If
    Atn2Deg = a * 180# / PI
End Function

' ============================================================
' BUILD ALIGNMENT PATH
' Collects line/arc elements newer than wztcAlignmentStartMaxID,
' sorts them into a connected chain starting at the first
' recorded click point, and fills pathSegs().
' Returns True on success.
' ============================================================
Public Function BuildAlignmentPath(aIdx As Integer) As Boolean
    On Error GoTo BuildErr

    pathSegCount = 0
    totalPathLen = 0
    ReDim pathSegs(0)

    ' --- collect candidate elements ---
    ' Use graphic group (preferred) or legacy maxID scan for alignment 1 fallback
    Dim elems()  As Element
    Dim nElems   As Integer
    nElems = 0
    Dim useGG As Boolean
    Dim gg As Integer
    useGG = (wztcAlignGraphicGroup(aIdx) > 0)
    gg = wztcAlignGraphicGroup(aIdx)

    Dim el As Element
    Dim oEnum As ElementEnumerator
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim include As Boolean
        include = False
        If useGG Then
            If el.GraphicGroup = gg Then include = True
        Else
            ' Legacy: use maxID scan for alignment 1 when graphic group not set
            If ElIDAsDouble(el.ID) > wztcAlignmentStartMaxID Then include = True
        End If
        If include Then
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
    ' Use per-alignment first point if available, else backward-compat global
    If wztcAlignFirstPtX(aIdx) <> 0 Or wztcAlignFirstPtY(aIdx) <> 0 Then
        chainX = wztcAlignFirstPtX(aIdx)
        chainY = wztcAlignFirstPtY(aIdx)
        chainZ = wztcAlignFirstPtZ(aIdx)
    Else
        chainX = wztcAlignmentFirstPointX
        chainY = wztcAlignmentFirstPointY
        chainZ = wztcAlignmentFirstPointZ
    End If

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

    ' Track element ID so Finish can delete only these perp reference lines
    wztcPerpLineIDCount = wztcPerpLineIDCount + 1
    If wztcPerpLineIDCount = 1 Then
        ReDim wztcPerpLineIDs(1 To 1)
    Else
        ReDim Preserve wztcPerpLineIDs(1 To wztcPerpLineIDCount)
    End If
    wztcPerpLineIDs(wztcPerpLineIDCount) = ElIDAsDouble(lineEl.ID)

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
    Dim aIdx As Integer: aIdx = currentProcessingAlignIdx
    Dim rowNum As Integer: rowNum = currentItemIdx + 1   ' 1-based row
    If aIdx >= 1 And rowNum >= 1 And rowNum <= wztcAlignRowCounts(aIdx) Then
        GetCurrentItemLabel = wztcAlignRowLabels(aIdx, rowNum)
    Else
        GetCurrentItemLabel = ""
    End If
End Function

Public Function GetCurrentItemSuggestedSpacing() As Double
    Dim aIdx As Integer: aIdx = currentProcessingAlignIdx
    Dim rowNum As Integer: rowNum = currentItemIdx + 1
    If aIdx >= 1 And rowNum >= 1 And rowNum <= wztcAlignRowCounts(aIdx) Then
        GetCurrentItemSuggestedSpacing = Val(wztcAlignRowSpacings(aIdx, rowNum))
    Else
        GetCurrentItemSuggestedSpacing = 0
    End If
End Function

' Place the perpendicular line for the current item using the given spacing,
' then advance to the next item.
' Only records sign geometry for rows where Type = "Sign" in the alignment table.
Public Sub PlaceLineForCurrentItem(spacing As Double)
    currentPathPos = currentPathPos + spacing

    Dim ptX As Double, ptY As Double, ptZ As Double
    Dim tanX As Double, tanY As Double
    Call GetPointAndTangent(currentPathPos, ptX, ptY, ptZ, tanX, tanY)
    Call PlacePerpendicularLine(ptX, ptY, ptZ, tanX, tanY, PERP_HALF_LEN)

    ' Only record sign geometry if this row is Type = "Sign" in the alignment table
    Dim aIdx As Integer: aIdx = currentProcessingAlignIdx
    Dim rowNum As Integer: rowNum = currentItemIdx + 1
    If aIdx >= 1 And rowNum >= 1 And rowNum <= wztcAlignRowCounts(aIdx) Then
        If wztcAlignRowTypes(aIdx, rowNum) = "Sign" Then
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
            wztcPlacedSignNums(n) = wztcAlignRowLabels(aIdx, rowNum)
            wztcPlacedSignPtX(n) = ptX
            wztcPlacedSignPtY(n) = ptY
            wztcPlacedSignPtZ(n) = ptZ
            ' Perpendicular vector = tangent rotated 90 degrees CCW
            wztcPlacedSignPerpX(n) = -tanY
            wztcPlacedSignPerpY(n) = tanX
            ' Read side and size directly from alignment table
            wztcPlacedSignSide(n) = wztcAlignRowSides(aIdx, rowNum)
            wztcPlacedSignSize(n) = wztcAlignRowSizes(aIdx, rowNum)
            ' Fallback defaults if blank
            If Len(Trim(wztcPlacedSignSide(n))) = 0 Then wztcPlacedSignSide(n) = "Both Sides"
        End If
    End If

    currentItemIdx = currentItemIdx + 1
End Sub

' Advance past the current item without placing a line or advancing the path position.
Public Sub SkipCurrentItem()
    currentItemIdx = currentItemIdx + 1
End Sub

Public Function GetCurrentItemNumber() As Integer
    GetCurrentItemNumber = currentItemIdx + 1   ' 1-based for display
End Function

Public Function GetTotalItemCount() As Integer
    If currentProcessingAlignIdx >= 1 Then
        GetTotalItemCount = wztcAlignRowCounts(currentProcessingAlignIdx)
    Else
        GetTotalItemCount = wztcOrderLabelCount
    End If
End Function

Public Function GetCurrentPosition() As Double
    GetCurrentPosition = currentPathPos
End Function

Public Function GetTotalPathLength() As Double
    GetTotalPathLength = totalPathLen
End Function

' ============================================================
' GET ALIGNMENT VERTICES -- one row per path segment, straight or arc,
' in design-file master units. Lets a Python-side compiler (placement-plan
' redesign, Stage 1) fetch an alignment's geometry ONCE and do its own
' station->XY interpolation locally, instead of one bridge round-trip per
' point via STATION_TO_POINT -- the real cost there is the ~0.4s COM keyin
' round trip per call, not VBA-side work, so this collapses N calls to 1
' for a whole sheet's worth of stations.
'
' Same row/error convention as WZTCQuery.StationToPoint /
' GetAlignmentStationing: rows(0)="error" + rows(1)=message on failure;
' otherwise rows(0) is the header and rows(1..pathSegCount) are data, one
' per BuildAlignmentPath segment in path order (start to end of the
' alignment). Arc fields (cx/cy/radius/startAngle/sweepAngle) are 0 for
' straight segments -- check isArc before using them.
' ============================================================
Public Function GetAlignmentVertices(alignIdx As Integer) As String()
    Dim rows() As String
    Dim errMsg As String
    If Not WZTCQuery.AlignmentIsReady(alignIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        GetAlignmentVertices = rows
        Exit Function
    End If

    Call BuildAlignmentPath(alignIdx)

    ReDim rows(0 To pathSegCount)
    rows(0) = "segIndex" & vbTab & "isArc" & vbTab & _
              "sx" & vbTab & "sy" & vbTab & "sz" & vbTab & _
              "ex" & vbTab & "ey" & vbTab & "ez" & vbTab & _
              "segLen" & vbTab & _
              "cx" & vbTab & "cy" & vbTab & "radius" & vbTab & _
              "startAngle" & vbTab & "sweepAngle"

    Dim i As Integer
    For i = 1 To pathSegCount
        Dim seg As PathSeg: seg = pathSegs(i)
        rows(i) = (i - 1) & vbTab & IIf(seg.IsArc, "Y", "N") & vbTab & _
                  Format(seg.SX, "0.0000") & vbTab & Format(seg.SY, "0.0000") & vbTab & Format(seg.SZ, "0.0000") & vbTab & _
                  Format(seg.EX, "0.0000") & vbTab & Format(seg.EY, "0.0000") & vbTab & Format(seg.EZ, "0.0000") & vbTab & _
                  Format(seg.SegLen, "0.0000") & vbTab & _
                  Format(seg.CX, "0.0000") & vbTab & Format(seg.CY, "0.0000") & vbTab & Format(seg.Radius, "0.0000") & vbTab & _
                  Format(seg.StartAngle, "0.000000") & vbTab & Format(seg.SweepAngle, "0.000000")
    Next i

    GetAlignmentVertices = rows
End Function

Public Function IsAllDone() As Boolean
    If currentProcessingAlignIdx >= 1 Then
        IsAllDone = (currentItemIdx >= wztcAlignRowCounts(currentProcessingAlignIdx))
    Else
        IsAllDone = (currentItemIdx >= wztcOrderLabelCount)
    End If
End Function

' ============================================================
' MULTI-ALIGNMENT: advance to next committed alignment
' Called by PlacePerp "Next Alignment" button after IsAllDone()
' ============================================================
Public Sub AdvanceToNextAlignment()
    Dim nextIdx As Integer
    nextIdx = currentProcessingAlignIdx + 1
    Do While nextIdx <= wztcAlignCount
        If wztcAlignDrawn(nextIdx) Then Exit Do
        nextIdx = nextIdx + 1
    Loop
    If nextIdx > wztcAlignCount Then Exit Sub   ' no more alignments

    currentProcessingAlignIdx = nextIdx
    currentItemIdx = 0
    currentPathPos = 0
    Call BuildAlignmentPath(currentProcessingAlignIdx)
End Sub

' Returns True if there are no more committed alignments after the current one
Public Function IsLastAlignment() As Boolean
    Dim nextIdx As Integer
    nextIdx = currentProcessingAlignIdx + 1
    Do While nextIdx <= wztcAlignCount
        If wztcAlignDrawn(nextIdx) Then
            IsLastAlignment = False
            Exit Function
        End If
        nextIdx = nextIdx + 1
    Loop
    IsLastAlignment = True
End Function

Public Function GetCurrentAlignmentName() As String
    Dim aIdx As Integer: aIdx = currentProcessingAlignIdx
    If aIdx >= 1 And aIdx <= 10 And Len(Trim(wztcAlignNames(aIdx))) > 0 Then
        GetCurrentAlignmentName = wztcAlignNames(aIdx)
    ElseIf aIdx = 1 Then
        GetCurrentAlignmentName = "Upstream Alignment"
    ElseIf aIdx = 2 Then
        GetCurrentAlignmentName = "Downstream Alignment"
    Else
        GetCurrentAlignmentName = "Alignment " & aIdx
    End If
End Function

' ============================================================
' HELPER: Is the given label a sign number (kept for backward compat)?
' Now superceded by checking wztcAlignRowTypes directly in PlaceLineForCurrentItem.
' ============================================================
Private Function IsSignLabel(lbl As String) As Boolean
    Select Case OrderLabelKind(lbl)
        Case "RollAhead", "VehicleSpace", "Buffer", "MergingTaper", _
             "ShoulderTaper", "DownstreamTaper", "WorkArea"
            IsSignLabel = False
        Case Else
            Dim u As String: u = UCase$(Trim(lbl))
            If InStr(1, u, "UPSTREAM TAPER TEMP", vbBinaryCompare) > 0 _
                Or InStr(1, u, "UPSTREAM TAPER BOX", vbBinaryCompare) > 0 Then
                IsSignLabel = False
            Else
                IsSignLabel = (Trim(lbl) <> "")
            End If
    End Select
End Function

' ============================================================
' DELETE ALL PERPENDICULAR REFERENCE LINES
' Scans the model for line elements whose IDs were recorded during
' PlacePerpendicularLine and removes them. Called by PlaceCells Finish.
' Only the exact perp lines are deleted — all other elements are untouched.
' ============================================================
Public Sub DeletePerpLines()
    If wztcPerpLineIDCount = 0 Then
        MsgBox "No perpendicular lines were tracked (count = 0)." & vbCrLf & _
               "Nothing was deleted. This may happen if placement was not run in this session.", _
               vbInformation, "Delete Perp Lines"
        Exit Sub
    End If

    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim toDelete() As Element
    Dim nDelete As Integer
    nDelete = 0

    Dim el As Element
    Dim elIDDbl As Double
    Dim i As Integer
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        elIDDbl = ElIDAsDouble(el.ID)
        For i = 1 To wztcPerpLineIDCount
            If elIDDbl = wztcPerpLineIDs(i) Then
                nDelete = nDelete + 1
                ReDim Preserve toDelete(1 To nDelete)
                Set toDelete(nDelete) = el
                Exit For
            End If
        Next i
    Loop

    For i = 1 To nDelete
        ActiveModelReference.RemoveElement toDelete(i)
    Next i

    wztcPerpLineIDCount = 0
End Sub

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


