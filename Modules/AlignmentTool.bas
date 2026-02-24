Option Explicit

' ============================================================
' ALIGNMENT DRAWING TOOL
' ------------------------------------------------------------
' Supports drawing multiple named alignments (Upstream,
' Downstream, plus any additional alignments the user added
' in WZTCDesigner). Each alignment is drawn independently:
' the user selects an alignment from the dropdown in AlignDraw,
' draws line/arc segments, then clicks "Commit This Alignment"
' to assign those elements to a MicroStation graphic group.
' After all desired alignments are committed, the user clicks
' "Next: Place Perp Lines" to launch PerpPlacement.
' ============================================================

' Per-alignment endpoint memory (1-based, indexed by alignment number)
Private lastPointArr(1 To 10)            As Point3d
Private hasLastPointArr(1 To 10)         As Boolean
Private alignFirstPtCapturedArr(1 To 10) As Boolean

' Which alignment the user is currently drawing
Private currentAlignIdx As Integer

' ============================================================
' SETUP VIEW AND ELEMENT PROPERTIES (shared helper)
' ============================================================
Private Sub SetupView()
    On Error Resume Next
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
End Sub

' ============================================================
' GET CURRENT MAX ELEMENT ID IN MODEL
' ============================================================
Public Function GetCurrentMaxID() As Double
    Dim el As Element
    Dim oEnum As ElementEnumerator
    Dim oScan As ElementScanCriteria
    Dim maxID As Double
    maxID = 0
    On Error GoTo MaxIDErr
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim elID As Double
        elID = ElIDAsDouble(el.ID)
        If elID > maxID Then maxID = elID
    Loop
MaxIDErr:
    GetCurrentMaxID = maxID
End Function

' ============================================================
' SWITCH TO A DIFFERENT ALIGNMENT
' Called when user changes dropdown in AlignDraw form.
' On first visit: snapshots current max element ID.
' Subsequent visits restore the per-alignment endpoint state.
' ============================================================
Public Sub SetCurrentAlignment(aIdx As Integer)
    If aIdx < 1 Or aIdx > 10 Then Exit Sub
    currentAlignIdx = aIdx
    wztcCurrentAlignDrawIdx = aIdx

    ' First time drawing this alignment: snapshot the max element ID
    ' so CommitCurrentAlignment can identify which elements belong here.
    If wztcAlignMaxIDSnapshot(aIdx) = 0 Then
        wztcAlignMaxIDSnapshot(aIdx) = GetCurrentMaxID()
    End If
End Sub

' ============================================================
' START WZTC DRAWING SESSION
' Called from DrawWorkSpace.frm "Next: Draw Alignments" button.
' Resets state, sets up view, shows AlignDraw form.
' ============================================================
Public Sub StartWZTCDrawing()
    Dim i As Integer
    ' Reset all per-alignment state for a fresh drawing session
    For i = 1 To 10
        hasLastPointArr(i) = False
        alignFirstPtCapturedArr(i) = False
        wztcAlignMaxIDSnapshot(i) = 0
        wztcAlignGraphicGroup(i) = 0
        wztcAlignDrawn(i) = False
    Next i
    currentAlignIdx = 1
    wztcCurrentAlignDrawIdx = 1

    Call SetupView

    ' Reset per-alignment session tracking
    For i = 1 To 10
        wztcAlignSessionCount(i) = 0
    Next i

    ' Show AlignDraw (it will call SetCurrentAlignment(1) in Initialize)
    AlignDraw.Show vbModeless
End Sub

' ============================================================
' RECORD A DRAWING SESSION FOR AN ALIGNMENT
' Called from AlignDraw cmdStartSegment_Click after each
' drawing session. startID = maxID before drawing; endID = maxID after.
' Ignored if no new elements were drawn (endID <= startID).
' ============================================================
Public Sub RecordAlignmentSession(aIdx As Integer, startID As Double, endID As Double)
    If aIdx < 1 Or aIdx > 10 Then Exit Sub
    If endID <= startID Then Exit Sub   ' nothing drawn this session
    Dim s As Integer
    s = wztcAlignSessionCount(aIdx) + 1
    If s > 50 Then s = 50   ' cap at 50 sessions per alignment
    wztcAlignSessionCount(aIdx) = s
    wztcAlignSessionStartIDs(aIdx, s) = startID
    wztcAlignSessionEndIDs(aIdx, s) = endID
End Sub

' ============================================================
' START A LINE SEGMENT (for current alignment)
' ============================================================
Public Sub StartLineSegment()
    Dim oMsg As CadInputMessage
    Dim currentPoint As Point3d
    Dim aIdx As Integer
    aIdx = currentAlignIdx
    If aIdx < 1 Or aIdx > 10 Then aIdx = 1

    CadInputQueue.SendReset

    If hasLastPointArr(aIdx) Then
        ' Continue from last endpoint of this alignment
        currentPoint = lastPointArr(aIdx)
        CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
        CadInputQueue.SendDataPoint currentPoint, 1
    Else
        ' First segment for this alignment — wait for first click
        CadInputQueue.SendKeyin "ECHO Click first point, right-click to cancel"
        CadInputQueue.SendCommand "NULL"

        Set oMsg = CadInputQueue.GetInput
        Do While oMsg.InputType <> msdCadInputTypeDataPoint
            If oMsg.InputType = msdCadInputTypeReset Then
                CommandState.StartDefaultCommand
                Exit Sub
            End If
            Set oMsg = CadInputQueue.GetInput
        Loop

        currentPoint = oMsg.Point
        lastPointArr(aIdx) = currentPoint
        hasLastPointArr(aIdx) = True

        ' Capture alignment start point (per-alignment + alignment 1 backward-compat)
        If Not alignFirstPtCapturedArr(aIdx) Then
            wztcAlignFirstPtX(aIdx) = currentPoint.X
            wztcAlignFirstPtY(aIdx) = currentPoint.Y
            wztcAlignFirstPtZ(aIdx) = currentPoint.Z
            If aIdx = 1 Then
                wztcAlignmentFirstPointX = currentPoint.X
                wztcAlignmentFirstPointY = currentPoint.Y
                wztcAlignmentFirstPointZ = currentPoint.Z
            End If
            alignFirstPtCapturedArr(aIdx) = True
        End If

        CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
        CadInputQueue.SendDataPoint currentPoint, 1
    End If

    CadInputQueue.SendKeyin "ECHO Click next points, right-click to finish segment"
    Do
        Set oMsg = CadInputQueue.GetInput
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
            lastPointArr(aIdx) = oMsg.Point
        ElseIf oMsg.InputType = msdCadInputTypeReset Then
            Exit Do
        End If
    Loop

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
End Sub

' ============================================================
' START AN ARC SEGMENT (for current alignment)
' ============================================================
Public Sub StartArcSegment()
    Dim oMsg As CadInputMessage
    Dim firstPoint As Point3d
    Dim pointCount As Integer
    Dim aIdx As Integer
    aIdx = currentAlignIdx
    If aIdx < 1 Or aIdx > 10 Then aIdx = 1

    CadInputQueue.SendReset

    If hasLastPointArr(aIdx) Then
        firstPoint = lastPointArr(aIdx)
    Else
        CadInputQueue.SendKeyin "ECHO Click first point of arc, right-click to cancel"
        CadInputQueue.SendCommand "NULL"

        Set oMsg = CadInputQueue.GetInput
        Do While oMsg.InputType <> msdCadInputTypeDataPoint
            If oMsg.InputType = msdCadInputTypeReset Then
                CommandState.StartDefaultCommand
                Exit Sub
            End If
            Set oMsg = CadInputQueue.GetInput
        Loop

        firstPoint = oMsg.Point
        lastPointArr(aIdx) = firstPoint
        hasLastPointArr(aIdx) = True

        ' Capture alignment start point (per-alignment + alignment 1 backward-compat)
        If Not alignFirstPtCapturedArr(aIdx) Then
            wztcAlignFirstPtX(aIdx) = firstPoint.X
            wztcAlignFirstPtY(aIdx) = firstPoint.Y
            wztcAlignFirstPtZ(aIdx) = firstPoint.Z
            If aIdx = 1 Then
                wztcAlignmentFirstPointX = firstPoint.X
                wztcAlignmentFirstPointY = firstPoint.Y
                wztcAlignmentFirstPointZ = firstPoint.Z
            End If
            alignFirstPtCapturedArr(aIdx) = True
        End If
    End If

    ' Activate 3-point arc mode and seed with start point
    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"
    CadInputQueue.SendDataPoint firstPoint, 1

    ' Collect end point (stored as next start) then arc point
    pointCount = 1
    CadInputQueue.SendKeyin "ECHO Click arc end point, then a point on the arc"
    Do
        Set oMsg = CadInputQueue.GetInput
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
            pointCount = pointCount + 1
            If pointCount = 2 Then
                lastPointArr(aIdx) = oMsg.Point  ' arc endpoint = next start
            End If
            If pointCount >= 3 Then Exit Do
        ElseIf oMsg.InputType = msdCadInputTypeReset Then
            Exit Do
        End If
    Loop

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
End Sub

' ============================================================
' COMMIT CURRENT ALIGNMENT
' Called by AlignDraw "Commit All Alignments" button.
' Groups elements from all recorded drawing sessions for this
' alignment into a new MicroStation graphic group.
'
' Uses session-based ID ranges (wztcAlignSessionStartIDs /
' wztcAlignSessionEndIDs) so that alignments drawn in any order
' — even interleaved — are correctly separated.
' ============================================================
Public Sub CommitCurrentAlignment()
    On Error GoTo CommitErr
    Dim aIdx As Integer
    aIdx = currentAlignIdx
    If aIdx < 1 Or aIdx > 10 Then Exit Sub

    If wztcAlignSessionCount(aIdx) = 0 Then
        MsgBox "No segments recorded for this alignment." & vbCrLf & _
               "Click 'Start Segment' to draw before committing.", _
               vbExclamation, "Commit Alignment"
        Exit Sub
    End If

    ' Find the highest graphic group number currently in the model
    Dim maxGG As Long
    maxGG = 0
    Dim el As Element
    Dim oEnum As ElementEnumerator
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If el.GraphicGroup > maxGG Then maxGG = el.GraphicGroup
    Loop

    Dim newGG As Long
    newGG = maxGG + 1

    ' Assign elements whose IDs fall within any recorded session for this alignment.
    ' Elements already in a group are skipped (already committed to another alignment).
    Dim groupedCount As Long
    groupedCount = 0
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If el.GraphicGroup = 0 Then
            Dim elIDDbl As Double
            elIDDbl = ElIDAsDouble(el.ID)
            Dim s As Integer
            Dim inSession As Boolean
            inSession = False
            For s = 1 To wztcAlignSessionCount(aIdx)
                If elIDDbl > wztcAlignSessionStartIDs(aIdx, s) And _
                   elIDDbl <= wztcAlignSessionEndIDs(aIdx, s) Then
                    inSession = True
                    Exit For
                End If
            Next s
            If inSession Then
                el.GraphicGroup = newGG
                el.Rewrite
                groupedCount = groupedCount + 1
            End If
        End If
    Loop

    If groupedCount = 0 Then
        MsgBox "No elements found for this alignment." & vbCrLf & _
               "Please draw some line/arc segments first.", _
               vbExclamation, "Commit Alignment"
        Exit Sub
    End If

    ' Store graphic group and mark alignment as drawn
    wztcAlignGraphicGroup(aIdx) = CInt(newGG)
    wztcAlignDrawn(aIdx) = True

    ' Update snapshot for backward compat (used by legacy BuildAlignmentPath fallback)
    wztcAlignMaxIDSnapshot(aIdx) = GetCurrentMaxID()
    If aIdx = 1 Then wztcAlignmentStartMaxID = wztcAlignMaxIDSnapshot(aIdx)

    CadInputQueue.SendKeyin "ECHO " & wztcAlignNames(aIdx) & " committed: " & _
                            groupedCount & " elements in graphic group " & newGG

    Exit Sub
CommitErr:
    MsgBox "Error committing alignment: " & Err.Description, vbCritical, "Commit Error"
End Sub

' ============================================================
' LEGACY ENTRY POINT (kept for back-compatibility references)
' Previously called from AlignDraw.cmdDone_Click.
' Now replaced by CommitCurrentAlignment + AlignDraw cmdNextStep.
' ============================================================
Public Sub GroupAndLaunchPlacement()
    Call CommitCurrentAlignment
    ' Note: perp placement is now launched from AlignDraw.cmdNextStep_Click
    ' directly via StartAlignmentPlacement (PerpPlacement.bas).
End Sub
