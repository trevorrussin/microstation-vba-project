Option Explicit

' Remembers the last placed endpoint so the next segment starts there
Private lastPoint As Point3d
Private hasLastPoint As Boolean

' Tracks whether the very first click point of THIS alignment session
' has been captured and saved to wztcAlignmentFirstPoint*.
Private alignmentFirstPointCaptured As Boolean

' ============================================================
' ALIGNMENT DRAWING TOOL - MAIN ENTRY POINT
' ============================================================
Public Sub StartAlignmentTool()
    Dim v As View

    ' Reset endpoint memory for this session
    hasLastPoint = False

    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw

    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"

    AlignmentForm.Show vbModeless
End Sub

' ============================================================
' START A LINE SEGMENT
' ============================================================
' If a previous endpoint is stored, starts the line from there
' automatically. Otherwise waits for the user to click a first point.
' Subsequent clicks extend the line; right-click finishes and
' stores the last point for the next segment.
Public Sub StartLineSegment()
    Dim oMsg As CadInputMessage
    Dim currentPoint As Point3d

    CadInputQueue.SendReset

    If hasLastPoint Then
        ' Continue from where the last segment ended
        currentPoint = lastPoint
        CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
        CadInputQueue.SendDataPoint currentPoint, 1
    Else
        ' First segment - wait for user to click start point
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
        lastPoint = currentPoint
        hasLastPoint = True

        ' Capture the alignment start point on the very first segment
        If Not alignmentFirstPointCaptured Then
            wztcAlignmentFirstPointX = currentPoint.X
            wztcAlignmentFirstPointY = currentPoint.Y
            wztcAlignmentFirstPointZ = currentPoint.Z
            alignmentFirstPointCaptured = True
        End If

        CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
        CadInputQueue.SendDataPoint currentPoint, 1
    End If

    ' Collect further clicks; each one extends the line
    CadInputQueue.SendKeyin "ECHO Click next points, right-click to finish"
    Do
        Set oMsg = CadInputQueue.GetInput
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
            lastPoint = oMsg.Point      ' remember for next segment
        ElseIf oMsg.InputType = msdCadInputTypeReset Then
            Exit Do
        End If
    Loop

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
End Sub

' ============================================================
' START AN ARC SEGMENT
' ============================================================
' Same endpoint-memory behaviour as StartLineSegment.
' 3-point arc: start -> end -> point on arc.
' The end point (2nd click) is stored as the next start.
Public Sub StartArcSegment()
    Dim oMsg As CadInputMessage
    Dim firstPoint As Point3d
    Dim pointCount As Integer

    CadInputQueue.SendReset

    If hasLastPoint Then
        firstPoint = lastPoint
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
        lastPoint = firstPoint
        hasLastPoint = True

        ' Capture the alignment start point on the very first segment
        If Not alignmentFirstPointCaptured Then
            wztcAlignmentFirstPointX = firstPoint.X
            wztcAlignmentFirstPointY = firstPoint.Y
            wztcAlignmentFirstPointZ = firstPoint.Z
            alignmentFirstPointCaptured = True
        End If
    End If

    ' Activate 3-point arc and seed with start point
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
                lastPoint = oMsg.Point  ' arc endpoint becomes next start
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
' START WZTC ALIGNMENT DRAWING TOOL
' ============================================================
' Called after frmWorkzoneDesigner submits. Resets state,
' snapshots the current max element ID (so we can identify
' elements drawn as part of this alignment later), and
' launches UserForm2 for alignment drawing.
Public Sub StartWZTCDrawing()
    Dim v As View

    hasLastPoint = False
    alignmentFirstPointCaptured = False

    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw

    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"

    ' Snapshot max element ID before user draws the alignment.
    ' All elements with ID > this value were drawn as part of the alignment.
    Dim el As Element
    Dim maxID As Double
    Dim oEnum As ElementEnumerator
    Dim oScan As ElementScanCriteria
    maxID = 0
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim elID As Double
        elID = ElIDAsDouble(el.ID)
        If elID > maxID Then maxID = elID
    Loop
    wztcAlignmentStartMaxID = maxID

    UserForm2.Show vbModeless
End Sub

' ============================================================
' GROUP ALIGNMENT ELEMENTS AND LAUNCH PLACEMENT TOOL
' ============================================================
' Called by UserForm2.cmdDone_Click after the user finishes
' drawing the alignment. Assigns all newly drawn elements
' (those with ID > wztcAlignmentStartMaxID) to a new
' MicroStation graphic group, then launches the perpendicular
' line placement form.
Public Sub GroupAndLaunchPlacement()
    On Error GoTo GroupError

    ' Find max graphic group number currently in use
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

    ' Assign all elements drawn since the snapshot to the new group
    Dim groupedCount As Long
    groupedCount = 0
    Set oEnum = ActiveModelReference.Scan(oScan)
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If ElIDAsDouble(el.ID) > wztcAlignmentStartMaxID Then
            el.GraphicGroup = newGG
            el.Rewrite
            groupedCount = groupedCount + 1
        End If
    Loop

    If groupedCount = 0 Then
        MsgBox "No alignment elements found." & vbCrLf & _
               "Please draw the alignment before clicking Done.", _
               vbExclamation, "Group Alignment"
        Exit Sub
    End If

    CadInputQueue.SendKeyin "ECHO Alignment grouped: " & groupedCount & _
                            " elements in graphic group " & newGG

    ' Launch the perpendicular line placement tool
    StartAlignmentPlacement

    Exit Sub

GroupError:
    MsgBox "Error grouping alignment elements: " & Err.Description, _
           vbCritical, "Group Error"
End Sub
