Option Explicit

' ============================================================
' WZTC EXEC — HEADLESS COORDINATE-DRIVEN DRAWING PRIMITIVES
' ------------------------------------------------------------
' Zero GetInput calls anywhere in this module — that is the
' defining constraint. Coordinates and directions always come in
' as parameters, already resolved by the caller (an agent
' reasoning about station/obstruction placement, or a future
' VBA chat panel). This module never asks MicroStation to wait
' for a click.
'
' Return convention: "OK<TAB>key=val..." or "ERROR<TAB>note=...".
' No reqId here — that's WZTCBridge's job, so these functions stay
' callable from anything, not just the file-based bridge protocol.
'
' NOTE: cell placement (PLACE_CELL) intentionally still lives in
' WZTCBridge.bas rather than here. It was built and proven
' end-to-end (Python -> COM -> VBA -> response) in M1, before this
' module existed. Moving it here for architectural consistency is
' a real improvement but not a free one — it touches already-
' verified code for no functional gain, so it's deferred rather
' than bundled into this change. New primitives from here on
' belong in this module.
' ============================================================

' ============================================================
' PLACE PERPENDICULAR REFERENCE LINE AT A STATION
' Reuses PerpPlacement's existing pure element-API primitive
' (PlacePerpendicularLine) and its arc-length path engine
' (BuildAlignmentPath / GetPointAndTangent) — this function is a
' thin wrapper, no new geometry code.
' ============================================================
Public Function ExecPlacePerpLine(alignIdx As Integer, sta As Double, _
                                  Optional halfLen As Double = 40) As String
    On Error GoTo PerpError

    Dim errMsg As String
    If Not WZTCQuery.AlignmentIsReady(alignIdx, errMsg) Then
        ExecPlacePerpLine = "ERROR" & vbTab & "note=" & errMsg
        Exit Function
    End If

    Call PerpPlacement.BuildAlignmentPath(alignIdx)

    Dim ptX As Double, ptY As Double, ptZ As Double, tanX As Double, tanY As Double
    If Not PerpPlacement.GetPointAndTangent(sta, ptX, ptY, ptZ, tanX, tanY) Then
        ExecPlacePerpLine = "ERROR" & vbTab & "note=GetPointAndTangent failed at station " & sta
        Exit Function
    End If

    Call PerpPlacement.PlacePerpendicularLine(ptX, ptY, ptZ, tanX, tanY, halfLen)

    ' PlacePerpendicularLine already appends to wztcPerpLineIDs() — no
    ' separate max-ID scan needed, unlike cell placement.
    Dim newID As Double
    newID = wztcPerpLineIDs(wztcPerpLineIDCount)

    ExecPlacePerpLine = "OK" & vbTab & "elementId=" & CStr(newID) & vbTab & _
                        "ptX=" & Format(ptX, "0.00") & vbTab & "ptY=" & Format(ptY, "0.00") & vbTab & _
                        "note=placed " & (halfLen * 2) & "ft perp line at station " & sta
    Exit Function

PerpError:
    ExecPlacePerpLine = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE A SIGN — face cell, post cell, post line, text label,
' and (Both Sides) the connecting arc — all in one call.
'
' Directly reuses DrawSign.PlaceSignFaceAndText / DrawSign.DrawSignPost /
' DrawSign.DrawConnectingArc, which are already implicitly Public
' and already zero-GetInput. Only DrawSignAtPerpLine's own click
' collection is bypassed — everything downstream of the click is
' unmodified working code, called directly, not duplicated.
'
' pt1/dir1 are the resolved post location and outward direction —
' the caller (agent) is responsible for deciding where that is
' (e.g. via WZTCQuery.StationToPoint plus an offset to dodge an
' obstruction), same division of labor as WZTCQuery/WZTCExec
' generally: Exec only executes, it never decides placement.
' pt2/dir2 are required only when side = "Both Sides".
' ============================================================
Public Function ExecPlaceSign(signNum As String, roadType As String, side As String, _
                              pt1X As Double, pt1Y As Double, pt1Z As Double, _
                              dir1X As Double, dir1Y As Double, _
                              Optional pt2X As Double = 0, Optional pt2Y As Double = 0, Optional pt2Z As Double = 0, _
                              Optional dir2X As Double = 0, Optional dir2Y As Double = 0) As String
    On Error GoTo SignError

    If SignLibrary.GetSignCount() = 0 Then Call SignLibrary.InitializeSignLibrary
    If Not SignLibrary.SignExists(signNum) Then
        ExecPlaceSign = "ERROR" & vbTab & "note=sign not found in library: " & signNum & _
            " -- sign numbers in SignLibrary.bas are zero-padded (e.g. W20-01 not W20-1, " & _
            "R04-09 not R4-9) and some MUTCD numbers need a suffix for the specific variant " & _
            "(e.g. W20-01RA for the Road/Ahead 'ROAD WORK AHEAD' face) -- check SignLibrary.bas " & _
            "for the exact key rather than guessing a second time"
        Exit Function
    End If

    Dim sd As signData
    sd = SignLibrary.GetSignData(signNum, roadType)
    Dim signSize As String: signSize = sd.TextLine2

    Dim bothSides As Boolean
    bothSides = (Trim(side) = "Both Sides")
    If bothSides And pt2X = 0 And pt2Y = 0 And dir2X = 0 And dir2Y = 0 Then
        ExecPlaceSign = "ERROR" & vbTab & "note=side=Both Sides requires pt2X/pt2Y/dir2X/dir2Y"
        Exit Function
    End If

    ' Setup — identical sequence to DrawSign.DrawSignAtPerpLine
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    ' Capture view rotation BEFORE resetting to identity, same rationale as
    ' DrawSign.DrawSignAtPerpLine: sign face cell must match the CURRENT
    ' VIEW's angle so it always reads upright, not a hardcoded 0.
    Dim viewAngleDeg As Double
    viewAngleDeg = DrawSign.ViewRotationAngleDegrees(v)
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendCommand "ATTACH LIBRARY " & sd.CellLibraryPath

    Dim pt1 As Point3d
    pt1.X = pt1X: pt1.Y = pt1Y: pt1.Z = pt1Z

    ' pt1 = attachment on the perp tick (outward tip); dir = outward along perp.
    ' PlaceSignAssembly places post/face/stem/label with edge-connected stem.
    Call DrawSign.PlaceSignAssembly(pt1, signNum, signSize, dir1X, dir1Y, viewAngleDeg)

    If bothSides Then
        Dim pt2 As Point3d
        pt2.X = pt2X: pt2.Y = pt2Y: pt2.Z = pt2Z
        Call DrawSign.PlaceSignAssembly(pt2, signNum, signSize, dir2X, dir2Y, viewAngleDeg)
        Call DrawSign.DrawConnectingArc(pt1, pt2)
    End If

    CommandState.StartDefaultCommand
    ExecPlaceSign = "OK" & vbTab & "signNum=" & signNum & vbTab & "size=" & signSize & vbTab & _
                    "note=placed" & IIf(bothSides, " (both sides)", " (one side)")
    Exit Function

SignError:
    ExecPlaceSign = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE A CHANNELIZING DEVICES / REMOVAL STRIPING / BARRIER RUN
' elementIdx: 2=Channelizing Devices, 3=Removal Striping,
'             4=Temporary Barrier, 5=Barrier w/Warning Lights
'             (1=Work Space is a shape, not a line -- use
'             ExecPlaceWorkspace instead)
' verticesTSV: pipe-separated points, "x,y,z" per point, e.g.
'              "1000,1000,0|1050,1010,0|1100,1005,0"
'
' Same ACTIVE LEVEL/COLOR/WEIGHT + PLACE LINE CONSTRAINED sequence
' DrawElements.DrawCurrentElementSegment already uses for these four
' element types -- including its use of SendCommand (not SendKeyin)
' for the ACTIVE settings, which is a known pre-existing inconsistency
' with the rest of the codebase but is what's actually proven working
' for this specific call.
' ============================================================
Public Function ExecPlaceElementRun(elementIdx As Integer, verticesTSV As String) As String
    On Error GoTo ElemError

    If elementIdx < 2 Or elementIdx > 5 Then
        ExecPlaceElementRun = "ERROR" & vbTab & _
            "note=elementIdx must be 2-5 (1=Work Space, use ExecPlaceWorkspace instead)"
        Exit Function
    End If

    Dim lvlName As String
    lvlName = DrawElements.GetElementLevel(elementIdx)

    Dim pts() As String
    pts = Split(verticesTSV, "|")
    Dim n As Integer: n = UBound(pts) + 1
    If n < 2 Then
        ExecPlaceElementRun = "ERROR" & vbTab & "note=need at least 2 vertices"
        Exit Function
    End If

    ' Element API polyline — avoids AccuDraw distance lock from a prior
    ' long alignment PLACE/define corrupting CadInputQueue PLACE LINE
    ' (live 2026-08-03: channelizing came out as ~3000ft diagonals).
    Dim vpts() As Point3d
    ReDim vpts(0 To n - 1)
    Dim i As Integer, coords() As String
    For i = 0 To n - 1
        coords = Split(pts(i), ",")
        vpts(i).X = CDbl(coords(0))
        vpts(i).Y = CDbl(coords(1))
        If UBound(coords) >= 2 Then vpts(i).Z = CDbl(coords(2)) Else vpts(i).Z = 0
    Next i

    Dim lineEl As LineElement
    Set lineEl = CreateLineElement1(Nothing, vpts)
    lineEl.Color = 6
    lineEl.LineWeight = 2
    On Error Resume Next
    lineEl.Level = ActiveDesignFile.Levels(lvlName)
    On Error GoTo ElemError
    ActiveModelReference.AddElement lineEl
    lineEl.Rewrite

    ExecPlaceElementRun = "OK" & vbTab & "level=" & lvlName & vbTab & "vertexCount=" & n & vbTab & _
                          "elementId=" & CStr(ElIDAsDouble(lineEl.ID)) & vbTab & _
                          "note=placed " & DrawElements.GetElementName(elementIdx)
    Exit Function

ElemError:
    ExecPlaceElementRun = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE THE WORK SPACE BOUNDARY SHAPE + HATCH
' verticesTSV: pipe-separated "x,y,z" points, same format as
' ExecPlaceElementRun. Shape is closed by repeating the first
' vertex as the final data point, matching Legacy Files/LegacyElements.bas.
'
' Hatch uses ClosedElement.SetPattern + CreateHatchPattern1 (Element API),
' not CadInputQueue HATCH ICON — live DELETE.dgn probes 2026-08-02 showed
' HATCH ICON completing with +0 elements / no reliable pattern. Associative
' hatch attaches to the shape (HasPattern=True); it does not add a new
' graphical element ID.
' ============================================================
Public Function ExecPlaceWorkspace(verticesTSV As String) As String
    On Error GoTo WsError

    Dim pts() As String
    pts = Split(verticesTSV, "|")
    Dim n As Integer: n = UBound(pts) + 1
    If n < 3 Then
        ExecPlaceWorkspace = "ERROR" & vbTab & "note=work space needs at least 3 vertices"
        Exit Function
    End If

    Dim vpts() As Point3d
    ReDim vpts(0 To n - 1)
    Dim i As Integer, coords() As String
    For i = 0 To n - 1
        coords = Split(pts(i), ",")
        vpts(i).X = CDbl(coords(0))
        vpts(i).Y = CDbl(coords(1))
        If UBound(coords) >= 2 Then vpts(i).Z = CDbl(coords(2)) Else vpts(i).Z = 0
    Next i

    ' Element API shape — PLACE SHAPE CONSTRAINED was returning OK with no
    ' element after AccuDraw-tainted sessions (live 2026-08-03, 3 retries).
    ' FillMode 0 = none. Named constant msdFillModeNone is NOT defined in
    ' this MicroStation VBA project (compile: Variable not defined,
    ' live 2026-08-03) — use the numeric enum. Passing fill as 3rd arg;
    ' assigning .FillMode after create also raised a blocking error dialog.
    Dim shapeEl As ShapeElement
    Set shapeEl = CreateShapeElement1(Nothing, vpts, 0)
    shapeEl.Color = 6
    shapeEl.LineWeight = 2
    On Error Resume Next
    shapeEl.Level = ActiveDesignFile.Levels("TWZWS2_P")
    On Error GoTo WsError
    ActiveModelReference.AddElement shapeEl
    shapeEl.Rewrite

    ' Spacing 2 ft so diagonal stripes read in a typical 12-ft-tall lane
    ' work space (spacing 10 only showed ~1 line and looked solid).
    ' Associative SetPattern alone is invisible when the view's Patterns
    ' attribute is off (live 2026-08-03) — also draw real stripe lines.
    Dim hatchNote As String
    hatchNote = ApplyHatchPatternToClosed(shapeEl, 2#, 45#)
    Dim lineNote As String
    lineNote = DrawDiagonalHatchLines(shapeEl, 2#, 45#)

    ExecPlaceWorkspace = "OK" & vbTab & "vertexCount=" & n & vbTab & _
                         "elementId=" & CStr(ElIDAsDouble(shapeEl.ID)) & vbTab & _
                         "note=placed work space shape (unfilled); " & hatchNote & "; " & lineNote
    Exit Function

WsError:
    ExecPlaceWorkspace = "ERROR" & vbTab & "note=" & Err.Description
End Function

' Draw visible 45-deg hatch stripes clipped to shape bbox (WYSIWYG even
' when view Patterns attribute is off). spacing = perpendicular distance.
Private Function DrawDiagonalHatchLines(shapeEl As Element, spacing As Double, angleDeg As Double) As String
    On Error GoTo DrawErr
    Dim rng As Range3d
    rng = shapeEl.Range
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
    xmin = rng.Low.X: xmax = rng.High.X
    ymin = rng.Low.Y: ymax = rng.High.Y

    Dim stepC As Double
    stepC = spacing * Sqr(2#)
    Dim cMin As Double, cMax As Double, cVal As Double
    cMin = xmin + ymin
    cMax = xmax + ymax

    Dim nLines As Integer: nLines = 0
    Dim lvl As Level
    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels("TWZWS2_P")
    On Error GoTo DrawErr

    cVal = cMin
    Do While cVal <= cMax + 0.001
        Dim ptsX(0 To 3) As Double, ptsY(0 To 3) As Double
        Dim nPts As Integer: nPts = 0
        Dim yL As Double, yR As Double, xB As Double, xT As Double
        yL = cVal - xmin
        If yL >= ymin - 0.0000001 And yL <= ymax + 0.0000001 Then
            ptsX(nPts) = xmin: ptsY(nPts) = yL: nPts = nPts + 1
        End If
        yR = cVal - xmax
        If yR >= ymin - 0.0000001 And yR <= ymax + 0.0000001 Then
            ptsX(nPts) = xmax: ptsY(nPts) = yR: nPts = nPts + 1
        End If
        xB = cVal - ymin
        If xB >= xmin - 0.0000001 And xB <= xmax + 0.0000001 Then
            ptsX(nPts) = xB: ptsY(nPts) = ymin: nPts = nPts + 1
        End If
        xT = cVal - ymax
        If xT >= xmin - 0.0000001 And xT <= xmax + 0.0000001 Then
            ptsX(nPts) = xT: ptsY(nPts) = ymax: nPts = nPts + 1
        End If

        If nPts >= 2 Then
            Dim i As Integer, j As Integer, keep As Integer
            Dim ux(0 To 3) As Double, uy(0 To 3) As Double
            keep = 0
            For i = 0 To nPts - 1
                Dim dup As Boolean: dup = False
                For j = 0 To keep - 1
                    If Abs(ptsX(i) - ux(j)) < 0.000001 And Abs(ptsY(i) - uy(j)) < 0.000001 Then
                        dup = True: Exit For
                    End If
                Next j
                If Not dup Then
                    ux(keep) = ptsX(i): uy(keep) = ptsY(i): keep = keep + 1
                End If
            Next i
            If keep >= 2 Then
                Dim p1 As Point3d, p2 As Point3d
                p1.X = ux(0): p1.Y = uy(0): p1.Z = 0
                p2.X = ux(1): p2.Y = uy(1): p2.Z = 0
                Dim ln As LineElement
                Set ln = CreateLineElement2(Nothing, p1, p2)
                ln.Color = shapeEl.Color
                ln.LineWeight = shapeEl.LineWeight
                If Not lvl Is Nothing Then ln.Level = lvl
                ActiveModelReference.AddElement ln
                ln.Rewrite
                nLines = nLines + 1
            End If
        End If
        cVal = cVal + stepC
    Loop

    DrawDiagonalHatchLines = "drew " & nLines & " hatch lines (spacing=" & spacing & ")"
    Exit Function
DrawErr:
    DrawDiagonalHatchLines = "ERROR:" & Err.Description
End Function

' ============================================================
' HATCH A CLOSED ELEMENT BY ID — CreateHatchPattern1 + SetPattern
' with Matrix3dIdentity. Live-proven on DELETE.dgn (HasPattern=True).
' spacing in master units; angleDeg in degrees (converted to radians).
' ============================================================
Public Function ExecHatchClosedElementByID(elementId As Double, _
                                           Optional spacing As Double = 10#, _
                                           Optional angleDeg As Double = 45#) As String
    On Error GoTo HatchErr

    If spacing <= 0 Then
        ExecHatchClosedElementByID = "ERROR" & vbTab & "note=spacing must be > 0"
        Exit Function
    End If

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecHatchClosedElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim note As String
    note = ApplyHatchPatternToClosed(el, spacing, angleDeg)
    If Left(note, 5) = "ERROR" Then
        ExecHatchClosedElementByID = "ERROR" & vbTab & "note=" & Mid(note, 7)
        Exit Function
    End If

    ExecHatchClosedElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                 "spacing=" & spacing & vbTab & "angleDeg=" & angleDeg & vbTab & _
                                 "note=" & note
    Exit Function

HatchErr:
    ExecHatchClosedElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE ARC (3-point / placeArcModeEx=3) — sequence from
' LegacySignPlace.bas / DrawSign.bas, live-verified +1 element on
' DELETE.dgn 2026-08-02. Point order: start, end, bulge.
' ============================================================
Public Function ExecPlaceArc3Point(x1 As Double, y1 As Double, _
                                   x2 As Double, y2 As Double, _
                                   x3 As Double, y3 As Double, _
                                   Optional z As Double = 0) As String
    On Error GoTo ArcErr

    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"

    Dim pt As Point3d
    pt.X = x1: pt.Y = y1: pt.Z = z
    CadInputQueue.SendDataPoint pt, 1
    pt.X = x2: pt.Y = y2: pt.Z = z
    CadInputQueue.SendDataPoint pt, 1
    pt.X = x3: pt.Y = y3: pt.Z = z
    CadInputQueue.SendDataPoint pt, 1

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    ExecPlaceArc3Point = "OK" & vbTab & _
                         "x1=" & x1 & vbTab & "y1=" & y1 & vbTab & _
                         "x2=" & x2 & vbTab & "y2=" & y2 & vbTab & _
                         "x3=" & x3 & vbTab & "y3=" & y3 & vbTab & _
                         "note=placed 3-point arc (placeArcModeEx=3)"
    Exit Function

ArcErr:
    ExecPlaceArc3Point = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE TEXT LABEL via TEXTEDITOR PLACE + INSERT_TEXT — same
' path as LegacySignPlace / LegacyPrototype (single line). Live
' verified +1 element on DELETE.dgn 2026-08-02.
' ============================================================
Public Function ExecPlaceTextLabel(text As String, x As Double, y As Double, _
                                   Optional z As Double = 0) As String
    On Error GoTo TxtErr

    If Trim(text) = "" Then
        ExecPlaceTextLabel = "ERROR" & vbTab & "note=text is empty"
        Exit Function
    End If

    ' Element API text, then Move so the glyph range is centered on (x,y).
    ' Justification enum alone does not change origin vs range on this install.
    Dim pt As Point3d
    pt.X = x: pt.Y = y: pt.Z = z
    Dim te As TextElement
    Set te = CreateTextElement1(Nothing, text, pt, Matrix3dIdentity)
    te.Color = 0
    te.LineWeight = 0
    ActiveModelReference.AddElement te
    te.Rewrite

    Dim rng As Range3d
    rng = te.Range
    Dim cx As Double, cy As Double
    cx = 0.5 * (rng.Low.X + rng.High.X)
    cy = 0.5 * (rng.Low.Y + rng.High.Y)
    Dim delta As Point3d
    delta.X = x - cx: delta.Y = y - cy: delta.Z = 0
    te.Move delta
    te.Rewrite

    ExecPlaceTextLabel = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(te.ID)) & vbTab & _
                         "x=" & x & vbTab & "y=" & y & vbTab & _
                         "text=" & text & vbTab & "note=placed text label (center)"
    Exit Function

TxtErr:
    ExecPlaceTextLabel = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' PLACE LINEAR DIMENSION — real DimensionElement (Linear Size /
' msdDimTypeSizeArrow), matching PlaceElements.frm
' "DIMENSION SIZE WITH LINES". CadInputQueue datapoints create
' 0 elements headlessly; CreateDimensionElement1 works.
' styleName: DGN DimensionStyles name (default ny_Plan = active
' NYSDOT plan style). DimHeight = offset of dim line from segment.
' ============================================================
Public Function ExecPlaceDimension(x1 As Double, y1 As Double, _
                                   x2 As Double, y2 As Double, _
                                   ox As Double, oy As Double, _
                                   Optional z As Double = 0, _
                                   Optional styleName As String = "ny_Plan") As String
    On Error GoTo DimErr

    Dim dx As Double, dy As Double, L As Double
    dx = x2 - x1: dy = y2 - y1
    L = Sqr(dx * dx + dy * dy)
    If L < 0.001 Then
        ExecPlaceDimension = "ERROR" & vbTab & "note=dimension length too small"
        Exit Function
    End If

    Dim ux As Double, uy As Double
    ux = dx / L: uy = dy / L
    Dim px As Double, py As Double
    px = -uy: py = ux

    Dim mx As Double, my As Double
    mx = 0.5 * (x1 + x2): my = 0.5 * (y1 + y2)
    Dim side As Double
    If ((ox - mx) * px + (oy - my) * py) >= 0 Then side = 1# Else side = -1#
    Dim off As Double
    off = Abs((ox - mx) * px + (oy - my) * py)
    If off < 0.01 Then off = 30#

    ' Rotation: X along measured segment so text/terminators follow alignment
    Dim rot As Matrix3d
    rot = Matrix3dFromRowValues(ux, uy, 0#, _
                                -uy, ux, 0#, _
                                0#, 0#, 1#)

    Dim oDim As DimensionElement
    Set oDim = CreateDimensionElement1(Nothing, rot, msdDimTypeSizeArrow)

    Dim p1 As Point3d, p2 As Point3d
    p1.X = x1: p1.Y = y1: p1.Z = z
    p2.X = x2: p2.Y = y2: p2.Z = z
    oDim.AddReferencePoint ActiveModelReference, p1
    oDim.AddReferencePoint ActiveModelReference, p2
    oDim.DimHeight = side * off

    ' ny_Plan has ShowSecondaryText=True (primary above + secondary
    ' below). Force one measurement only for the placement copy —
    ' do not OverrideTextHeight here (annotation scale would explode it).
    On Error Resume Next
    Dim oStyle As DimensionStyle
    Dim styleErrNum As Long: styleErrNum = 0
    Dim styleErrDesc As String: styleErrDesc = ""
    Set oStyle = ActiveDesignFile.DimensionStyles(styleName)
    If Err.Number <> 0 Then
        styleErrNum = Err.Number
        styleErrDesc = Err.Description
        Err.Clear
    End If
    Dim savedSec As Boolean
    Dim styleApplied As Boolean: styleApplied = False
    If Not oStyle Is Nothing Then
        savedSec = oStyle.ShowSecondaryText
        oStyle.ShowSecondaryText = False
        Set oDim.DimensionStyle = oStyle
        oStyle.ShowSecondaryText = savedSec
        If Err.Number = 0 Then styleApplied = True Else styleErrNum = Err.Number: styleErrDesc = Err.Description
    End If
    On Error GoTo DimErr

    oDim.Color = 2
    oDim.LineWeight = 0
    ActiveModelReference.AddElement oDim
    oDim.Rewrite

    If Not oDim.IsDimensionElement Then
        ExecPlaceDimension = "ERROR" & vbTab & "note=created element is not a DimensionElement"
        Exit Function
    End If

    ' DIAGNOSTIC (temporary): styleApplied/styleErr* reveal whether the
    ' named DimensionStyle actually resolved -- the On Error Resume Next
    ' above previously swallowed a failed lookup silently, so a dimension
    ' could render with no style (no arrow terminators/text) with no error
    ' ever surfacing. Remove once root-caused.
    ExecPlaceDimension = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(oDim.ID)) & vbTab & _
                         "lengthFt=" & Format(L, "0.0") & vbTab & _
                         "style=" & styleName & vbTab & _
                         "dimType=SizeArrow" & vbTab & _
                         "styleApplied=" & IIf(styleApplied, "Y", "N") & vbTab & _
                         "styleErrNum=" & styleErrNum & vbTab & _
                         "styleErrDesc=" & styleErrDesc & vbTab & _
                         "note=real DimensionElement (ny_Plan Linear Size)"
    Exit Function

DimErr:
    ExecPlaceDimension = "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function FindNewestElementIdAbove(beforeMaxID As Double) As Double
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Dim best As Double: best = beforeMaxID
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > best Then best = idVal
    Loop
    FindNewestElementIdAbove = best
End Function

' Returns a short note string; prefix "ERROR:" on failure.
Private Function ApplyHatchPatternToClosed(el As Element, spacing As Double, angleDeg As Double) As String
    On Error GoTo ApplyErr

    Dim closed As ClosedElement
    Set closed = el.AsClosedElement
    If closed Is Nothing Then
        ApplyHatchPatternToClosed = "ERROR:element is not a closed shape"
        Exit Function
    End If

    Dim angRad As Double
    angRad = angleDeg * Atn(1) * 4# / 180#

    Dim hatch As HatchPattern
    Set hatch = CreateHatchPattern1(spacing, angRad)
    hatch.Color = el.Color
    hatch.LineWeight = el.LineWeight

    closed.SetPattern hatch, Matrix3dIdentity
    el.Rewrite

    If closed.HasPattern Then
        ApplyHatchPatternToClosed = "hatch applied (HasPattern=True spacing=" & spacing & " angleDeg=" & angleDeg & ")"
    Else
        ApplyHatchPatternToClosed = "ERROR:SetPattern completed but HasPattern=False"
    End If
    Exit Function

ApplyErr:
    ApplyHatchPatternToClosed = "ERROR:" & Err.Description
End Function

Private Function ScanMaxElementID() As Double
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Dim maxID As Double: maxID = 0
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > maxID Then maxID = idVal
    Loop
    ScanMaxElementID = maxID
End Function

Private Function FindNewestClosedElementAbove(beforeMaxID As Double) As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim best As Element
    Dim bestID As Double: bestID = beforeMaxID
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > bestID Then
            On Error Resume Next
            Dim c As ClosedElement
            Set c = el.AsClosedElement
            If Err.Number = 0 And Not c Is Nothing Then
                Set best = el
                bestID = idVal
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Loop
    Set FindNewestClosedElementAbove = best
End Function

' ============================================================
' FIND AN INTERIOR POINT OF A SIMPLE POLYGON
' Horizontal scanline through the vertical bbox midpoint; collects
' all edge crossings, sorts them, and returns the midpoint of the
' WIDEST even-odd interior span. Correct for non-convex shapes
' (L-shaped work areas, etc.) where a plain vertex-average centroid
' can land outside the polygon. If the first scanline is degenerate
' (hits a vertex exactly, odd crossing count), nudges Y and retries
' a few times before giving up.
' ============================================================
Private Function FindInteriorPoint(vx() As Double, vy() As Double, n As Integer, _
                                   ByRef outX As Double, ByRef outY As Double) As Boolean
    Dim minY As Double, maxY As Double
    Dim i As Integer
    minY = vy(0): maxY = vy(0)
    For i = 1 To n - 1
        If vy(i) < minY Then minY = vy(i)
        If vy(i) > maxY Then maxY = vy(i)
    Next i

    Dim scanY As Double
    scanY = (minY + maxY) / 2#

    ' Allocated once at the max possible size (an n-vertex polygon can have
    ' at most n scanline crossings) and never resized again. An earlier
    ' version resized this array by one element per crossing found (ReDim
    ' Preserve in the loop) and hit an unexplained "Subscript out of range"
    ' once 2+ crossings accumulated. A single plain ReDim up front avoids
    ' that pattern entirely. Must be declared dynamic (empty parens) --
    ' a static-size Dim here cannot be ReDim'd at all, which is what the
    ' "array already dimensioned" compile error was.
    Dim maxCrossings As Integer
    If n > 50 Then
        maxCrossings = n
    Else
        maxCrossings = 50
    End If
    Dim xs() As Double
    ReDim xs(1 To maxCrossings)

    Dim checkpoint As String: checkpoint = "xs sized 1 To " & maxCrossings & " (UBound=" & UBound(xs) & ")"
    On Error GoTo FipError

    Dim attempt As Integer
    For attempt = 0 To 4
        Dim xCount As Integer: xCount = 0
        Dim j As Integer

        For i = 0 To n - 1
            j = (i + 1) Mod n
            Dim y1 As Double, y2 As Double, x1 As Double, x2 As Double
            y1 = vy(i): y2 = vy(j): x1 = vx(i): x2 = vx(j)
            If (y1 <= scanY And y2 > scanY) Or (y2 <= scanY And y1 > scanY) Then
                Dim t As Double
                t = (scanY - y1) / (y2 - y1)
                xCount = xCount + 1
                checkpoint = "attempt " & attempt & " storing crossing " & xCount & _
                    " (xs UBound=" & UBound(xs) & ")"
                xs(xCount) = x1 + t * (x2 - x1)
            End If
        Next i

        If xCount >= 2 And xCount Mod 2 = 0 Then
            Dim a As Integer, b As Integer, tmp As Double
            For a = 2 To xCount
                checkpoint = "attempt " & attempt & " sort a=" & a & " (xs UBound=" & UBound(xs) & ")"
                tmp = xs(a)
                b = a - 1
                Do While b >= 1 And xs(b) > tmp
                    xs(b + 1) = xs(b): b = b - 1
                Loop
                xs(b + 1) = tmp
            Next a

            Dim bestWidth As Double: bestWidth = -1
            Dim bestMidX As Double
            Dim k As Integer
            For k = 1 To xCount - 1 Step 2
                checkpoint = "attempt " & attempt & " widest k=" & k & " (xs UBound=" & UBound(xs) & ")"
                Dim w As Double: w = xs(k + 1) - xs(k)
                If w > bestWidth Then
                    bestWidth = w
                    bestMidX = (xs(k) + xs(k + 1)) / 2#
                End If
            Next k

            outX = bestMidX
            outY = scanY
            FindInteriorPoint = True
            Exit Function
        End If

        scanY = scanY + (maxY - minY) * 0.0037 * (attempt + 1)
    Next attempt

    FindInteriorPoint = False
    Exit Function

FipError:
    Err.Raise Err.Number, , "[in FindInteriorPoint at: " & checkpoint & "] " & Err.Description
End Function

' ============================================================
' SET SIGN DISPLAY ATTRIBUTES ON ALREADY-PLACED ELEMENTS
' Replaces the 22-keyin CHANGE ATTRIBUTES relay in
' frmSignSubColors.frm with direct element property assignment --
' no clicking, no GetInput. Finds targets via full model scan
' filtered by element ID, the same proven pattern FindElementsNear
' and PerpPlacement.DeletePerpLines already use (no direct
' get-by-ID call is used anywhere in this codebase to reuse).
'
' Only Level/Color/LineWeight are set -- the three properties
' confirmed as directly assignable Element members elsewhere in
' this codebase (PerpPlacement.bas: lineEl.Color, .LineWeight,
' .Level). The original CHANGE ATTRIBUTES sequence also sets
' FillColor=6, ElementClass=CONSTRUCTION, LineStyle=ByLevel,
' Transparency=0, Priority=0 -- none of those have a confirmed VBA
' property path anywhere in this repo, and guessing risks the same
' category of compile error as msdElementTypeCell / Level.Color
' earlier. Most likely to matter in practice: ElementClass=
' CONSTRUCTION affects print/plot visibility -- flagged as a known
' gap, not silently dropped.
'
' elementIdsTSV: comma-separated element IDs (as produced by
' FindElementsNear's elementId column). Typical caller pattern:
' FIND_ELEMENTS_NEAR with typeFilter=CELL near the sign location,
' filter results for the sign face cell's name, pass those IDs here.
' ============================================================
Public Function ExecSetSignAttributes(elementIdsTSV As String) As String
    On Error GoTo AttrError

    If Trim(elementIdsTSV) = "" Then
        ExecSetSignAttributes = "ERROR" & vbTab & "note=no element IDs supplied"
        Exit Function
    End If

    Dim idStrs() As String
    idStrs = Split(elementIdsTSV, ",")

    Dim targetIds As Object
    Set targetIds = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 0 To UBound(idStrs)
        Dim s As String: s = Trim(idStrs(i))
        If s <> "" Then targetIds(CDbl(s)) = True
    Next i

    Dim sfLevel As Level
    Set sfLevel = ActiveDesignFile.Levels("SF_P")

    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim applied As Integer: applied = 0
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If targetIds.Exists(idVal) Then
            ' Face cells must keep library symbology: ByCell weights and
            ' SF_P/SFB_P levels on subelements. Forcing Level=SF_P +
            ' LineWeight=3 on the cell (earlier) thickened every shape
            ' inside the diamond and dropped SFB_P — live mismatch vs
            ' engineer reference W20-01RA (2026-08-03).
            If el.IsTextElement Or el.IsTextNodeElement Then
                el.Level = sfLevel
                el.Color = 0
                el.LineWeight = 3
                el.Rewrite
                applied = applied + 1
            ElseIf el.IsLineElement Then
                el.Level = sfLevel
                el.Color = 0
                el.LineWeight = 0
                el.Rewrite
                applied = applied + 1
            ElseIf el.IsCellElement Then
                If UCase$(el.AsCellElement.Name) = "TWZSGN_P" Then
                    el.Color = 6
                    el.Rewrite
                    applied = applied + 1
                End If
                ' else: sign-face cell — leave Level/Color/Weight alone
            End If
        End If
    Loop

    ExecSetSignAttributes = "OK" & vbTab & "applied=" & applied & vbTab & _
                            "requested=" & (UBound(idStrs) + 1) & vbTab & _
                            "note=labels/stems SF_P white; post TWZSGN_P color=6; " & _
                            "face cells untouched (library weights/levels)"
    Exit Function

AttrError:
    ExecSetSignAttributes = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' DELETE ELEMENTS BY ID
' Same scan-and-match-and-RemoveElement pattern as
' PerpPlacement.DeletePerpLines -- the only established precedent
' in this codebase for deleting a specific element by ID (there is
' no direct GetElementByID call anywhere to reuse instead).
'
' Built for WZTCBridge's UNDO_LAST_OP: undo is implemented as
' "delete the exact elements the last op created" rather than by
' relying on MicroStation's own undo stack, since the stack's
' grouping behavior across a multi-element op (e.g. PLACE_SIGN's
' post + face + text + arc) has not been verified in the IDE.
'
' idsCSV: comma-separated element IDs, e.g. "88213,88214,88215"
' ============================================================
Public Function ExecDeleteElementsByID(idsCSV As String) As String
    On Error GoTo DelError

    If Trim(idsCSV) = "" Then
        ExecDeleteElementsByID = "OK" & vbTab & "deleted=0" & vbTab & "note=no element IDs given"
        Exit Function
    End If

    Dim idStrs() As String
    idStrs = Split(idsCSV, ",")

    Dim targetIds As Object
    Set targetIds = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 0 To UBound(idStrs)
        Dim s As String: s = Trim(idStrs(i))
        If s <> "" Then targetIds(CDbl(s)) = True
    Next i

    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim toDelete() As Element
    Dim nDelete As Integer: nDelete = 0
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If targetIds.Exists(ElIDAsDouble(el.ID)) Then
            nDelete = nDelete + 1
            ReDim Preserve toDelete(1 To nDelete)
            Set toDelete(nDelete) = el
        End If
    Loop

    For i = 1 To nDelete
        ActiveModelReference.RemoveElement toDelete(i)
    Next i

    ExecDeleteElementsByID = "OK" & vbTab & "deleted=" & nDelete & vbTab & _
                            "requested=" & (UBound(idStrs) + 1)
    Exit Function

DelError:
    ExecDeleteElementsByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' MOVE ELEMENT BY ID (M6)
' Direct Element.Move + Rewrite — same pattern
' BBMarkupProcessor.ExecuteMove already proves works, minus its
' interactive-click fallback. deltaX/deltaY/deltaZ in design
' units (ft). Returns priorDelta* as the reverse displacement
' so UNDO_LAST_OP can re-apply it without needing a snapshot.
' ============================================================
Public Function ExecMoveElementByID(elementId As Double, _
                                    deltaX As Double, deltaY As Double, _
                                    Optional deltaZ As Double = 0) As String
    On Error GoTo MoveError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecMoveElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim delta As Point3d
    delta.X = deltaX: delta.Y = deltaY: delta.Z = deltaZ
    el.Move delta
    el.Rewrite

    ExecMoveElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                          "deltaX=" & deltaX & vbTab & "deltaY=" & deltaY & vbTab & "deltaZ=" & deltaZ & vbTab & _
                          "priorDeltaX=" & (-deltaX) & vbTab & "priorDeltaY=" & (-deltaY) & vbTab & _
                          "priorDeltaZ=" & (-deltaZ) & vbTab & _
                          "note=moved element " & elementId
    Exit Function

MoveError:
    ExecMoveElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' CHANGE ELEMENT LEVEL BY ID (M6)
' Direct el.Level + Rewrite — same approach as
' BBMarkupProcessor.ExecuteChangeLevel / ExecSetSignAttributes,
' without On Error Resume Next around the mutation.
'
' Reading el.Level.Name back on a freshly re-scanned element is a
' confirmed-flaky read on this install (runtime error 91) -- already
' hit and guarded the same way in WZTCQuery.FindElementsNear (:120)
' and flagged in Debug/DebugExecTest.bas:224-233. Confirmed live
' 2026-07-31 via the actual MCP path: this crashed on every call
' before the guard below was added. Writing el.Level is proven safe
' (PerpPlacement.bas, BBMarkupProcessor); reading it back is not.
' ============================================================
Public Function ExecChangeElementLevelByID(elementId As Double, levelName As String) As String
    On Error GoTo LevelError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecChangeElementLevelByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim priorLevel As String: priorLevel = ""
    Dim havePriorLevel As Boolean: havePriorLevel = False
    On Error Resume Next
    priorLevel = el.Level.Name
    If Err.Number = 0 Then havePriorLevel = True
    Err.Clear
    On Error GoTo LevelError

    Dim lvl As Level
    Set lvl = Nothing
    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels(levelName)
    On Error GoTo LevelError
    If lvl Is Nothing Then
        ExecChangeElementLevelByID = "ERROR" & vbTab & "note=level not found: " & levelName
        Exit Function
    End If

    el.Level = lvl
    el.Rewrite

    If havePriorLevel Then
        ExecChangeElementLevelByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                     "level=" & levelName & vbTab & "priorLevel=" & priorLevel & vbTab & _
                                     "note=changed level " & priorLevel & " -> " & levelName
    Else
        ' Prior level unreadable on this install -- declare not-undoable
        ' rather than emit a bogus/empty priorLevel= that UNDO_LAST_OP's
        ' field scan would otherwise misread as "restore to blank level".
        ExecChangeElementLevelByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                     "level=" & levelName & vbTab & "notUndoable=Y" & vbTab & _
                                     "note=changed level to " & levelName & _
                                     " (prior level unreadable on this install -- not undoable via UNDO_LAST_OP)"
    End If
    Exit Function

LevelError:
    ExecChangeElementLevelByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' EDIT TEXT ELEMENT BY ID (M6)
' Handles TextElement and TextNodeElement (first line), matching
' BBMarkupProcessor.ExecuteEditText. Returns priorText= so
' UNDO_LAST_OP can restore the previous string.
'
' priorText reads are guarded the same way ExecChangeElementLevelByID
' guards el.Level.Name -- reading el.Level back on a freshly re-scanned
' element is confirmed-flaky on this install (runtime error 91; see
' that function's header), and .Text hasn't been proven safe to read
' back either, so the same defensive pattern applies here on general
' principle rather than waiting to hit it live.
' ============================================================
Public Function ExecEditTextByID(elementId As Double, newText As String) As String
    On Error GoTo TextError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecEditTextByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim priorText As String: priorText = ""
    Dim havePriorText As Boolean: havePriorText = False

    If el.Type = msdElementTypeText Then
        Dim te As TextElement
        Set te = el
        On Error Resume Next
        priorText = te.Text
        If Err.Number = 0 Then havePriorText = True
        Err.Clear
        On Error GoTo TextError
        te.Text = newText
        te.Rewrite
    ElseIf el.Type = msdElementTypeTextNode Then
        Dim tn As TextNodeElement
        Set tn = el
        Dim lineEl As TextElement
        Dim lineEnum As ElementEnumerator
        Set lineEnum = tn.GetSubElements
        If lineEnum.MoveNext Then
            Set lineEl = lineEnum.Current
            On Error Resume Next
            priorText = lineEl.Text
            If Err.Number = 0 Then havePriorText = True
            Err.Clear
            On Error GoTo TextError
            lineEl.Text = newText
            lineEl.Rewrite
        Else
            ExecEditTextByID = "ERROR" & vbTab & "note=text node has no sub-elements: " & elementId
            Exit Function
        End If
        tn.Rewrite
    Else
        ExecEditTextByID = "ERROR" & vbTab & "note=element is not text (type " & el.Type & ")"
        Exit Function
    End If

    If havePriorText Then
        ExecEditTextByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                           "newText=" & newText & vbTab & "priorText=" & priorText & vbTab & _
                           "note=edited text element " & elementId
    Else
        ExecEditTextByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                           "newText=" & newText & vbTab & "notUndoable=Y" & vbTab & _
                           "note=edited text element " & elementId & _
                           " (prior text unreadable on this install -- not undoable via UNDO_LAST_OP)"
    End If
    Exit Function

TextError:
    ExecEditTextByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' COPY ELEMENT BY ID (Phase C) — Clone + Move + AddElement.
' Live-proven on DELETE.dgn via COM (Element.Clone / .Move /
' ActiveModelReference.AddElement). Not a CadInputQueue recipe.
' ============================================================
Public Function ExecCopyElementByID(elementId As Double, _
                                    deltaX As Double, deltaY As Double, _
                                    Optional deltaZ As Double = 0) As String
    On Error GoTo CopyError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecCopyElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim c As Element
    Set c = el.Clone

    Dim delta As Point3d
    delta.X = deltaX: delta.Y = deltaY: delta.Z = deltaZ
    c.Move delta
    ActiveModelReference.AddElement c

    Dim newId As Double
    newId = ElIDAsDouble(c.ID)

    ExecCopyElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                          "newElementId=" & CStr(newId) & vbTab & _
                          "deltaX=" & deltaX & vbTab & "deltaY=" & deltaY & vbTab & "deltaZ=" & deltaZ & vbTab & _
                          "note=copied element " & elementId & " -> " & newId
    Exit Function

CopyError:
    ExecCopyElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' ROTATE ELEMENT BY ID (Phase C) — Matrix3d Z-axis rotation about
' a fixed point + Element.Transform. Live-proven on DELETE.dgn.
' angleDeg in degrees (converted to radians here).
' ============================================================
Public Function ExecRotateElementByID(elementId As Double, _
                                      originX As Double, originY As Double, _
                                      angleDeg As Double, _
                                      Optional originZ As Double = 0) As String
    On Error GoTo RotError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecRotateElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim origin As Point3d
    origin.X = originX: origin.Y = originY: origin.Z = originZ

    Dim angRad As Double
    angRad = angleDeg * Atn(1) * 4# / 180#

    Dim m As Matrix3d
    m = Matrix3dFromAxisAndRotationAngle(2, angRad)

    Dim t As Transform3d
    t = Transform3dFromMatrix3dAndFixedPoint3d(m, origin)

    el.Transform t
    el.Rewrite

    ExecRotateElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                            "originX=" & originX & vbTab & "originY=" & originY & vbTab & _
                            "angleDeg=" & angleDeg & vbTab & _
                            "priorAngleDeg=" & (-angleDeg) & vbTab & _
                            "note=rotated element " & elementId & " by " & angleDeg & " deg"
    Exit Function

RotError:
    ExecRotateElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' SCALE ELEMENT BY ID (Phase C) — Element.ScaleUniform about a
' point. Same API DrawSign.bas already uses on cell elements.
' ============================================================
Public Function ExecScaleElementByID(elementId As Double, _
                                     originX As Double, originY As Double, _
                                     scaleFactor As Double, _
                                     Optional originZ As Double = 0) As String
    On Error GoTo ScaleError

    If scaleFactor = 0 Then
        ExecScaleElementByID = "ERROR" & vbTab & "note=scaleFactor must be non-zero"
        Exit Function
    End If

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecScaleElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim origin As Point3d
    origin.X = originX: origin.Y = originY: origin.Z = originZ

    el.ScaleUniform origin, scaleFactor
    el.Rewrite

    ExecScaleElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                           "originX=" & originX & vbTab & "originY=" & originY & vbTab & _
                           "scaleFactor=" & scaleFactor & vbTab & _
                           "priorScaleFactor=" & (1# / scaleFactor) & vbTab & _
                           "note=scaled element " & elementId & " by " & scaleFactor
    Exit Function

ScaleError:
    ExecScaleElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' MIRROR ELEMENT BY ID (Phase C) — Element.Mirror about the axis
' through (x1,y1)-(x2,y2). Live-proven on DELETE.dgn (two-point form).
' ============================================================
Public Function ExecMirrorElementByID(elementId As Double, _
                                      x1 As Double, y1 As Double, _
                                      x2 As Double, y2 As Double, _
                                      Optional z1 As Double = 0, _
                                      Optional z2 As Double = 0) As String
    On Error GoTo MirrorError

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecMirrorElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim p1 As Point3d, p2 As Point3d
    p1.X = x1: p1.Y = y1: p1.Z = z1
    p2.X = x2: p2.Y = y2: p2.Z = z2

    el.Mirror p1, p2
    el.Rewrite

    ExecMirrorElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                            "x1=" & x1 & vbTab & "y1=" & y1 & vbTab & _
                            "x2=" & x2 & vbTab & "y2=" & y2 & vbTab & _
                            "note=mirrored element " & elementId & _
                            " (re-run same mirror to undo)"
    Exit Function

MirrorError:
    ExecMirrorElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' ARRAY ELEMENT BY ID (Phase C) — rectangular copies via repeated
' Clone+Move (same as ExecCopyElementByID). count = number of NEW
' copies (not including the original). spacing along X/Y in ft.
' ============================================================
Public Function ExecArrayElementByID(elementId As Double, _
                                     count As Integer, _
                                     spacingX As Double, spacingY As Double) As String
    On Error GoTo ArrayError

    If count < 1 Then
        ExecArrayElementByID = "ERROR" & vbTab & "note=count must be >= 1"
        Exit Function
    End If

    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecArrayElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    Dim newIds As String: newIds = ""
    Dim i As Integer
    For i = 1 To count
        Dim c As Element
        Set c = el.Clone
        Dim delta As Point3d
        delta.X = spacingX * i: delta.Y = spacingY * i: delta.Z = 0
        c.Move delta
        ActiveModelReference.AddElement c
        If newIds <> "" Then newIds = newIds & ","
        newIds = newIds & CStr(ElIDAsDouble(c.ID))
    Next i

    ExecArrayElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                           "count=" & count & vbTab & _
                           "spacingX=" & spacingX & vbTab & "spacingY=" & spacingY & vbTab & _
                           "newElementIds=" & newIds & vbTab & _
                           "note=arrayed " & count & " copies of " & elementId
    Exit Function

ArrayError:
    ExecArrayElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' GENERAL GEOMETRY (Tier 1-3) — Element API / Fence COM.
' Prefer Create*Element + Rewrite over CadInputQueue tools.
' ============================================================

Public Function ExecPlaceCircle(cx As Double, cy As Double, radius As Double, _
                                Optional z As Double = 0) As String
    On Error GoTo E
    If radius <= 0 Then
        ExecPlaceCircle = "ERROR" & vbTab & "note=radius must be > 0"
        Exit Function
    End If
    Dim origin As Point3d
    origin.X = cx: origin.Y = cy: origin.Z = z
    Dim el As EllipseElement
    Set el = CreateEllipseElement2(Nothing, origin, radius, radius, Matrix3dIdentity)
    el.Color = ActiveSettings.Color
    el.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement el
    el.Rewrite
    ExecPlaceCircle = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(el.ID)) & vbTab & _
                      "cx=" & cx & vbTab & "cy=" & cy & vbTab & "radius=" & radius & vbTab & _
                      "note=placed circle (CreateEllipseElement2)"
    Exit Function
E:
    ExecPlaceCircle = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecPlaceEllipse(cx As Double, cy As Double, _
                                 primaryRadius As Double, secondaryRadius As Double, _
                                 Optional angleDeg As Double = 0, Optional z As Double = 0) As String
    On Error GoTo E
    If primaryRadius <= 0 Or secondaryRadius <= 0 Then
        ExecPlaceEllipse = "ERROR" & vbTab & "note=radii must be > 0"
        Exit Function
    End If
    Dim origin As Point3d
    origin.X = cx: origin.Y = cy: origin.Z = z
    Dim rot As Matrix3d
    If angleDeg = 0 Then
        rot = Matrix3dIdentity
    Else
        rot = Matrix3dFromAxisAndRotationAngle(2, angleDeg * Atn(1) * 4# / 180#)
    End If
    Dim el As EllipseElement
    Set el = CreateEllipseElement2(Nothing, origin, primaryRadius, secondaryRadius, rot)
    el.Color = ActiveSettings.Color
    el.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement el
    el.Rewrite
    ExecPlaceEllipse = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(el.ID)) & vbTab & _
                       "note=placed ellipse"
    Exit Function
E:
    ExecPlaceEllipse = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecPlaceBlock(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                               Optional z As Double = 0) As String
    On Error GoTo E
    Dim pts(0 To 3) As Point3d
    pts(0).X = x1: pts(0).Y = y1: pts(0).Z = z
    pts(1).X = x2: pts(1).Y = y1: pts(1).Z = z
    pts(2).X = x2: pts(2).Y = y2: pts(2).Z = z
    pts(3).X = x1: pts(3).Y = y2: pts(3).Z = z
    Dim el As ShapeElement
    Set el = CreateShapeElement1(Nothing, pts)
    el.Color = ActiveSettings.Color
    el.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement el
    el.Rewrite
    ExecPlaceBlock = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(el.ID)) & vbTab & _
                     "note=placed block/rectangle"
    Exit Function
E:
    ExecPlaceBlock = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecPlacePolyline(verticesTSV As String) As String
    On Error GoTo E
    Dim pts() As Point3d
    Dim n As Integer
    n = ParseVerticesTSV(verticesTSV, pts)
    If n < 2 Then
        ExecPlacePolyline = "ERROR" & vbTab & "note=polyline needs at least 2 vertices"
        Exit Function
    End If
    Dim el As LineElement
    Set el = CreateLineElement1(Nothing, pts)
    el.Color = ActiveSettings.Color
    el.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement el
    el.Rewrite
    ExecPlacePolyline = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(el.ID)) & vbTab & _
                        "vertexCount=" & n & vbTab & "note=placed polyline"
    Exit Function
E:
    ExecPlacePolyline = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecPlacePolygon(cx As Double, cy As Double, radius As Double, sides As Integer, _
                                 Optional z As Double = 0) As String
    On Error GoTo E
    If radius <= 0 Or sides < 3 Then
        ExecPlacePolygon = "ERROR" & vbTab & "note=need radius>0 and sides>=3"
        Exit Function
    End If
    Dim pts() As Point3d
    ReDim pts(0 To sides - 1)
    Dim i As Integer
    Dim twoPi As Double: twoPi = Atn(1) * 8#
    For i = 0 To sides - 1
        Dim ang As Double: ang = twoPi * i / sides
        pts(i).X = cx + radius * Cos(ang)
        pts(i).Y = cy + radius * Sin(ang)
        pts(i).Z = z
    Next i
    Dim el As ShapeElement
    Set el = CreateShapeElement1(Nothing, pts)
    el.Color = ActiveSettings.Color
    el.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement el
    el.Rewrite
    ExecPlacePolygon = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(el.ID)) & vbTab & _
                       "sides=" & sides & vbTab & "note=placed regular polygon"
    Exit Function
E:
    ExecPlacePolygon = "ERROR" & vbTab & "note=" & Err.Description
End Function

' Creates a chain of 2-point line elements from vertices, Default
' level / color 0 / weight 0 (CLAUDE.md's alignment-line property row)
' -- the same shape AlignmentTool.StartLineSegment's interactive
' clicks already produce (one line element per consecutive vertex
' pair), so the existing BuildAlignmentPath/CommitCurrentAlignment
' machinery (PerpPlacement.bas, AlignmentTool.bas) consumes it
' identically regardless of whether a human clicked or an agent
' supplied vertices. Straight segments only -- no arc support; nothing
' in this plan's vertex sources (FindReferenceLinework, click-picked
' points) currently produces arc data, so that's deferred until an
' actual caller needs it, not built speculatively.
' verticesTSV is "x,y,z|x,y,z|..." (>= 2 points), the same convention
' place_workspace/place_polyline already use.
Public Function ExecDefineAlignmentSegments(verticesTSV As String) As String
    On Error GoTo E
    Dim parts() As String
    parts = Split(verticesTSV, "|")
    If UBound(parts) < 1 Then
        ExecDefineAlignmentSegments = "ERROR" & vbTab & "note=need at least 2 vertices"
        Exit Function
    End If

    Dim pts() As Point3d
    ReDim pts(0 To UBound(parts))
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim xyz() As String: xyz = Split(parts(i), ",")
        If UBound(xyz) < 2 Then
            ExecDefineAlignmentSegments = "ERROR" & vbTab & "note=malformed vertex: " & parts(i)
            Exit Function
        End If
        pts(i).X = CDbl(xyz(0)): pts(i).Y = CDbl(xyz(1)): pts(i).Z = CDbl(xyz(2))
    Next i

    Dim defLvl As Level
    On Error Resume Next
    Set defLvl = ActiveDesignFile.Levels("Default")
    On Error GoTo E

    Dim createdCount As Integer: createdCount = 0
    For i = 0 To UBound(pts) - 1
        Dim le As LineElement
        Set le = CreateLineElement2(Nothing, pts(i), pts(i + 1))
        le.Color = 0
        le.LineWeight = 0
        If Not defLvl Is Nothing Then le.Level = defLvl
        ActiveModelReference.AddElement le
        le.Rewrite
        createdCount = createdCount + 1
    Next i

    ExecDefineAlignmentSegments = "OK" & vbTab & "segmentCount=" & createdCount & vbTab & _
                                  "note=created " & createdCount & " line element(s)"
    Exit Function
E:
    ExecDefineAlignmentSegments = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecChangeElementSymbology(elementId As Double, _
                                           Optional color As Long = -1, _
                                           Optional weight As Long = -1, _
                                           Optional lineStyleIndex As Long = -999, _
                                           Optional lineStyleName As String = "") As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecChangeElementSymbology = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    Dim priorColor As Long: priorColor = el.Color
    Dim priorWeight As Long: priorWeight = el.LineWeight
    If color >= 0 Then el.Color = color
    If weight >= 0 Then el.LineWeight = weight
    ' Prefer exact Name key from resolve_line_style (e.g. "( Dashed )").
    ' LineStyles(Number) fails for custom styles — Number is not a lookup key.
    ' A bad line style must not abort the whole call: color/weight above are
    ' only saved by Rewrite below, so an unguarded lookup failure here would
    ' silently discard those too. Report the failure instead of hiding it.
    Dim lineStyleNote As String: lineStyleNote = ""
    If Len(Trim$(lineStyleName)) > 0 Then
        On Error Resume Next
        el.LineStyle = ActiveDesignFile.LineStyles(Trim$(lineStyleName))
        If Err.Number <> 0 Then
            lineStyleNote = vbTab & "lineStyleError=" & Err.Description
            Err.Clear
        End If
        On Error GoTo E
    ElseIf lineStyleIndex <> -999 Then
        On Error Resume Next
        el.LineStyle = ActiveDesignFile.LineStyles(lineStyleIndex)
        If Err.Number <> 0 Then
            lineStyleNote = vbTab & "lineStyleError=" & Err.Description
            Err.Clear
        End If
        On Error GoTo E
    End If
    el.Rewrite
    ExecChangeElementSymbology = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                 "priorColor=" & priorColor & vbTab & "priorWeight=" & priorWeight & vbTab & _
                                 "note=updated symbology" & lineStyleNote
    Exit Function
E:
    ExecChangeElementSymbology = "ERROR" & vbTab & "note=" & Err.Description
End Function

' Perpendicular offset copy of a LINE (type 3). distance>0 = left of
' start->end direction in XY; <0 = right.
Public Function ExecCopyParallelLineByID(elementId As Double, distance As Double) As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecCopyParallelLineByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    Dim le As LineElement
    Set le = el.AsLineElement
    If le Is Nothing Then
        ExecCopyParallelLineByID = "ERROR" & vbTab & "note=element is not a line"
        Exit Function
    End If
    Dim s As Point3d, ept As Point3d
    s = le.StartPoint: ept = le.EndPoint
    Dim dx As Double, dy As Double, lineLen As Double
    dx = ept.X - s.X: dy = ept.Y - s.Y
    lineLen = Sqr(dx * dx + dy * dy)
    If lineLen = 0 Then
        ExecCopyParallelLineByID = "ERROR" & vbTab & "note=zero-length line"
        Exit Function
    End If
    Dim ox As Double, oy As Double
    ox = -dy / lineLen * distance
    oy = dx / lineLen * distance
    Dim c As Element
    Set c = el.Clone()
    Dim delta As Point3d
    delta.X = ox: delta.Y = oy: delta.Z = 0
    c.Move delta
    ActiveModelReference.AddElement c
    ExecCopyParallelLineByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                               "newElementId=" & CStr(ElIDAsDouble(c.ID)) & vbTab & _
                               "distance=" & distance & vbTab & "note=copy-parallel of line"
    Exit Function
E:
    ExecCopyParallelLineByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecCrossHatchClosedElementByID(elementId As Double, _
                                                Optional spacing As Double = 10#, _
                                                Optional angleDeg As Double = 45#) As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecCrossHatchClosedElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    Dim closed As ClosedElement
    Set closed = el.AsClosedElement
    If closed Is Nothing Then
        ExecCrossHatchClosedElementByID = "ERROR" & vbTab & "note=not a closed element"
        Exit Function
    End If
    Dim a1 As Double, a2 As Double
    a1 = angleDeg * Atn(1) * 4# / 180#
    a2 = a1 + Atn(1) * 2#   ' +90 deg
    Dim pat As Object
    Set pat = CreateCrossHatchPattern(spacing, spacing, a1, a2)
    On Error Resume Next
    pat.Color = el.Color
    pat.LineWeight = el.LineWeight
    On Error GoTo E
    closed.SetPattern pat, Matrix3dIdentity
    el.Rewrite
    ExecCrossHatchClosedElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                      "HasPattern=" & IIf(closed.HasPattern, "Y", "N") & vbTab & _
                                      "note=crosshatch applied"
    Exit Function
E:
    ExecCrossHatchClosedElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecRemoveHatchByID(elementId As Double) As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecRemoveHatchByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    Dim closed As ClosedElement
    Set closed = el.AsClosedElement
    If closed Is Nothing Then
        ExecRemoveHatchByID = "ERROR" & vbTab & "note=not a closed element"
        Exit Function
    End If
    If closed.HasPattern Then closed.RemovePattern
    el.Rewrite
    ExecRemoveHatchByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                          "HasPattern=" & IIf(closed.HasPattern, "Y", "N") & vbTab & _
                          "note=hatch removed"
    Exit Function
E:
    ExecRemoveHatchByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecBreakLineAtPoint(elementId As Double, x As Double, y As Double, _
                                     Optional z As Double = 0) As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecBreakLineAtPoint = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    Dim le As LineElement
    Set le = el.AsLineElement
    If le Is Nothing Then
        ExecBreakLineAtPoint = "ERROR" & vbTab & "note=element is not a line"
        Exit Function
    End If
    Dim s As Point3d, ept As Point3d, mid As Point3d
    s = le.StartPoint: ept = le.EndPoint
    mid.X = x: mid.Y = y: mid.Z = z
    Dim c1 As LineElement, c2 As LineElement
    Set c1 = CreateLineElement2(Nothing, s, mid)
    Set c2 = CreateLineElement2(Nothing, mid, ept)
    c1.Color = el.Color: c1.LineWeight = el.LineWeight
    c2.Color = el.Color: c2.LineWeight = el.LineWeight
    On Error Resume Next
    c1.Level = el.Level: c2.Level = el.Level
    On Error GoTo E
    ActiveModelReference.AddElement c1
    ActiveModelReference.AddElement c2
    c1.Rewrite: c2.Rewrite
    ActiveModelReference.RemoveElement el
    ExecBreakLineAtPoint = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                           "newElementIds=" & CStr(ElIDAsDouble(c1.ID)) & "," & CStr(ElIDAsDouble(c2.ID)) & vbTab & _
                           "note=broke line into two segments"
    Exit Function
E:
    ExecBreakLineAtPoint = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecExtendLineToLength(elementId As Double, newLength As Double) As String
    On Error GoTo E
    If newLength <= 0 Then
        ExecExtendLineToLength = "ERROR" & vbTab & "note=newLength must be > 0"
        Exit Function
    End If
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecExtendLineToLength = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If

    ' Prefer VertexList — LineElement.StartPoint/EndPoint + RemoveElement
    ' hung the VBA thread on this install during live verify.
    Dim vl As VertexList
    Set vl = el.AsVertexList
    If vl Is Nothing Then
        ExecExtendLineToLength = "ERROR" & vbTab & "note=element has no VertexList"
        Exit Function
    End If
    Dim verts() As Point3d
    verts = vl.GetVertices()
    If UBound(verts) < 1 Then
        ExecExtendLineToLength = "ERROR" & vbTab & "note=need at least 2 vertices"
        Exit Function
    End If
    Dim s As Point3d, ept As Point3d
    s = verts(LBound(verts))
    ept = verts(UBound(verts))
    Dim dx As Double, dy As Double, dz As Double, lineLen As Double
    dx = ept.X - s.X: dy = ept.Y - s.Y: dz = ept.Z - s.Z
    lineLen = Sqr(dx * dx + dy * dy + dz * dz)
    If lineLen = 0 Then
        ExecExtendLineToLength = "ERROR" & vbTab & "note=zero-length line"
        Exit Function
    End If
    Dim lenScale As Double: lenScale = newLength / lineLen
    Dim newEnd As Point3d
    newEnd.X = s.X + dx * lenScale
    newEnd.Y = s.Y + dy * lenScale
    newEnd.Z = s.Z + dz * lenScale

    Dim neu As LineElement
    Set neu = CreateLineElement2(Nothing, s, newEnd)
    neu.Color = el.Color
    neu.LineWeight = el.LineWeight
    On Error Resume Next
    neu.Level = el.Level
    On Error GoTo E
    ActiveModelReference.AddElement neu
    neu.Rewrite

    ' Delete original via proven scan-remove helper path
    Dim delNote As String
    delNote = ExecDeleteElementsByID(CStr(elementId))
    If Left(delNote, 2) <> "OK" Then
        ExecExtendLineToLength = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                                 "newElementId=" & CStr(ElIDAsDouble(neu.ID)) & vbTab & _
                                 "priorLength=" & lineLen & vbTab & "newLength=" & newLength & vbTab & _
                                 "note=extended but could not delete original: " & delNote
        Exit Function
    End If

    ExecExtendLineToLength = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                             "newElementId=" & CStr(ElIDAsDouble(neu.ID)) & vbTab & _
                             "priorLength=" & lineLen & vbTab & "newLength=" & newLength & vbTab & _
                             "note=recreated line at new length from start"
    Exit Function
E:
    ExecExtendLineToLength = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecFilletTwoElements(elementId1 As Double, elementId2 As Double, _
                                      radius As Double, pickX As Double, pickY As Double, _
                                      Optional pickZ As Double = 0) As String
    On Error GoTo E
    If radius <= 0 Then
        ExecFilletTwoElements = "ERROR" & vbTab & "note=radius must be > 0"
        Exit Function
    End If
    Dim el1 As Element, el2 As Element
    Set el1 = FindElementByID(elementId1)
    Set el2 = FindElementByID(elementId2)
    If el1 Is Nothing Or el2 Is Nothing Then
        ExecFilletTwoElements = "ERROR" & vbTab & "note=one or both elements not found"
        Exit Function
    End If
    Dim trav As TraversableElement
    Set trav = el1.AsTraversableElement
    If trav Is Nothing Then
        ExecFilletTwoElements = "ERROR" & vbTab & "note=element1 not traversable"
        Exit Function
    End If
    Dim pick As Point3d
    pick.X = pickX: pick.Y = pickY: pick.Z = pickZ
    Dim arc As ArcElement
    Set arc = trav.ConstructFillet(el2, Nothing, radius, pick, Matrix3dIdentity)
    If arc Is Nothing Then
        ExecFilletTwoElements = "ERROR" & vbTab & "note=ConstructFillet returned Nothing"
        Exit Function
    End If
    arc.Color = ActiveSettings.Color
    arc.LineWeight = ActiveSettings.LineWeight
    ActiveModelReference.AddElement arc
    arc.Rewrite
    ExecFilletTwoElements = "OK" & vbTab & "elementId1=" & CStr(elementId1) & vbTab & _
                            "elementId2=" & CStr(elementId2) & vbTab & _
                            "newElementId=" & CStr(ElIDAsDouble(arc.ID)) & vbTab & _
                            "note=fillet arc created (source lines not auto-trimmed)"
    Exit Function
E:
    ExecFilletTwoElements = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecCreateComplexString(elementIdsCSV As String) As String
    On Error GoTo E
    Dim idStrs() As String
    idStrs = Split(elementIdsCSV, ",")
    If UBound(idStrs) < 0 Then
        ExecCreateComplexString = "ERROR" & vbTab & "note=no element IDs"
        Exit Function
    End If

    ' CreateComplexStringElement1 requires ChainableElement(), not Element()
    ' (Bentley / geo-sol sample pattern).
    Dim chainables() As ChainableElement
    ReDim chainables(0 To UBound(idStrs))
    Dim i As Integer, n As Integer: n = 0
    For i = 0 To UBound(idStrs)
        Dim s As String: s = Trim(idStrs(i))
        If s <> "" Then
            Dim el As Element
            Set el = FindElementByID(CDbl(s))
            If el Is Nothing Then
                ExecCreateComplexString = "ERROR" & vbTab & "note=element not found: " & s
                Exit Function
            End If
            Dim ch As ChainableElement
            Set ch = el.AsChainableElement
            If ch Is Nothing Then
                ExecCreateComplexString = "ERROR" & vbTab & _
                    "note=element " & s & " is not chainable (need line/arc/curve/etc.)"
                Exit Function
            End If
            Set chainables(n) = ch
            n = n + 1
        End If
    Next i
    If n < 1 Then
        ExecCreateComplexString = "ERROR" & vbTab & "note=need at least 1 chainable element"
        Exit Function
    End If
    ReDim Preserve chainables(0 To n - 1)

    Dim cs As ComplexStringElement
    Set cs = CreateComplexStringElement1(chainables)
    ActiveModelReference.AddElement cs
    cs.Rewrite

    ExecCreateComplexString = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(cs.ID)) & vbTab & _
                              "partCount=" & n & vbTab & _
                              "note=complex string created (source elements left in place)"
    Exit Function
E:
    ExecCreateComplexString = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecPlaceFenceBlock(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                                    Optional z As Double = 0, Optional viewNum As Integer = 1) As String
    On Error GoTo E
    Dim pts(0 To 3) As Point3d
    pts(0).X = x1: pts(0).Y = y1: pts(0).Z = z
    pts(1).X = x2: pts(1).Y = y1: pts(1).Z = z
    pts(2).X = x2: pts(2).Y = y2: pts(2).Z = z
    pts(3).X = x1: pts(3).Y = y2: pts(3).Z = z
    Dim f As Fence
    Set f = ActiveDesignFile.Fence
    f.DefineFromModelPoints viewNum, pts
    ExecPlaceFenceBlock = "OK" & vbTab & "IsDefined=" & IIf(f.IsDefined, "Y", "N") & vbTab & _
                          "note=fence defined from block corners"
    Exit Function
E:
    ExecPlaceFenceBlock = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecFenceUndefine() As String
    On Error GoTo E
    ActiveDesignFile.Fence.Undefine
    ExecFenceUndefine = "OK" & vbTab & "note=fence undefined"
    Exit Function
E:
    ExecFenceUndefine = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecFenceCopyContents(deltaX As Double, deltaY As Double, _
                                      Optional deltaZ As Double = 0) As String
    On Error GoTo E
    Dim f As Fence
    Set f = ActiveDesignFile.Fence
    If Not f.IsDefined Then
        ExecFenceCopyContents = "ERROR" & vbTab & "note=no fence defined"
        Exit Function
    End If
    Dim ee As ElementEnumerator
    Set ee = f.GetContents()
    Dim delta As Point3d
    delta.X = deltaX: delta.Y = deltaY: delta.Z = deltaZ
    Dim newIds As String: newIds = ""
    Dim n As Integer: n = 0
    Do While ee.MoveNext
        Dim c As Element
        Set c = ee.Current.Clone()
        c.Move delta
        ActiveModelReference.AddElement c
        If newIds <> "" Then newIds = newIds & ","
        newIds = newIds & CStr(ElIDAsDouble(c.ID))
        n = n + 1
    Loop
    ExecFenceCopyContents = "OK" & vbTab & "copied=" & n & vbTab & "createdElementIds=" & newIds & vbTab & _
                            "note=copied fence contents"
    Exit Function
E:
    ExecFenceCopyContents = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecFenceMoveContents(deltaX As Double, deltaY As Double, _
                                      Optional deltaZ As Double = 0) As String
    On Error GoTo E
    Dim f As Fence
    Set f = ActiveDesignFile.Fence
    If Not f.IsDefined Then
        ExecFenceMoveContents = "ERROR" & vbTab & "note=no fence defined"
        Exit Function
    End If
    Dim ee As ElementEnumerator
    Set ee = f.GetContents()
    Dim delta As Point3d
    delta.X = deltaX: delta.Y = deltaY: delta.Z = deltaZ
    Dim n As Integer: n = 0
    Dim ids As String: ids = ""
    Do While ee.MoveNext
        Dim el As Element
        Set el = ee.Current
        el.Move delta
        el.Rewrite
        If ids <> "" Then ids = ids & ","
        ids = ids & CStr(ElIDAsDouble(el.ID))
        n = n + 1
    Loop
    ExecFenceMoveContents = "OK" & vbTab & "moved=" & n & vbTab & "elementIds=" & ids & vbTab & _
                            "priorDeltaX=" & (-deltaX) & vbTab & "priorDeltaY=" & (-deltaY) & vbTab & _
                            "note=moved fence contents"
    Exit Function
E:
    ExecFenceMoveContents = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecFenceDeleteContents() As String
    On Error GoTo E
    Dim f As Fence
    Set f = ActiveDesignFile.Fence
    If Not f.IsDefined Then
        ExecFenceDeleteContents = "ERROR" & vbTab & "note=no fence defined"
        Exit Function
    End If
    Dim ee As ElementEnumerator
    Set ee = f.GetContents()
    Dim n As Integer: n = 0
    Dim ids As String: ids = ""
    ' Collect first — enumerator may invalidate on remove
    Dim toKill As Object
    Set toKill = CreateObject("Scripting.Dictionary")
    Do While ee.MoveNext
        Dim idVal As Double: idVal = ElIDAsDouble(ee.Current.ID)
        toKill(CStr(idVal)) = True
        If ids <> "" Then ids = ids & ","
        ids = ids & CStr(idVal)
    Loop
    Dim k As Variant
    For Each k In toKill.Keys
        Dim el As Element
        Set el = FindElementByID(CDbl(k))
        If Not el Is Nothing Then
            ActiveModelReference.RemoveElement el
            n = n + 1
        End If
    Next k
    ExecFenceDeleteContents = "OK" & vbTab & "deleted=" & n & vbTab & "elementIds=" & ids & vbTab & _
                              "notUndoable=Y" & vbTab & "note=deleted fence contents"
    Exit Function
E:
    ExecFenceDeleteContents = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecSelectElementByID(elementId As Double, Optional clearFirst As Boolean = True) As String
    On Error GoTo E
    Dim el As Element
    Set el = FindElementByID(elementId)
    If el Is Nothing Then
        ExecSelectElementByID = "ERROR" & vbTab & "note=element not found: " & elementId
        Exit Function
    End If
    If clearFirst Then ActiveModelReference.UnselectAllElements
    ActiveModelReference.SelectElement el, True
    ExecSelectElementByID = "OK" & vbTab & "elementId=" & CStr(elementId) & vbTab & _
                            "AnyElementsSelected=" & IIf(ActiveModelReference.AnyElementsSelected, "Y", "N") & vbTab & _
                            "note=element selected"
    Exit Function
E:
    ExecSelectElementByID = "ERROR" & vbTab & "note=" & Err.Description
End Function

Public Function ExecClearSelection() As String
    On Error GoTo E
    ActiveModelReference.UnselectAllElements
    ExecClearSelection = "OK" & vbTab & "note=selection cleared"
    Exit Function
E:
    ExecClearSelection = "ERROR" & vbTab & "note=" & Err.Description
End Function

' verticesTSV: pipe-separated "x,y,z" — fills pts() 0-based, returns count
Private Function ParseVerticesTSV(verticesTSV As String, ByRef pts() As Point3d) As Integer
    Dim parts() As String
    parts = Split(verticesTSV, "|")
    Dim n As Integer: n = UBound(parts) + 1
    If n < 1 Then
        ParseVerticesTSV = 0
        Exit Function
    End If
    ReDim pts(0 To n - 1)
    Dim i As Integer
    For i = 0 To n - 1
        Dim coords() As String
        coords = Split(parts(i), ",")
        pts(i).X = CDbl(coords(0))
        pts(i).Y = CDbl(coords(1))
        If UBound(coords) >= 2 Then pts(i).Z = CDbl(coords(2)) Else pts(i).Z = 0
    Next i
    ParseVerticesTSV = n
End Function

' ============================================================
' FIND ELEMENT BY NUMERIC ID — scan-and-match, same pattern as
' ExecDeleteElementsByID / ExecSetSignAttributes. No direct
' GetElementByID call exists anywhere in this codebase to reuse.
' ============================================================
Private Function FindElementByID(elementId As Double) As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If ElIDAsDouble(el.ID) = elementId Then
            Set FindElementByID = el
            Exit Function
        End If
    Loop
    Set FindElementByID = Nothing
End Function
