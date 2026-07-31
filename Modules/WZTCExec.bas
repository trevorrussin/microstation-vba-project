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
        ExecPlaceSign = "ERROR" & vbTab & "note=sign not found in library: " & signNum
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
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendCommand "ATTACH LIBRARY " & sd.CellLibraryPath

    Dim pt1 As Point3d
    pt1.X = pt1X: pt1.Y = pt1Y: pt1.Z = pt1Z

    ' Legacy order (DrawSign.DrawSignAtPerpLine): text label -> sign face cell -> post
    Call DrawSign.PlaceSignFaceAndText(pt1, signNum, signSize, dir1X, dir1Y)
    Call DrawSign.DrawSignPost(pt1, dir1X, dir1Y)

    If bothSides Then
        Dim pt2 As Point3d
        pt2.X = pt2X: pt2.Y = pt2Y: pt2.Z = pt2Z
        Call DrawSign.PlaceSignFaceAndText(pt2, signNum, signSize, dir2X, dir2Y)
        Call DrawSign.DrawSignPost(pt2, dir2X, dir2Y)
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

    Dim lvl As String
    lvl = DrawElements.GetElementLevel(elementIdx)

    Dim pts() As String
    pts = Split(verticesTSV, "|")
    Dim n As Integer: n = UBound(pts) + 1
    If n < 2 Then
        ExecPlaceElementRun = "ERROR" & vbTab & "note=need at least 2 vertices"
        Exit Function
    End If

    CadInputQueue.SendCommand "ACTIVE LEVEL """ & lvl & """"
    CadInputQueue.SendCommand "ACTIVE COLOR 6"
    CadInputQueue.SendCommand "ACTIVE WEIGHT 2"
    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"

    Dim i As Integer, coords() As String, pt As Point3d
    For i = 0 To n - 1
        coords = Split(pts(i), ",")
        pt.X = CDbl(coords(0))
        pt.Y = CDbl(coords(1))
        If UBound(coords) >= 2 Then pt.Z = CDbl(coords(2)) Else pt.Z = 0
        CadInputQueue.SendDataPoint pt, 1
    Next i
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    ExecPlaceElementRun = "OK" & vbTab & "level=" & lvl & vbTab & "vertexCount=" & n & vbTab & _
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
' The hatch seed point is COMPUTED (FindInteriorPoint below), not
' clicked. DrawElements.bas's comment explains why the production
' code leaves this to a user click: "avoids centroid issues with
' non-convex shapes." FindInteriorPoint solves that properly with a
' horizontal-scanline / widest-even-odd-span method instead of a
' naive centroid, so it holds up on L-shaped and other non-convex
' work areas -- but it has only been checked against the test shapes
' in DebugExecTest.bas, not a large sample of real project geometry.
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

    Dim vx() As Double, vy() As Double, vz() As Double
    ReDim vx(0 To n - 1): ReDim vy(0 To n - 1): ReDim vz(0 To n - 1)
    Dim i As Integer, coords() As String
    For i = 0 To n - 1
        coords = Split(pts(i), ",")
        vx(i) = CDbl(coords(0))
        vy(i) = CDbl(coords(1))
        If UBound(coords) >= 2 Then vz(i) = CDbl(coords(2)) Else vz(i) = 0
    Next i

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZWS2_P"""
    CadInputQueue.SendCommand "ACTIVE COLOR 6"
    CadInputQueue.SendCommand "ACTIVE WEIGHT 2"
    CadInputQueue.SendCommand "PLACE SHAPE CONSTRAINED"

    Dim pt As Point3d
    For i = 0 To n - 1
        pt.X = vx(i): pt.Y = vy(i): pt.Z = vz(i)
        CadInputQueue.SendDataPoint pt, 1
    Next i
    ' Close explicitly by repeating the first vertex (matches
    ' Legacy Files/LegacyElements.bas's recorded sequence)
    pt.X = vx(0): pt.Y = vy(0): pt.Z = vz(0)
    CadInputQueue.SendDataPoint pt, 1
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    Dim seedX As Double, seedY As Double
    If Not FindInteriorPoint(vx, vy, n, seedX, seedY) Then
        ExecPlaceWorkspace = "OK" & vbTab & "vertexCount=" & n & vbTab & _
            "note=shape placed but hatch skipped -- could not compute an interior point " & _
            "(degenerate or self-intersecting boundary?)"
        Exit Function
    End If

    Dim hatchPt As Point3d
    hatchPt.X = seedX: hatchPt.Y = seedY: hatchPt.Z = vz(0)

    CadInputQueue.SendCommand "HATCH ICON"
    CadInputQueue.SendDataPoint hatchPt, 1
    CadInputQueue.SendDataPoint hatchPt, 1   ' same point twice -- matches LegacyElements.bas
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    ExecPlaceWorkspace = "OK" & vbTab & "vertexCount=" & n & vbTab & _
                         "hatchSeedX=" & Format(seedX, "0.00") & vbTab & "hatchSeedY=" & Format(seedY, "0.00") & vbTab & _
                         "note=placed work space shape and hatch"
    Exit Function

WsError:
    ExecPlaceWorkspace = "ERROR" & vbTab & "note=" & Err.Description
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
            el.Level = sfLevel
            el.Color = 240
            el.LineWeight = 3
            el.Rewrite
            applied = applied + 1
        End If
    Loop

    ExecSetSignAttributes = "OK" & vbTab & "applied=" & applied & vbTab & _
                            "requested=" & (UBound(idStrs) + 1) & vbTab & _
                            "note=set level=SF_P color=240 weight=3 " & _
                            "(fillColor/elementClass not replicated - no confirmed VBA property path)"
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
