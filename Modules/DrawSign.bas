Option Explicit

' Library path for current sign (set at start of DrawSignAtPerpLine; used when re-attaching after post)
Private currentSignFaceLibraryPath As String

' ============================================================
' SIGN PLACEMENT STATE AND ENTRY
' ------------------------------------------------------------
' Manages the sign drawing step that follows alignment placement.
' After the user clicks "Next: Draw Signs" in PlacePerp, StartSignPlacement
' shows PlaceSign and steps through each sign that had a perpendicular
' line placed. State is in wztcPlacedSign* (SharedState) and currentSignIdx.
' ============================================================
Public currentSignIdx As Integer   ' 0-based index into wztcPlacedSign* arrays

Public Sub StartSignPlacement()
    If wztcPlacedSignCount <= 0 Then
        MsgBox "No signs were recorded during alignment placement." & vbCrLf & _
               "Make sure sign numbers appear in the WZTC order and that you" & vbCrLf & _
               "clicked 'Place Line' (not 'Skip') for those items.", _
               vbExclamation, "Sign Placement"
        Exit Sub
    End If
    currentSignIdx = 0
    PlaceSign.Show vbModeless
End Sub

Public Function GetCurrentSignNum() As String
    If currentSignIdx >= 0 And currentSignIdx < wztcPlacedSignCount Then
        GetCurrentSignNum = wztcPlacedSignNums(currentSignIdx + 1)
    Else
        GetCurrentSignNum = ""
    End If
End Function

Public Function GetCurrentSignSide() As String
    If currentSignIdx >= 0 And currentSignIdx < wztcPlacedSignCount Then
        GetCurrentSignSide = wztcPlacedSignSide(currentSignIdx + 1)
    Else
        GetCurrentSignSide = ""
    End If
End Function

Public Function GetCurrentSignSize() As String
    If currentSignIdx >= 0 And currentSignIdx < wztcPlacedSignCount Then
        GetCurrentSignSize = wztcPlacedSignSize(currentSignIdx + 1)
    Else
        GetCurrentSignSize = ""
    End If
End Function

Public Function GetCurrentSignNumber() As Integer
    GetCurrentSignNumber = currentSignIdx + 1   ' 1-based for display
End Function

Public Function GetTotalSignCount() As Integer
    GetTotalSignCount = wztcPlacedSignCount
End Function

Public Function IsAllSignsDone() As Boolean
    IsAllSignsDone = (currentSignIdx >= wztcPlacedSignCount)
End Function

Public Sub AdvanceSign()
    currentSignIdx = currentSignIdx + 1
End Sub

Public Sub DrawCurrentSign()
    If currentSignIdx < 0 Or currentSignIdx >= wztcPlacedSignCount Then Exit Sub
    Dim n As Integer
    n = currentSignIdx + 1
    Call DrawSignAtPerpLine( _
        wztcPlacedSignNums(n), _
        wztcPlacedSignSize(n), _
        wztcPlacedSignSide(n), _
        wztcPlacedSignPtX(n), _
        wztcPlacedSignPtY(n), _
        wztcPlacedSignPtZ(n), _
        wztcPlacedSignPerpX(n), _
        wztcPlacedSignPerpY(n))
End Sub

' ============================================================
' PLACE WORKZONE SIGN
' ------------------------------------------------------------
' Called by DrawCurrentSign with the geometry of the perpendicular
' line where the sign goes. Uses SignLibrary for cell name and path.
'
' Parameters:
'   signNum   - sign number string (e.g. "W20-05", zero-padded to match cell library)
'   signSize  - sign size string from sign table (may contain " chars)
'   side      - "One Side" or "Both Sides"
'   midX/Y/Z  - alignment point = midpoint of the perpendicular line
'   perpX/Y   - unit perpendicular vector (perpendicular to alignment)
'
' Behaviour:
'   One Side   - collects 1 click, projects onto perp line,
'                draws: text label + sign face cell + post line + post cell.
'                No arc.
'   Both Sides - collects 2 clicks, projects each onto perp line,
'                draws both signs and a connecting arc between the posts.
'
' Sign placement order matches Legacy pattern:
'   1. Text label   2. Sign face cell   3. Post line + post cell
' ============================================================
Sub DrawSignAtPerpLine(signNum As String, signSize As String, side As String, _
                       midX As Double, midY As Double, midZ As Double, _
                       perpX As Double, perpY As Double)

    ' Ensure sign library is loaded before any lookup
    If SignLibrary.GetSignCount() = 0 Then Call SignLibrary.InitializeSignLibrary

    Const HALF_LEN As Double = 40   ' matches PERP_HALF_LEN in PerpPlacement

    ' Setup view
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    ' Capture the view's rotation BEFORE resetting it to identity below --
    ' the sign face cell is placed at THIS angle (see PlaceSignFaceAndText),
    ' not derived from the alignment/perp direction, so it always reads
    ' upright in whatever view the engineer currently has (confirmed via
    ' live testing 2026-08-01: a direction-derived rotation is
    ' mathematically guaranteed to flip text upside-down for some
    ' directions -- see Legacy Files/LegacySignPlace.bas, which never
    ' rotates the sign face cell at all, always ACTIVE ANGLE 0. Using the
    ' view's angle instead of a hardcoded 0 generalizes that same "always
    ' upright" intent to a rotated view instead of assuming the view is
    ' never rotated).
    Dim viewAngleDeg As Double
    viewAngleDeg = ViewRotationAngleDegrees(v)
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg   ' text label also reads upright in-view
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"

    ' Set element properties: Default level, color 0 (white), weight 0
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"

    ' Attach sign face library (use path from SignLibrary if sign is in library)
    currentSignFaceLibraryPath = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    If SignLibrary.SignExists(signNum) Then
        currentSignFaceLibraryPath = SignLibrary.GetSignData(signNum).CellLibraryPath
    End If
    CadInputQueue.SendCommand "ATTACH LIBRARY " & currentSignFaceLibraryPath

    Dim bothSides As Boolean
    bothSides = (Trim(side) = "Both Sides")

    Dim oMsg As CadInputMessage
    Dim pt1 As Point3d

    ' ---- Collect first click ----
    CadInputQueue.SendKeyin "ECHO Click post location on perpendicular line for sign " & signNum
    CadInputQueue.SendCommand "NULL"
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendKeyin "ECHO Sign placement cancelled"
            CommandState.StartDefaultCommand
            Exit Sub
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    ' Project clicked point onto the perpendicular line segment
    pt1 = ProjectOntoPerp(oMsg.Point, midX, midY, midZ, perpX, perpY, HALF_LEN)

    If Not bothSides Then
        ' =====================================================
        ' ONE SIDE
        ' =====================================================
        CadInputQueue.SendReset
        CommandState.StartDefaultCommand

        ' Outward direction = away from alignment midpoint along perp vector
        Dim t1 As Double
        t1 = (pt1.X - midX) * perpX + (pt1.Y - midY) * perpY
        Dim d1X As Double, d1Y As Double
        If t1 >= 0 Then
            d1X = perpX:  d1Y = perpY
        Else
            d1X = -perpX: d1Y = -perpY
        End If

        ' attachmentPt = click on perp tick; dir = outward along perp
        Call PlaceSignAssembly(pt1, signNum, signSize, d1X, d1Y, viewAngleDeg)

    Else
        ' =====================================================
        ' BOTH SIDES
        ' =====================================================
        ' Show dynamic line feedback while waiting for second click
        Dim p As Point3d
        p.X = pt1.X: p.Y = pt1.Y: p.Z = pt1.Z
        CadInputQueue.SendCommand "PLACE LINE"
        CadInputQueue.SendDataPoint p, 1

        CadInputQueue.SendKeyin "ECHO Click SECOND post location on perpendicular line"
        Set oMsg = CadInputQueue.GetInput
        Do While oMsg.InputType <> msdCadInputTypeDataPoint
            If oMsg.InputType = msdCadInputTypeReset Then
                CadInputQueue.SendKeyin "ECHO Sign placement cancelled"
                CadInputQueue.SendReset
                CommandState.StartDefaultCommand
                Exit Sub
            End If
            Set oMsg = CadInputQueue.GetInput
        Loop

        Dim pt2 As Point3d
        pt2 = ProjectOntoPerp(oMsg.Point, midX, midY, midZ, perpX, perpY, HALF_LEN)

        CadInputQueue.SendReset
        CommandState.StartDefaultCommand

        ' Outward direction for each point = away from alignment midpoint
        Dim tA As Double, tB As Double
        tA = (pt1.X - midX) * perpX + (pt1.Y - midY) * perpY
        tB = (pt2.X - midX) * perpX + (pt2.Y - midY) * perpY

        Dim dAX As Double, dAY As Double
        Dim dBX As Double, dBY As Double
        If tA >= 0 Then
            dAX = perpX:  dAY = perpY
        Else
            dAX = -perpX: dAY = -perpY
        End If
        If tB >= 0 Then
            dBX = perpX:  dBY = perpY
        Else
            dBX = -perpX: dBY = -perpY
        End If

        ' attachmentPts = clicks on perp tick; dirs = outward along perp
        Call PlaceSignAssembly(pt1, signNum, signSize, dAX, dAY, viewAngleDeg)
        Call PlaceSignAssembly(pt2, signNum, signSize, dBX, dBY, viewAngleDeg)

        Call DrawConnectingArc(pt1, pt2)
    End If

    CommandState.StartDefaultCommand
    CadInputQueue.SendKeyin "ECHO Sign " & signNum & " placed."

End Sub

' ============================================================
' PROJECT A CLICKED POINT ONTO THE PERPENDICULAR LINE SEGMENT
' Returns the nearest point on the segment to the clicked point,
' clamped to ±halfLen from the alignment midpoint.
' ============================================================
Private Function ProjectOntoPerp(clickPt As Point3d, _
                                   midX As Double, midY As Double, midZ As Double, _
                                   perpX As Double, perpY As Double, _
                                   halfLen As Double) As Point3d
    Dim t As Double
    t = (clickPt.X - midX) * perpX + (clickPt.Y - midY) * perpY
    If t < -halfLen Then t = -halfLen
    If t > halfLen Then t = halfLen

    Dim result As Point3d
    result.X = midX + t * perpX
    result.Y = midY + t * perpY
    result.Z = midZ
    ProjectOntoPerp = result
End Function

' ============================================================
' DRAW SIGN POST CELL + STEM LINE  (thin wrapper -- real work is
' PlaceSignAssembly, which also places the face/label with edge-
' connected stem geometry).
' ============================================================
Sub DrawSignPost(postPt As Point3d, dirX As Double, dirY As Double)
    ' No-op when PlaceSignAssembly already ran (preferred path).
    ' Kept so any stray callers that only want a post still compile;
    ' Prefer PlaceSignAssembly for a full one-side assembly.
End Sub

' ============================================================
' PLACE FULL SIGN ASSEMBLY AT A PERP-TICK ATTACHMENT POINT
' ------------------------------------------------------------
' attachmentPt = point ON the perpendicular tick where the post
'   should meet the tick (typically the outward tip of the 80ft
'   perp line -- NOT the alignment center, NOT the face center).
' dirX/Y = unit outward direction along the perp, away from the
'   alignment (the direction the stem and face go).
'
' Geometry (confirmed against engineer reference 2026-08-03):
'   - TWZSGN_P is asymmetric (vertical shaft + east crossbar at 0 deg).
'     Rotate the POST with travel tangent (arm toward downstream, stem
'     toward upstream, T along the curve). Face + MUTCD text stay at
'     viewAngleDeg (view-horizontal). Live miss 2026-08-13: post used
'     viewAngle so G20-2 on a 90 deg bend sat 90 deg off the roadside.
'   - Align the SHAFT CENTERLINE to the tip laterally, and the POST'S
'     INWARD EDGE to the tip along dir -- not the cell bbox center.
'   - Face origin further out so FACE INWARD EDGE is STEM_GAP beyond the
'     post outward edge; stem stays on the tip's lateral line.
'   - Stem LINE connects those two edges only -- never to face center.
' STEM_GAP=50 matches the live reference; old "20 ft to face origin"
' put the line through a Scale=960 face.
' ============================================================
Public Sub PlaceSignAssembly(attachmentPt As Point3d, signNum As String, signSize As String, _
                             dirX As Double, dirY As Double, viewAngleDeg As Double, _
                             Optional postAngleDeg As Double = -9999#)
    Const STEM_GAP As Double = 50#
    Dim postAng As Double
    postAng = postAngleDeg
    If postAng <= -9000# Then postAng = viewAngleDeg

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & postAng

    If Len(currentSignFaceLibraryPath) = 0 Then
        currentSignFaceLibraryPath = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    End If

    ' --- Post cell: tip = shaft centerline + inward edge (not bbox center) ---
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    SetCExpressionValue "tcb->activeCellUtf16", "TWZSGN_P", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendDataPoint attachmentPt, 1
    CadInputQueue.SendReset

    Dim postEl As Element
    Set postEl = FindNewestElement()
    If postEl Is Nothing Then Exit Sub

    Dim halfPost As Double
    halfPost = HalfExtentAlongDir(postEl, dirX, dirY)
    ' Lateral unit (perp to outward dir). Shaft sits west of cell origin
    ' for TWZSGN_P at 0 deg -- offset measured from live cell, not guessed.
    Dim latX As Double, latY As Double
    latX = -dirY
    latY = dirX
    Dim shaftLatFromOrigin As Double
    shaftLatFromOrigin = ShaftLateralOffsetFromOrigin(postEl, latX, latY)

    Dim postOrigin As Point3d
    postOrigin.X = attachmentPt.X + dirX * halfPost - latX * shaftLatFromOrigin
    postOrigin.Y = attachmentPt.Y + dirY * halfPost - latY * shaftLatFromOrigin
    postOrigin.Z = attachmentPt.Z
    Call MoveCellOriginTo(postEl, postOrigin)
    ' Nudge along dir so the measured inward edge lands exactly on the tip
    ' (half-extent estimate can be ~0.004ft short of the true range edge).
    Call SnapInwardEdgeToTip(postEl, attachmentPt, dirX, dirY)
    halfPost = HalfExtentAlongDir(postEl, dirX, dirY)

    ' Outer edge of post, forced onto the stem ray (same as L: line on perp).
    Dim postOuter As Point3d
    postOuter = ExtremePointAlongDir(postEl, dirX, dirY, True)
    Dim tPost As Double
    tPost = (postOuter.X - attachmentPt.X) * dirX + (postOuter.Y - attachmentPt.Y) * dirY
    If tPost < 0.5 Then tPost = 2# * halfPost
    postOuter.X = attachmentPt.X + dirX * tPost
    postOuter.Y = attachmentPt.Y + dirY * tPost
    postOuter.Z = attachmentPt.Z

    ' --- Face cell (centered on tip laterally) ---
    Dim cellName As String
    cellName = ""
    If SignLibrary.SignExists(signNum) Then
        cellName = SignLibrary.GetSignData(signNum).CellName
    End If
    If Len(cellName) = 0 Then
        CadInputQueue.SendKeyin "ECHO WARNING: Sign " & signNum & " not found in library - face cell skipped"
        Exit Sub
    End If

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg
    CadInputQueue.SendCommand "ATTACH LIBRARY " & currentSignFaceLibraryPath
    SetCExpressionValue "tcb->activeCellUtf16", cellName, ""

    Dim guessPt As Point3d
    guessPt.X = postOuter.X + dirX * (STEM_GAP + 30#)
    guessPt.Y = postOuter.Y + dirY * (STEM_GAP + 30#)
    guessPt.Z = postOuter.Z
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendDataPoint guessPt, 1
    CadInputQueue.SendReset

    Dim faceEl As Element
    Set faceEl = FindNewestElement()
    If faceEl Is Nothing Then Exit Sub

    Dim halfFace As Double
    halfFace = HalfExtentAlongDir(faceEl, dirX, dirY)
    Dim faceOrigin As Point3d
    faceOrigin.X = postOuter.X + dirX * (STEM_GAP + halfFace)
    faceOrigin.Y = postOuter.Y + dirY * (STEM_GAP + halfFace)
    faceOrigin.Z = postOuter.Z
    Call MoveCellOriginTo(faceEl, faceOrigin)

    Dim faceInner As Point3d
    ' Snap face so its inward edge is exactly STEM_GAP past postOuter
    Dim faceTarget As Point3d
    faceTarget.X = postOuter.X + dirX * STEM_GAP
    faceTarget.Y = postOuter.Y + dirY * STEM_GAP
    faceTarget.Z = postOuter.Z
    Call SnapInwardEdgeToTip(faceEl, faceTarget, dirX, dirY)
    halfFace = HalfExtentAlongDir(faceEl, dirX, dirY)
    If faceEl.IsCellElement Then
        faceOrigin = faceEl.AsCellElement.Origin
    ElseIf faceEl.IsSharedCellElement Then
        faceOrigin = faceEl.AsSharedCellElement.Origin
    End If
    ' Stem ends on the perp ray, 1.5 ft into the orange fill so the white
    ' line meets the face (G20 rect + diamonds). Prefer color-6 vertices;
    ' construction/hole verts sit short of the fill and leave a gap.
    Dim inPt As Point3d
    Dim tFace As Double
    inPt = ExtremePointAlongDirPreferColor(faceEl, dirX, dirY, False, 6)
    tFace = (inPt.X - attachmentPt.X) * dirX + (inPt.Y - attachmentPt.Y) * dirY
    If tFace < tPost + 1# Then tFace = tPost + STEM_GAP
    tFace = tFace + 1.5
    faceInner.X = attachmentPt.X + dirX * tFace
    faceInner.Y = attachmentPt.Y + dirY * tFace
    faceInner.Z = faceTarget.Z

    ' W04-02* cells ship with yellow SF_P copies of the merge symbol on
    ' top of the black SFB_P legend; hide the small yellow duplicates so
    ' the black symbol reads (live 2026-08-03 south 619-311 QA).
    Call HideDuplicateYellowLegend(faceEl, cellName)

    ' G20-* faces: library place can leave the SF_P "grouped hole" (inner
    ' border cell) at color 6/orange instead of 240/black — engineer
    ' reference next to the Y=297000 build had SF_P hole=240 with orange
    ' outer (SF_P complex) + black SFB_P legend (2026-08-10).
    Call FixG20FaceBlackHole(faceEl, signNum)

    ' --- Stem: post outward edge -> face inward edge only (on tip line) ---
    ' Element API (CreateLineElement2), NOT PLACE LINE CONSTRAINED: AccuDraw
    ' distance lock left over from a long define_alignment_segment was making
    ' CadInputQueue stems 3000ft instead of STEM_GAP (live 2026-08-03).
    Dim stemEl As LineElement
    Set stemEl = CreateLineElement2(Nothing, postOuter, faceInner)
    stemEl.Color = 0
    stemEl.LineWeight = 0
    On Error Resume Next
    stemEl.Level = ActiveDesignFile.Levels("Default")
    On Error GoTo 0
    ActiveModelReference.AddElement stemEl
    stemEl.Rewrite

    ' --- Text label beyond the face ---
    Dim textPt As Point3d
    textPt.X = faceOrigin.X + dirX * (halfFace + 20#)
    textPt.Y = faceOrigin.Y + dirY * (halfFace + 20#)
    textPt.Z = faceOrigin.Z

    ' Same two-line label as every other MUTCD sign: code above, size
    ' below (engineer hand-fixed G20-02 on the Y=299000 corridor to
    ' "G20-02" / "36\" x 18\"" — size-only was wrong). Use TEXTEDITOR
    ' with signNum FIRST: a prior G20-only size-first TEXTEDITOR path
    ' silently produced no text (2026-08-10); Element API size-only also
    ' omitted the code the engineer requires above the size callout.
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & signNum & """"
    If Len(signSize) > 0 Then
        CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
        Call InsertTextWithInchMarks(signSize)
    End If
    CadInputQueue.SendDataPoint textPt, 1
    CadInputQueue.SendReset

    CadInputQueue.SendCommand "ATTACH LIBRARY " & currentSignFaceLibraryPath
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg
End Sub

' Backward-compatible name used by ExecPlaceSign / DrawSignAtPerpLine:
' attachmentPt semantics (perp tip), not "post origin / face at +20".
Sub PlaceSignFaceAndText(postPt As Point3d, signNum As String, signSize As String, _
                          dirX As Double, dirY As Double, viewAngleDeg As Double)
    Call PlaceSignAssembly(postPt, signNum, signSize, dirX, dirY, viewAngleDeg)
End Sub

' ============================================================
' Cell-on-a-post assembly (Arrow Panel and similar plan symbols) —
' same stem-then-face geometry as PlaceSignAssembly's post/stem/face
' chain, minus the SignLibrary lookup and code/size text a MUTCD sign
' needs. Engineer asked for the arrow panel to sit the same way the
' roadside signs do (2026-08-10) instead of floating at a bare lateral
' offset with no stem.
' ============================================================
Public Function PlaceCellOnPost(basePt As Point3d, cellName As String, _
                                 dirX As Double, dirY As Double, _
                                 viewAngleDeg As Double) As String
    On Error GoTo PostErr

    ' Stem line from the tick/base out STEM_GAP, then snap the plan cell
    ' so its inward edge sits on the stem tip — same "post then face"
    ' look as roadside signs, but WITHOUT placing TWZSGN_P. Arrow panel
    ' sits on the channelizing tip; a TWZSGN_P cell there stacks on the
    ' cone and reads as a doubled orange marker (live 2026-08-10
    ' Y=298000 rebuild). Engineer demo on Y=299000 used stem line +
    ' panel only.
    Const STEM_GAP As Double = 50#
    Const WZTC_CELL_LIB As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"

    Dim postOuter As Point3d
    postOuter.X = basePt.X + dirX * STEM_GAP
    postOuter.Y = basePt.Y + dirY * STEM_GAP
    postOuter.Z = basePt.Z

    Dim stemEl As LineElement
    Set stemEl = CreateLineElement2(Nothing, basePt, postOuter)
    stemEl.Color = 0
    stemEl.LineWeight = 0
    On Error Resume Next
    stemEl.Level = ActiveDesignFile.Levels("Default")
    On Error GoTo PostErr
    ActiveModelReference.AddElement stemEl
    stemEl.Rewrite

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & viewAngleDeg
    CadInputQueue.SendCommand "ATTACH LIBRARY " & WZTC_CELL_LIB
    SetCExpressionValue "tcb->activeCellUtf16", cellName, ""

    Dim guessPt As Point3d
    guessPt.X = postOuter.X + dirX * 30#
    guessPt.Y = postOuter.Y + dirY * 30#
    guessPt.Z = postOuter.Z
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendDataPoint guessPt, 1
    CadInputQueue.SendReset

    Dim cellEl As Element
    Set cellEl = FindNewestElement()
    If cellEl Is Nothing Then
        PlaceCellOnPost = "ERROR" & vbTab & "note=cell placement failed"
        Exit Function
    End If

    Call SnapInwardEdgeToTip(cellEl, postOuter, dirX, dirY)

    Dim finalOrigin As Point3d
    If cellEl.IsCellElement Then
        finalOrigin = cellEl.AsCellElement.Origin
    ElseIf cellEl.IsSharedCellElement Then
        finalOrigin = cellEl.AsSharedCellElement.Origin
    End If

    PlaceCellOnPost = "OK" & vbTab & "elementId=" & CStr(ElIDAsDouble(cellEl.ID)) & vbTab & _
                       "stemElementId=" & CStr(ElIDAsDouble(stemEl.ID)) & vbTab & _
                       "x=" & finalOrigin.X & vbTab & "y=" & finalOrigin.Y & vbTab & _
                       "note=cell-on-post assembly (stem " & STEM_GAP & "ft, no TWZSGN_P)"
    Exit Function

PostErr:
    PlaceCellOnPost = "ERROR" & vbTab & "note=" & Err.Description
End Function

' Half-size of element along unit dir using real vertices (not AABB).
' AABB over-estimates a view-horizontal diamond on a diagonal stem, so
' the white post line stopped short of the face on C/S curves.
Private Function HalfExtentAlongDir(el As Element, dirX As Double, dirY As Double) As Double
    Dim minP As Double, maxP As Double
    Dim anyPts As Boolean
    anyPts = False
    minP = 0#: maxP = 0#
    Call AccumulateDirProj(el, dirX, dirY, minP, maxP, anyPts)
    If anyPts Then
        HalfExtentAlongDir = 0.5 * (maxP - minP)
    Else
        Dim rng As Range3d
        rng = el.Range
        HalfExtentAlongDir = 0.5 * (Abs(dirX) * (rng.High.X - rng.Low.X) + _
                                    Abs(dirY) * (rng.High.Y - rng.Low.Y))
    End If
    If HalfExtentAlongDir < 0.01 Then HalfExtentAlongDir = 0.01
End Function

Private Function ExtremePointAlongDirPreferColor(el As Element, dirX As Double, dirY As Double, _
                                                 wantMax As Boolean, preferColor As Long) As Point3d
    Dim best As Point3d
    Dim bestP As Double
    Dim found As Boolean
    found = False
    bestP = 0#
    Call FindExtremeAlongDirColor(el, dirX, dirY, wantMax, preferColor, best, bestP, found)
    If found Then
        ExtremePointAlongDirPreferColor = best
        Exit Function
    End If
    ExtremePointAlongDirPreferColor = ExtremePointAlongDir(el, dirX, dirY, wantMax)
End Function

Private Sub FindExtremeAlongDirColor(el As Element, dirX As Double, dirY As Double, _
                                     wantMax As Boolean, preferColor As Long, _
                                     ByRef best As Point3d, ByRef bestP As Double, _
                                     ByRef found As Boolean)
    On Error Resume Next
    Dim ce As ElementEnumerator
    Set ce = Nothing
    If el.IsCellElement Then
        Set ce = el.AsCellElement.GetSubElements
    ElseIf el.IsSharedCellElement Then
        Set ce = el.AsSharedCellElement.GetSubElements
    End If
    Err.Clear
    If Not ce Is Nothing Then
        Dim subEl As Element
        Do While ce.MoveNext
            Set subEl = ce.Current
            Call FindExtremeAlongDirColor(subEl, dirX, dirY, wantMax, preferColor, best, bestP, found)
        Loop
        Exit Sub
    End If
    If el.Color <> preferColor Then Exit Sub
    On Error GoTo NoVerts
    Dim vl As VertexList
    Dim pts() As Point3d
    Dim i As Long
    Dim p As Double
    Set vl = el.AsVertexList
    pts = vl.GetVertices()
    For i = LBound(pts) To UBound(pts)
        p = pts(i).X * dirX + pts(i).Y * dirY
        If (Not found) Or (wantMax And p > bestP) Or ((Not wantMax) And p < bestP) Then
            best = pts(i)
            bestP = p
            found = True
        End If
    Next i
NoVerts:
End Sub

Private Function ExtremePointAlongDir(el As Element, dirX As Double, dirY As Double, _
                                      wantMax As Boolean) As Point3d
    Dim best As Point3d
    Dim bestP As Double
    Dim found As Boolean
    found = False
    bestP = 0#
    Call FindExtremeAlongDir(el, dirX, dirY, wantMax, best, bestP, found)
    If found Then
        ExtremePointAlongDir = best
        Exit Function
    End If
    Dim rng As Range3d
    rng = el.Range
    Dim midX As Double, midY As Double, half As Double
    midX = 0.5 * (rng.Low.X + rng.High.X)
    midY = 0.5 * (rng.Low.Y + rng.High.Y)
    half = HalfExtentAlongDir(el, dirX, dirY)
    If wantMax Then
        best.X = midX + dirX * half
        best.Y = midY + dirY * half
    Else
        best.X = midX - dirX * half
        best.Y = midY - dirY * half
    End If
    best.Z = 0#
    ExtremePointAlongDir = best
End Function

Private Sub AccumulateDirProj(el As Element, dirX As Double, dirY As Double, _
                              ByRef minP As Double, ByRef maxP As Double, _
                              ByRef anyPts As Boolean)
    On Error GoTo TrySubs
    Dim vl As VertexList
    Dim pts() As Point3d
    Dim i As Long
    Dim p As Double
    Set vl = el.AsVertexList
    pts = vl.GetVertices()
    For i = LBound(pts) To UBound(pts)
        p = pts(i).X * dirX + pts(i).Y * dirY
        If Not anyPts Then
            minP = p: maxP = p: anyPts = True
        Else
            If p < minP Then minP = p
            If p > maxP Then maxP = p
        End If
    Next i
    Exit Sub
TrySubs:
    On Error Resume Next
    Dim ce As ElementEnumerator
    If el.IsCellElement Then
        Set ce = el.AsCellElement.GetSubElements
    ElseIf el.IsSharedCellElement Then
        Set ce = el.AsSharedCellElement.GetSubElements
    Else
        Exit Sub
    End If
    If ce Is Nothing Then Exit Sub
    Dim subEl As Element
    Do While ce.MoveNext
        Set subEl = ce.Current
        Call AccumulateDirProj(subEl, dirX, dirY, minP, maxP, anyPts)
    Loop
End Sub

Private Sub FindExtremeAlongDir(el As Element, dirX As Double, dirY As Double, _
                                wantMax As Boolean, ByRef best As Point3d, _
                                ByRef bestP As Double, ByRef found As Boolean)
    On Error GoTo TrySubs2
    Dim vl As VertexList
    Dim pts() As Point3d
    Dim i As Long
    Dim p As Double
    Set vl = el.AsVertexList
    pts = vl.GetVertices()
    For i = LBound(pts) To UBound(pts)
        p = pts(i).X * dirX + pts(i).Y * dirY
        If (Not found) Or (wantMax And p > bestP) Or ((Not wantMax) And p < bestP) Then
            best = pts(i)
            bestP = p
            found = True
        End If
    Next i
    Exit Sub
TrySubs2:
    On Error Resume Next
    Dim ce As ElementEnumerator
    If el.IsCellElement Then
        Set ce = el.AsCellElement.GetSubElements
    ElseIf el.IsSharedCellElement Then
        Set ce = el.AsSharedCellElement.GetSubElements
    Else
        Exit Sub
    End If
    If ce Is Nothing Then Exit Sub
    Dim subEl As Element
    Do While ce.MoveNext
        Set subEl = ce.Current
        Call FindExtremeAlongDir(subEl, dirX, dirY, wantMax, best, bestP, found)
    Loop
End Sub

' First hit of ray (ox,oy)+t*(dir) with cell outline. Keeps the stem on
' the perp line and on the diamond edge (not an inner hole AABB gap).
Private Function RayHitOutline(el As Element, ox As Double, oy As Double, _
                               dirX As Double, dirY As Double) As Point3d
    Dim bestT As Double
    Dim found As Boolean
    bestT = 1E+99
    found = False
    Call AccelRayHits(el, ox, oy, dirX, dirY, bestT, found)
    Dim p As Point3d
    If found And bestT > 0.01 And bestT < 1E+98 Then
        p.X = ox + dirX * bestT
        p.Y = oy + dirY * bestT
        p.Z = 0#
    Else
        p = ExtremePointAlongDir(el, dirX, dirY, False)
        Dim t As Double
        t = (p.X - ox) * dirX + (p.Y - oy) * dirY
        If t < 0.01 Then t = 50#
        p.X = ox + dirX * t
        p.Y = oy + dirY * t
        p.Z = 0#
    End If
    RayHitOutline = p
End Function

Private Sub AccelRayHits(el As Element, ox As Double, oy As Double, _
                         dirX As Double, dirY As Double, _
                         ByRef bestT As Double, ByRef found As Boolean)
    On Error GoTo TrySubs3
    Dim vl As VertexList
    Dim pts() As Point3d
    Dim i As Long
    Dim n As Long
    Dim t As Double
    Set vl = el.AsVertexList
    pts = vl.GetVertices()
    n = UBound(pts) - LBound(pts) + 1
    If n < 2 Then Exit Sub
    For i = LBound(pts) To UBound(pts) - 1
        t = SegRayT(ox, oy, dirX, dirY, pts(i).X, pts(i).Y, pts(i + 1).X, pts(i + 1).Y)
        If t > 0.05 And t < bestT Then
            bestT = t
            found = True
        End If
    Next i
    ' Close only a small loop (one diamond/rect). Do NOT chord last→first
    ' across a whole cell vertex dump — that chord sits short of the orange
    ' face and is the C/S stem gap (L was axis-aligned so the chord missed).
    Dim closeLen As Double
    closeLen = Sqr((pts(UBound(pts)).X - pts(LBound(pts)).X) ^ 2 + _
                   (pts(UBound(pts)).Y - pts(LBound(pts)).Y) ^ 2)
    If closeLen > 0.5 And closeLen < 80# Then
        t = SegRayT(ox, oy, dirX, dirY, pts(UBound(pts)).X, pts(UBound(pts)).Y, _
                    pts(LBound(pts)).X, pts(LBound(pts)).Y)
        If t > 0.05 And t < bestT Then
            bestT = t
            found = True
        End If
    End If
    Exit Sub
TrySubs3:
    On Error Resume Next
    Dim ce As ElementEnumerator
    If el.IsCellElement Then
        Set ce = el.AsCellElement.GetSubElements
    ElseIf el.IsSharedCellElement Then
        Set ce = el.AsSharedCellElement.GetSubElements
    Else
        Exit Sub
    End If
    If ce Is Nothing Then Exit Sub
    Dim subEl As Element
    Do While ce.MoveNext
        Set subEl = ce.Current
        Call AccelRayHits(subEl, ox, oy, dirX, dirY, bestT, found)
    Loop
End Sub

Private Function SegRayT(ox As Double, oy As Double, dx As Double, dy As Double, _
                         ax As Double, ay As Double, bx As Double, by As Double) As Double
    Dim rx As Double, ry As Double, sx As Double, sy As Double
    Dim den As Double, t As Double, u As Double
    rx = dx: ry = dy
    sx = bx - ax: sy = by - ay
    den = rx * sy - ry * sx
    SegRayT = -1#
    If Abs(den) < 0.0000001 Then Exit Function
    t = ((ax - ox) * sy - (ay - oy) * sx) / den
    u = ((ax - ox) * ry - (ay - oy) * rx) / den
    If t > 0# And u >= -0.02 And u <= 1.02 Then
        SegRayT = t
    End If
End Function

' Lateral offset of the post SHAFT centerline from the cell origin.
' TWZSGN_P = tall shaft + shorter crossbar; pick the subelement with the
' largest extent along the outward dir (the shaft), then return
' (shaftMid - origin) · lateral. Used so the shaft sits on the tip line
' instead of centering the whole T bbox on the tip.
Private Function ShaftLateralOffsetFromOrigin(el As Element, latX As Double, latY As Double) As Double
    Dim origin As Point3d
    If el.IsCellElement Then
        origin = el.AsCellElement.Origin
    ElseIf el.IsSharedCellElement Then
        origin = el.AsSharedCellElement.Origin
    Else
        ShaftLateralOffsetFromOrigin = 0#
        Exit Function
    End If

    Dim bestExtent As Double
    bestExtent = -1#
    Dim shaftMidX As Double, shaftMidY As Double
    shaftMidX = origin.X
    shaftMidY = origin.Y

    On Error GoTo FallbackBBox
    Dim ce As ElementEnumerator
    If el.IsCellElement Then
        Set ce = el.AsCellElement.GetSubElements
    Else
        Set ce = el.AsSharedCellElement.GetSubElements
    End If
    Dim subEl As Element
    Dim sr As Range3d
    Dim sideMax As Double
    Do While ce.MoveNext
        Set subEl = ce.Current
        sr = subEl.Range
        ' Prefer the taller shaft (larger max side of its bbox)
        If (sr.High.Y - sr.Low.Y) > (sr.High.X - sr.Low.X) Then
            sideMax = sr.High.Y - sr.Low.Y
        Else
            sideMax = sr.High.X - sr.Low.X
        End If
        If sideMax > bestExtent Then
            bestExtent = sideMax
            shaftMidX = 0.5 * (sr.Low.X + sr.High.X)
            shaftMidY = 0.5 * (sr.Low.Y + sr.High.Y)
        End If
    Loop

    ShaftLateralOffsetFromOrigin = (shaftMidX - origin.X) * latX + (shaftMidY - origin.Y) * latY
    Exit Function

FallbackBBox:
    ' No subelements -- fall back to bbox center (= 0 offset from origin for
    ' a centered cell; for TWZSGN_P this path should not run).
    Dim rng As Range3d
    rng = el.Range
    shaftMidX = 0.5 * (rng.Low.X + rng.High.X)
    shaftMidY = 0.5 * (rng.Low.Y + rng.High.Y)
    ShaftLateralOffsetFromOrigin = (shaftMidX - origin.X) * latX + (shaftMidY - origin.Y) * latY
End Function

' W04-02* cells ship with yellow SF_P copies of the merge symbol on
' top of the black SFB_P legend. Delete those small yellow duplicates
' and raise DisplayPriority on black SFB_P so the symbol reads above
' the yellow diamond fill (IsHidden on cell components does not stick;
' live 2026-08-03 south 619-311 QA).
' G20-* sign face: force the SF_P inner "grouped hole" cell to color 240
' (black) so the stack reads black lettering / orange field / black inner
' border / orange outer — matching engineer hand-placed reference.
Private Sub FixG20FaceBlackHole(faceEl As Element, signNum As String)
    On Error GoTo FixDone
    If Left$(UCase$(signNum), 4) <> "G20-" Then Exit Sub
    If Not faceEl.IsCellElement Then Exit Sub

    Dim cell As CellElement
    Set cell = faceEl.AsCellElement
    cell.ResetElementEnumeration

    Dim subEl As Element
    Dim lvlName As String
    Dim col As Long
    Dim guard As Integer: guard = 0
    Do While cell.MoveToNextElement(False) And guard < 40
        guard = guard + 1
        Set subEl = cell.CopyCurrentElement
        On Error Resume Next
        lvlName = ""
        lvlName = subEl.Level.Name
        col = -1
        col = CLng(subEl.Color)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextHole
        End If
        On Error GoTo FixDone

        ' Nested cell on SF_P at orange (6) is the missing black hole;
        ' leave SF_P type-14 outer complex orange, leave SFB_P alone.
        If UCase$(lvlName) = "SF_P" And col = 6 And subEl.IsCellElement Then
            subEl.Color = 240
            cell.ReplaceCurrentElement subEl
        End If
NextHole:
    Loop
    faceEl.Rewrite
FixDone:
End Sub

Private Sub HideDuplicateYellowLegend(faceEl As Element, cellName As String)
    On Error GoTo HideDone
    If Len(cellName) < 6 Then Exit Sub
    If UCase$(Left$(cellName, 6)) <> "W04-02" Then Exit Sub
    If Not faceEl.IsCellElement Then Exit Sub

    Dim fr As Range3d
    fr = faceEl.Range
    Dim faceMax As Double
    If (fr.High.X - fr.Low.X) > (fr.High.Y - fr.Low.Y) Then
        faceMax = fr.High.X - fr.Low.X
    Else
        faceMax = fr.High.Y - fr.Low.Y
    End If
    If faceMax < 0.01 Then Exit Sub

    Dim cell As CellElement
    Set cell = faceEl.AsCellElement
    cell.ResetElementEnumeration

    Dim subEl As Element
    Dim sr As Range3d
    Dim sideMax As Double
    Dim lvlName As String
    Dim col As Long
    Dim guard As Integer: guard = 0
    Do While cell.MoveToNextElement(False) And guard < 40
        guard = guard + 1
        Set subEl = cell.CopyCurrentElement
        On Error Resume Next
        lvlName = ""
        lvlName = subEl.Level.Name
        col = -1
        col = CLng(subEl.Color)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextSub
        End If
        On Error GoTo HideDone

        sr = subEl.Range
        If (sr.High.Y - sr.Low.Y) > (sr.High.X - sr.Low.X) Then
            sideMax = sr.High.Y - sr.Low.Y
        Else
            sideMax = sr.High.X - sr.Low.X
        End If

        If UCase$(lvlName) = "SF_P" And col = 4 And sideMax < 0.7 * faceMax Then
            cell.DeleteCurrentElement
            GoTo NextSub
        End If

        If UCase$(lvlName) = "SFB_P" And col = 240 Then
            On Error Resume Next
            subEl.DisplayPriority = 2000
            cell.ReplaceCurrentElement subEl
            Err.Clear
            On Error GoTo HideDone
        ElseIf UCase$(lvlName) = "SF_P" And col = 4 Then
            On Error Resume Next
            subEl.DisplayPriority = -2000
            cell.ReplaceCurrentElement subEl
            Err.Clear
            On Error GoTo HideDone
        End If
NextSub:
    Loop
    faceEl.Rewrite
HideDone:
End Sub

Private Sub MoveCellOriginTo(el As Element, newOrigin As Point3d)
    Dim oldOrigin As Point3d
    If el.IsCellElement Then
        oldOrigin = el.AsCellElement.Origin
    ElseIf el.IsSharedCellElement Then
        oldOrigin = el.AsSharedCellElement.Origin
    Else
        Exit Sub
    End If
    Dim delta As Point3d
    delta.X = newOrigin.X - oldOrigin.X
    delta.Y = newOrigin.Y - oldOrigin.Y
    delta.Z = newOrigin.Z - oldOrigin.Z
    Call el.Move(delta)
    Call el.Rewrite
End Sub

' Move el along dir only so its bbox inward edge (center - dir*half)
' coincides with tip. Lateral position is left alone (shaft alignment).
Private Sub SnapInwardEdgeToTip(el As Element, tip As Point3d, dirX As Double, dirY As Double)
    Dim inPt As Point3d
    inPt = ExtremePointAlongDirPreferColor(el, dirX, dirY, False, 6)
    Dim along As Double
    along = (tip.X - inPt.X) * dirX + (tip.Y - inPt.Y) * dirY
    If Abs(along) < 0.0000001 Then Exit Sub
    Dim delta As Point3d
    delta.X = dirX * along
    delta.Y = dirY * along
    delta.Z = 0#
    Call el.Move(delta)
    Call el.Rewrite
End Sub

Private Function FindNewestElement() As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim maxID As Double: maxID = -1
    Dim el As Element
    Dim newest As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > maxID Then
            maxID = idVal
            Set newest = el
        End If
    Loop
    Set FindNewestElement = newest
End Function

' Inserts signSize (e.g. 48" x 48") via TEXTEDITOR PLAYCOMMAND INSERT_TEXT
' without doubled inch marks. Each " is its own INSERT_TEXT """ keyin
' (three quote chars = one inch mark inside delimiters) -- confirmed live
' 2026-08-03 against CONNECT (doubling produced literal doubles).
Private Sub InsertTextWithInchMarks(sizeText As String)
    Dim i As Long
    Dim chunk As String
    Dim ch As String
    chunk = ""
    For i = 1 To Len(sizeText)
        ch = Mid$(sizeText, i, 1)
        If ch = Chr$(34) Then
            If Len(chunk) > 0 Then
                CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & chunk & """"
                chunk = ""
            End If
            CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT " & Chr$(34) & Chr$(34) & Chr$(34)
        Else
            chunk = chunk & ch
        End If
    Next i
    If Len(chunk) > 0 Then
        CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & chunk & """"
    End If
End Sub

' ============================================================
' CURRENT VIEW ROTATION, IN DEGREES (Z-axis / plan rotation)
' Reads the view's rotation matrix (captured BEFORE it gets reset to
' identity) so signs can be placed at the same angle the engineer is
' currently viewing the drawing at, per the CadInputQueue "ACTIVE ANGLE"
' keyin convention (degrees, CCW positive).
' ============================================================
Public Function ViewRotationAngleDegrees(v As View) As Double
    Dim rot As Matrix3d
    rot = v.Rotation
    Const PI As Double = 3.14159265358979
    ViewRotationAngleDegrees = Atan2(rot.RowY.X, rot.RowX.X) * 180# / PI
End Function

' ============================================================
' ATAN2 -- VBA has no built-in atan2; standard quadrant-correct
' implementation in terms of Atn.
' ============================================================
Private Function Atan2(y As Double, x As Double) As Double
    Const PI As Double = 3.14159265358979
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        Atan2 = Atn(y / x) + PI
    ElseIf x < 0 And y < 0 Then
        Atan2 = Atn(y / x) - PI
    ElseIf x = 0 And y > 0 Then
        Atan2 = PI / 2
    ElseIf x = 0 And y < 0 Then
        Atan2 = -PI / 2
    Else
        Atan2 = 0
    End If
End Function

' ============================================================
' DRAW ARC CONNECTING TWO SIGN POSTS (Both Sides only)
' ============================================================
Sub DrawConnectingArc(startPt As Point3d, endPt As Point3d)
    Dim point As Point3d
    Dim midPoint As Point3d
    Dim dx As Double, dy As Double
    Dim distance As Double
    Dim pX As Double, pY As Double
    Dim arcDepth As Double

    dx = endPt.X - startPt.X
    dy = endPt.Y - startPt.Y
    distance = Sqr(dx * dx + dy * dy)

    midPoint.X = (startPt.X + endPt.X) / 2
    midPoint.Y = (startPt.Y + endPt.Y) / 2
    midPoint.Z = (startPt.Z + endPt.Z) / 2

    arcDepth = distance * 0.1
    If distance > 0 Then
        pX = -dy / distance
        pY = dx / distance
    Else
        pX = 0: pY = 0
    End If

    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"

    point.X = startPt.X: point.Y = startPt.Y: point.Z = startPt.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = endPt.X: point.Y = endPt.Y: point.Z = endPt.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = midPoint.X + (pX * arcDepth)
    point.Y = midPoint.Y + (pY * arcDepth)
    point.Z = midPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset
End Sub
