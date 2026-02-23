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
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
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

        ' Legacy order: text label → sign face cell → post
        Call PlaceSignFaceAndText(pt1, signNum, signSize, d1X, d1Y)
        Call DrawSignPost(pt1, d1X, d1Y)

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

        ' Legacy order: text label → sign face cell → post
        Call PlaceSignFaceAndText(pt1, signNum, signSize, dAX, dAY)
        Call DrawSignPost(pt1, dAX, dAY)

        Call PlaceSignFaceAndText(pt2, signNum, signSize, dBX, dBY)
        Call DrawSignPost(pt2, dBX, dBY)

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
' DRAW SIGN POST CELL + 20-FT POST LINE
' postPt = (projected) post base location
' dirX/Y = unit vector pointing outward (toward sign face)
' ============================================================
Sub DrawSignPost(postPt As Point3d, dirX As Double, dirY As Double)
    Dim point As Point3d

    ' 20-ft line from post base to bottom of sign face
    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
    point.X = postPt.X: point.Y = postPt.Y: point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    point.X = postPt.X + dirX * 20#
    point.Y = postPt.Y + dirY * 20#
    point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset

    ' Post cell at base location
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    SetCExpressionValue "tcb->activeCellUtf16", "TWZSGN_P", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = postPt.X: point.Y = postPt.Y: point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset

    ' Re-attach sign face library so next PlaceSignFaceAndText uses correct library
    If currentSignFaceLibraryPath = "" Then
        currentSignFaceLibraryPath = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    End If
    CadInputQueue.SendCommand "ATTACH LIBRARY " & currentSignFaceLibraryPath
End Sub

' ============================================================
' PLACE SIGN FACE CELL AND TWO-LINE TEXT LABEL
' Follows Legacy pattern order: text label first, then sign face cell.
' Sign face is placed 20 ft from post in dirX/Y direction.
' Text (sign number + size) is placed 70 ft from post.
' ============================================================
Sub PlaceSignFaceAndText(postPt As Point3d, signNum As String, signSize As String, _
                          dirX As Double, dirY As Double)
    Dim point As Point3d

    ' --- Text label at 70 ft: sign number (line 1) + size (line 2) ---
    ' Escape " (inch marks) as "" for TEXTEDITOR keyin (matches Legacy pattern)
    Dim escapedSize As String
    escapedSize = Replace(signSize, Chr(34), Chr(34) & Chr(34))

    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & signNum & """"
    If Len(escapedSize) > 0 Then
        CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
        CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & escapedSize & """"
    End If
    point.X = postPt.X + dirX * 70#
    point.Y = postPt.Y + dirY * 70#
    point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset

    ' --- Sign face cell at 20 ft (same for every sign in library) ---
    ' Re-attach sign face library so it is active before placing cell
    If Len(currentSignFaceLibraryPath) = 0 Then
        currentSignFaceLibraryPath = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    End If
    CadInputQueue.SendCommand "ATTACH LIBRARY " & currentSignFaceLibraryPath
    ' Cell name from SignLibrary; warn and skip face if sign is not in library
    Dim cellName As String
    If SignLibrary.SignExists(signNum) Then
        cellName = SignLibrary.GetSignData(signNum).CellName
    Else
        CadInputQueue.SendKeyin "ECHO WARNING: Sign " & signNum & " not found in library — face cell skipped"
        cellName = ""
    End If
    If Len(cellName) > 0 Then
        SetCExpressionValue "tcb->activeCellUtf16", cellName, ""
        CadInputQueue.SendCommand "PLACE CELL ICON"
        point.X = postPt.X + dirX * 20#
        point.Y = postPt.Y + dirY * 20#
        point.Z = postPt.Z
        CadInputQueue.SendDataPoint point, 1
        CadInputQueue.SendReset
    End If
End Sub

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
