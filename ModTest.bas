Attribute VB_Name = "ModTest"
Option Explicit

' ============================================================
' PLACE WORKZONE SIGN
' ------------------------------------------------------------
' Called by ModuleSignPlacement.DrawCurrentSign with the
' geometry of the perpendicular line where the sign goes.
'
' Parameters:
'   signNum   - sign number string (e.g. "W6-3")
'   signSize  - sign size string from sign table (may contain " chars)
'   side      - "One Side" or "Both Sides"
'   midX/Y/Z  - alignment point = midpoint of the perpendicular line
'   perpX/Y   - unit perpendicular vector (perpendicular to alignment)
'
' Behaviour:
'   One Side   - collects 1 click, projects onto perp line,
'                draws: 20-ft post line + post cell + sign face cell + text.
'                No arc.
'   Both Sides - collects 2 clicks, projects each onto perp line,
'                draws both signs and a connecting arc between the posts.
' ============================================================
Sub PlaceWorkZoneSign(signNum As String, signSize As String, side As String, _
                       midX As Double, midY As Double, midZ As Double, _
                       perpX As Double, perpY As Double)

    Const HALF_LEN As Double = 20   ' matches PERP_HALF_LEN in ModuleAlignmentPlacement

    ' Setup view
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"

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

        Call DrawSignPost(pt1, d1X, d1Y)
        Call PlaceSignFaceAndText(pt1, signNum, signSize, d1X, d1Y)

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

        Call DrawSignPost(pt1, dAX, dAY)
        Call PlaceSignFaceAndText(pt1, signNum, signSize, dAX, dAY)

        Call DrawSignPost(pt2, dBX, dBY)
        Call PlaceSignFaceAndText(pt2, signNum, signSize, dBX, dBY)

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

    ' Re-attach sign face library for subsequent operations
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
End Sub

' ============================================================
' PLACE SIGN FACE CELL AND TWO-LINE TEXT LABEL
' Sign face is placed 20 ft from post in dirX/Y direction.
' Text (sign number + size) is placed 70 ft from post.
' ============================================================
Sub PlaceSignFaceAndText(postPt As Point3d, signNum As String, signSize As String, _
                          dirX As Double, dirY As Double)
    Dim point As Point3d

    ' Sign face cell at 20 ft
    SetCExpressionValue "tcb->activeCellUtf16", "R02-10sNY", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = postPt.X + dirX * 20#
    point.Y = postPt.Y + dirY * 20#
    point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset

    ' Text label at 70 ft: sign number (line 1) + size (line 2)
    ' Replace " (inch symbol) with ' to avoid TEXTEDITOR keyin quoting issues
    Dim cleanSize As String
    cleanSize = Replace(signSize, Chr(34), Chr(39))

    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & signNum & """"
    If cleanSize <> "" Then
        CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
        CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & cleanSize & """"
    End If
    point.X = postPt.X + dirX * 70#
    point.Y = postPt.Y + dirY * 70#
    point.Z = postPt.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
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
