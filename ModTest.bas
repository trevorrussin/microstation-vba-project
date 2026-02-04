Attribute VB_Name = "ModTest"
Option Explicit

Sub PlaceWorkZoneSign()
    Dim startPoint As Point3d
    Dim endPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long
    Dim oMessage As CadInputMessage
    
    ' --- Force active view to be unrotated ---
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw

    '   --- FIX: force world coordinates ---
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"
    
    ' Prompt user to select first sign post location
    CadInputQueue.SendKeyin "ECHO Select location for FIRST sign post"
    CadInputQueue.SendCommand "NULL"
    
    ' Get first data point from user
    Set oMessage = CadInputQueue.GetInput
    
    Do While oMessage.InputType <> msdCadInputTypeDataPoint
        Set oMessage = CadInputQueue.GetInput
        If oMessage.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendKeyin "ECHO Operation cancelled"
            CommandState.StartDefaultCommand
            Exit Sub
        End If
    Loop
    
    startPoint = oMessage.point
    
    ' Start a line command to show dynamic feedback
    CadInputQueue.SendCommand "PLACE LINE"
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1
    
    ' Prompt user to select second sign post location
    CadInputQueue.SendKeyin "ECHO Select location for SECOND sign post"
    
    ' Get second data point from user (with dynamic line feedback)
    Set oMessage = CadInputQueue.GetInput
    
    Do While oMessage.InputType <> msdCadInputTypeDataPoint
        Set oMessage = CadInputQueue.GetInput
        If oMessage.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendKeyin "ECHO Operation cancelled"
            CadInputQueue.SendReset
            CommandState.StartDefaultCommand
            Exit Sub
        End If
    Loop
    
    endPoint = oMessage.point
    
    ' Cancel the line command and clear it
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    
'   Determine which point has higher Y coordinate
    Dim upperPoint As Point3d
    Dim lowerPoint As Point3d
    
    If startPoint.Y > endPoint.Y Then
        upperPoint = startPoint
        lowerPoint = endPoint
    Else
        upperPoint = endPoint
        lowerPoint = startPoint
    End If
    
'   Start placing signs and posts
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    
'   ===== SIGN AT UPPER POINT (higher Y) =====
'   This sign face will be placed 20 feet ABOVE the clicked point
'   Place text label above the sign
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""NYR09-11"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = upperPoint.X
    point.Y = upperPoint.Y + 70#
    point.Z = upperPoint.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign face (bottom center of sign will be 20 feet above the post)
    SetCExpressionValue "tcb->activeCellUtf16", "R02-10sNY", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = upperPoint.X
    point.Y = upperPoint.Y + 20#
    point.Z = upperPoint.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
'   ===== SIGN AT LOWER POINT (lower Y) =====
'   This sign face will be placed 20 feet BELOW the clicked point
'   Place text label below the sign
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""NYR09-11"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = lowerPoint.X
    point.Y = lowerPoint.Y - 70#
    point.Z = lowerPoint.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign face (bottom center of sign will be 20 feet below the post)
    SetCExpressionValue "tcb->activeCellUtf16", "R02-10sNY", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = lowerPoint.X
    point.Y = lowerPoint.Y - 20#
    point.Z = lowerPoint.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
'   ===== DRAW SIGN POSTS AND CONNECTING LINES =====
'   Upper point: post at clicked point, line goes UP 20 feet to bottom of sign
    Call DrawSignPost(upperPoint.X, upperPoint.Y, upperPoint.Z, "UP")
'   Lower point: post at clicked point, line goes DOWN 20 feet to bottom of sign
    Call DrawSignPost(lowerPoint.X, lowerPoint.Y, lowerPoint.Z, "DOWN")
    
'   ===== DRAW ARC CONNECTING THE TWO SIGN POSTS =====
    Call DrawConnectingArc(upperPoint, lowerPoint)
    
    CommandState.StartDefaultCommand
    
    CadInputQueue.SendKeyin "ECHO Sign posts and connecting arc placed successfully!"
    
End Sub

Sub DrawSignPost(xCoord As Double, yCoord As Double, zCoord As Double, direction As String)
    Dim point As Point3d
    
'   Draw vertical line connecting post to sign face bottom
'   The clicked point (xCoord, yCoord) is where the post is located
'   Direction: "UP" means line goes from post UP (+Y) 20 feet to bottom of sign face
'   Direction: "DOWN" means line goes from post DOWN (-Y) 20 feet to bottom of sign face
    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
    
    If direction = "UP" Then
        ' Post is at clicked point, line goes UP 20 feet to bottom of sign
        point.X = xCoord
        point.Y = yCoord
        point.Z = zCoord
        CadInputQueue.SendDataPoint point, 1
        point.X = xCoord
        point.Y = yCoord + 20#
        point.Z = zCoord
        CadInputQueue.SendDataPoint point, 1
    Else ' direction = "DOWN"
        ' Post is at clicked point, line goes DOWN 20 feet to bottom of sign
        point.X = xCoord
        point.Y = yCoord
        point.Z = zCoord
        CadInputQueue.SendDataPoint point, 1
        point.X = xCoord
        point.Y = yCoord - 20#
        point.Z = zCoord
        CadInputQueue.SendDataPoint point, 1
    End If
    
    CadInputQueue.SendReset
    
'   Attach sign post library
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    SetCExpressionValue "tcb->activeCellUtf16", "TWZSGN_P", ""
    
'   Place post cell at the clicked location (where the post is)
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendCommand "LOCK SNAP PERPENDICULAR"
    point.X = xCoord
    point.Y = yCoord
    point.Z = zCoord
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
'   Re-attach sign face library for next operations
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    
End Sub

Sub DrawConnectingArc(startPt As Point3d, endPt As Point3d)
    Dim point As Point3d
    Dim midPoint As Point3d
    Dim dx As Double, dy As Double
    Dim distance As Double
    Dim perpX As Double, perpY As Double
    Dim arcDepth As Double
    
    ' Calculate vector between the two posts
    dx = endPt.X - startPt.X
    dy = endPt.Y - startPt.Y
    distance = Sqr(dx * dx + dy * dy)
    
    ' Calculate midpoint between the two posts
    midPoint.X = (startPt.X + endPt.X) / 2
    midPoint.Y = (startPt.Y + endPt.Y) / 2
    midPoint.Z = (startPt.Z + endPt.Z) / 2
    
    ' Offset for the arc (10% of distance for gentle curve)
    arcDepth = distance * 0.1
    
    ' Calculate perpendicular vector (rotate 90 degrees)
    If distance > 0 Then
        perpX = -dy / distance
        perpY = dx / distance
    Else
        perpX = 0
        perpY = 0
    End If
    
'   Set up arc placement mode - use 3-point arc
    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"
    
'   Place arc with three points connecting the two posts
'   First point: upper post location (startPt has higher Y)
    point.X = startPt.X
    point.Y = startPt.Y
    point.Z = startPt.Z
    CadInputQueue.SendDataPoint point, 1
    
'   Second point: lower post location (endPt has lower Y)
    point.X = endPt.X
    point.Y = endPt.Y
    point.Z = endPt.Z
    CadInputQueue.SendDataPoint point, 1
    
'   Third point: offset perpendicular at midpoint to create curve
    point.X = midPoint.X + (perpX * arcDepth)
    point.Y = midPoint.Y + (perpY * arcDepth)
    point.Z = midPoint.Z
    CadInputQueue.SendDataPoint point, 1
    
    CadInputQueue.SendReset
    
End Sub
