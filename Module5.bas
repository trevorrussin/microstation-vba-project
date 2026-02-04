Attribute VB_Name = "Module5"
Option Explicit

Sub PlaceWorkZoneSign()
    Dim startPoint As Point3d
    Dim point As Point3d
    Dim oMessage As CadInputMessage
    
'   Coordinates are in master units
    startPoint.X = 1031175.43628106
    startPoint.Y = 243097.650582335
    startPoint.Z = 0#
    
'   Start a command
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    
'   ===== SIGN 1: R02-10sNY (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""NYR09-11"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "R02-10sNY", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X, startPoint.Y, 0#)
    
'   ===== SIGN 2: R02-10sNY (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""NYR09-11"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "R02-10sNY", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 1 and 2
    Call DrawConnectingArc(startPoint.X, startPoint.Y, startPoint.X, startPoint.Y + 200#)
    
'   ===== SIGN 3: W20-01RA (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W20-01RA"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 100#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W20-01RA", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 100#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 100#, startPoint.Y, 0#)
    
'   ===== SIGN 4: W20-01RA (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W20-01RA"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 100#
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W20-01RA", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 100#
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 100#, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 3 and 4
    Call DrawConnectingArc(startPoint.X + 100#, startPoint.Y, startPoint.X + 100#, startPoint.Y + 200#)
    
'   ===== SIGN 5: R02-01 (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""R02-01"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""30"""" x 36"""""
    point.X = startPoint.X + 200#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "R02-01", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 200#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 200#, startPoint.Y, 0#)
    
'   ===== SIGN 6: R02-01 (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""R02-01"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""30"""" x 36"""""
    point.X = startPoint.X + 200#
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "R02-01", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 200#
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 200#, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 5 and 6
    Call DrawConnectingArc(startPoint.X + 200#, startPoint.Y, startPoint.X + 200#, startPoint.Y + 200#)
    
'   ===== SIGN 7: W20-05RA (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W20-05RA"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 300#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W20-05RA", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 300#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 300#, startPoint.Y, 0#)
    
'   ===== SIGN 8: W20-05RA (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W20-05RA"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 300#
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W20-05RA", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 300#
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 300#, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 7 and 8
    Call DrawConnectingArc(startPoint.X + 300#, startPoint.Y, startPoint.X + 300#, startPoint.Y + 200#)
    
'   ===== SIGN 9: W04-02R (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W04-02R"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 400#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W04-02R", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 400#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 400#, startPoint.Y, 0#)
    
'   ===== SIGN 10: W04-02R (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""W04-02R"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 400#
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "W04-02R", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 400#
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 400#, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 9 and 10
    Call DrawConnectingArc(startPoint.X + 400#, startPoint.Y, startPoint.X + 400#, startPoint.Y + 200#)
    
'   ===== SIGN 11: OM03-R (ONLY LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""OM03-R"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""12"""" x 36"""""
    point.X = startPoint.X + 500#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "OM03-R", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 500#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 500#, startPoint.Y, 0#)
    
'   ===== SIGN 12: G20-02 (LOWER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""G20-02"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 600#
    point.Y = startPoint.Y - 50#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "G20-02", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 600#
    point.Y = startPoint.Y
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 600#, startPoint.Y, 0#)
    
'   ===== SIGN 13: G20-02 (UPPER) =====
'   Place text label
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""G20-02"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""48"""" x 48"""""
    point.X = startPoint.X + 600#
    point.Y = startPoint.Y + 250#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
'   Place sign
    SetCExpressionValue "tcb->activeCellUtf16", "G20-02", ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = startPoint.X + 600#
    point.Y = startPoint.Y + 200#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    Call DrawSignPost(startPoint.X + 600#, startPoint.Y + 200#, 0#)
    
'   Draw arc connecting signs 12 and 13
    Call DrawConnectingArc(startPoint.X + 600#, startPoint.Y, startPoint.X + 600#, startPoint.Y + 200#)
    
    CommandState.StartDefaultCommand
    
End Sub

Sub DrawSignPost(xCoord As Double, yCoord As Double, zCoord As Double)
    Dim point As Point3d
    
'   Draw vertical line connecting post to sign face bottom (20 feet down from sign)
    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
    
    point.X = xCoord
    point.Y = yCoord
    point.Z = zCoord
    CadInputQueue.SendDataPoint point, 1
    point.X = xCoord
    point.Y = yCoord - 20#
    point.Z = zCoord
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
'   Attach sign post library
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    SetCExpressionValue "tcb->activeCellUtf16", "TWZSGN_P", ""
    
'   Place post cell 20 feet below the sign
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendCommand "LOCK SNAP PERPENDICULAR"
    point.X = xCoord
    point.Y = yCoord - 20#
    point.Z = zCoord
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
'   Re-attach sign face library for next operations
    CadInputQueue.SendCommand "ATTACH LIBRARY c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    
End Sub

Sub DrawConnectingArc(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    Dim point As Point3d
    Dim midY As Double
    Dim arcDepth As Double
    
    ' Calculate midpoint Y and arc depth
    midY = (y1 + y2) / 2
    arcDepth = Abs(y2 - y1) * 0.1
    
'   Set up arc placement mode - use 3-point arc
    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"
    
'   First point: lower post (20 feet below sign)
    point.X = x1
    point.Y = y1 - 20#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    
'   Second point: upper post (20 feet below sign)
    point.X = x2
    point.Y = y2 - 20#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    
'   Third point: offset to left for curve
    point.X = x1 - arcDepth
    point.Y = midY - 20#
    point.Z = 0#
    CadInputQueue.SendDataPoint point, 1
    
    CadInputQueue.SendReset
    
End Sub
