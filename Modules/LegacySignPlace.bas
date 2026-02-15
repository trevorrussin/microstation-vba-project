Attribute VB_Name = "Module2"
'   - upperPoint: First placement point
'   - lowerPoint: Second placement point
'   - bothSides: True to place on both sides, False for one side
' ============================================================
Public Sub PlaceSignFromLibrary(SignNumber As String, _
                                upperPoint As Point3d, _
                                lowerPoint As Point3d, _
                                bothSides As Boolean)
    
    Dim signData As signData
    
    ' Initialize library if not already done
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    ' Get sign data from library
    signData = GetSignData(SignNumber)
    
    ' Check if sign was found
    If signData.SignNumber = "" Then
        CadInputQueue.SendKeyin "ECHO ERROR: Sign " & SignNumber & " not found in library"
        Exit Sub
    End If
    
    ' Place signs based on bothSides parameter
    If bothSides Then
        ' Place on both sides
        Call PlaceSingleSign(signData, upperPoint, "UPPER")
        Call PlaceSingleSign(signData, lowerPoint, "LOWER")
        Call DrawConnectingArcFromLibrary(upperPoint, lowerPoint)
    Else
        ' Place on one side only (use upperPoint)
        Call PlaceSingleSign(signData, upperPoint, "SINGLE")
    End If
    
    CadInputQueue.SendKeyin "ECHO Sign " & SignNumber & " placed successfully"
    
End Sub

' ============================================================
' PLACE SINGLE SIGN
' Places one sign with its text label and post
' ============================================================
Private Sub PlaceSingleSign(signData As signData, _
                           location As Point3d, _
                           placement As String)
    
    Dim point As Point3d
    Dim textOffsetY As Double
    
    ' Determine text offset based on placement type
    Select Case placement
        Case "UPPER"
            textOffsetY = 50#  ' Text above sign
        Case "LOWER"
            textOffsetY = -50# ' Text below sign
        Case "SINGLE"
            textOffsetY = 50#  ' Text above sign (default)
    End Select
    
    ' --- Attach sign face library ---
    CadInputQueue.SendCommand "ATTACH LIBRARY " & signData.CellLibraryPath
    
    ' --- Place text label ---
    CadInputQueue.SendCommand "TEXTEDITOR PLACE"
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & signData.TextLabel & """"
    
    ' Add second line of text if it exists
    If signData.TextLine2 <> "" Then
        CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
        CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & signData.TextLine2 & """"
    End If
    
    ' Place text at offset position
    point.X = location.X
    point.Y = location.Y + textOffsetY
    point.Z = location.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
    ' --- Place sign cell ---
    SetCExpressionValue "tcb->activeCellUtf16", signData.CellName, ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    point.X = location.X
    point.Y = location.Y
    point.Z = location.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
    ' --- Draw sign post ---
    Call DrawSignPostFromLibrary(signData, location, placement)
    
End Sub

' ============================================================
' DRAW SIGN POST FROM LIBRARY
' Draws the vertical line and post cell using library data
' ============================================================
Private Sub DrawSignPostFromLibrary(signData As signData, _
                                    location As Point3d, _
                                    placement As String)
    
    Dim point As Point3d
    Dim lngTemp As Long
    Dim direction As String
    
    ' Determine direction based on placement
    If placement = "UPPER" Then
        direction = "DOWN"
    Else
        direction = "UP"
    End If
    
    ' --- Draw vertical connecting line (20 feet) ---
    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
    
    If direction = "UP" Then
        ' Line starts 20 feet below and goes UP to sign bottom
        point.X = location.X
        point.Y = location.Y - 20#
        point.Z = location.Z
        CadInputQueue.SendDataPoint point, 1
        
        point.X = location.X
        point.Y = location.Y
        point.Z = location.Z
        CadInputQueue.SendDataPoint point, 1
    Else ' direction = "DOWN"
        ' Line starts at sign bottom and goes DOWN 20 feet
        point.X = location.X
        point.Y = location.Y
        point.Z = location.Z
        CadInputQueue.SendDataPoint point, 1
        
        point.X = location.X
        point.Y = location.Y - 20#
        point.Z = location.Z
        CadInputQueue.SendDataPoint point, 1
    End If
    
    CadInputQueue.SendReset
    
    ' --- Set up cell filter ---
    On Error Resume Next
    lngTemp = Not 1
    lngTemp = GetCExpressionValue("cellMaint_filterEnable", "MGDSHOOK") And lngTemp
    If Err.Number <> 0 Then lngTemp = 0
    Err.Clear
    On Error GoTo 0
    SetCExpressionValue "cellMaint_filterEnable", lngTemp Or 0, "MGDSHOOK"
    
    ' --- Attach post library ---
    CadInputQueue.SendCommand "ATTACH LIBRARY " & signData.PostLibraryPath
    SetCExpressionValue "tcb->activeCellUtf16", signData.PostType, ""
    
    ' --- Place post cell (always 20 feet below sign bottom) ---
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendCommand "LOCK SNAP PERPENDICULAR"
    point.X = location.X
    point.Y = location.Y - 20#
    point.Z = location.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendReset
    
    ' --- Re-attach sign face library for next operations ---
    CadInputQueue.SendCommand "ATTACH LIBRARY " & signData.CellLibraryPath
    
End Sub

' ============================================================
' DRAW CONNECTING ARC FROM LIBRARY
' Draws arc between two sign posts (used for both-sides placement)
' ============================================================
Private Sub DrawConnectingArcFromLibrary(startPt As Point3d, endPt As Point3d)
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
    
    ' Calculate midpoint
    midPoint.X = (startPt.X + endPt.X) / 2
    midPoint.Y = (startPt.Y + endPt.Y) / 2
    midPoint.Z = (startPt.Z + endPt.Z) / 2
    
    ' Calculate perpendicular offset for arc (10% of distance for gentle curve)
    arcDepth = distance * 0.1
    
    ' Calculate perpendicular vector (rotate 90 degrees)
    If distance > 0 Then
        perpX = -dy / distance
        perpY = dx / distance
    Else
        perpX = 0
        perpY = 0
    End If
    
    ' Arc connection points are 20 feet below each sign
    Dim arcStartY As Double
    Dim arcEndY As Double
    arcStartY = startPt.Y - 20#
    arcEndY = endPt.Y - 20#
    
    ' --- Set up arc placement mode ---
    CadInputQueue.SendCommand "PLACE ARC ICON"
    SetCExpressionValue "tcb->msToolSettings.igen.placeArcModeEx", 3, "CONSGEOM"
    CadInputQueue.SendCommand "PLACE ARC ICON"
    
    ' --- Place arc with three points ---
    ' First point (at bottom of first sign post)
    point.X = startPt.X
    point.Y = arcStartY
    point.Z = startPt.Z
    CadInputQueue.SendDataPoint point, 1
    
    ' Second point (at bottom of second sign post)
    point.X = endPt.X
    point.Y = arcEndY
    point.Z = endPt.Z
    CadInputQueue.SendDataPoint point, 1
    
    ' Third point (offset perpendicular at midpoint for arc curvature)
    point.X = midPoint.X + (perpX * arcDepth)
    point.Y = midPoint.Y - 20# + (perpY * arcDepth)
    point.Z = midPoint.Z
    CadInputQueue.SendDataPoint point, 1
    
    CadInputQueue.SendReset
    
End Sub

' ============================================================
' PLACE MULTIPLE SIGNS FROM TABLE
' Places multiple signs based on user's table input
' Parameters:
'   - signNumbers(): Array of sign numbers from table
'   - spacings(): Array of spacing values (in feet)
'   - sides(): Array indicating "ONE" or "BOTH"
'   - basePoint: Starting point for first sign
'   - direction: "HORIZONTAL" or "VERTICAL" for sign layout
' ============================================================
Public Sub PlaceMultipleSignsFromTable(signNumbers() As String, _
                                       spacings() As Double, _
                                       sides() As String, _
                                       basePoint As Point3d, _
                                       direction As String)
    
    Dim i As Integer
    Dim currentPoint As Point3d
    Dim upperPoint As Point3d
    Dim lowerPoint As Point3d
    Dim offset As Double
    Dim bothSides As Boolean
    
    ' Initialize library
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    currentPoint = basePoint
    offset = 0
    
    ' Loop through each sign in the table
    For i = LBound(signNumbers) To UBound(signNumbers)
        
        ' Skip empty entries
        If signNumbers(i) <> "" Then
            
            ' Determine if placing on both sides
            bothSides = (UCase(sides(i)) = "BOTH")
            
            ' Calculate current position based on accumulated spacing
            If direction = "HORIZONTAL" Then
                currentPoint.X = basePoint.X + offset
                currentPoint.Y = basePoint.Y
            Else ' VERTICAL
                currentPoint.X = basePoint.X
                currentPoint.Y = basePoint.Y + offset
            End If
            currentPoint.Z = basePoint.Z
            
            ' Set up points for both-sides placement
            If bothSides Then
                upperPoint = currentPoint
                If direction = "HORIZONTAL" Then
                    lowerPoint.X = currentPoint.X
                    lowerPoint.Y = currentPoint.Y - 100# ' 100 feet offset perpendicular
                    lowerPoint.Z = currentPoint.Z
                Else
                    lowerPoint.X = currentPoint.X + 100# ' 100 feet offset perpendicular
                    lowerPoint.Y = currentPoint.Y
                    lowerPoint.Z = currentPoint.Z
                End If
            Else
                upperPoint = currentPoint
                lowerPoint = currentPoint
            End If
            
            ' Place the sign
            Call PlaceSignFromLibrary(signNumbers(i), upperPoint, lowerPoint, bothSides)
            
            ' Update offset for next sign
            offset = offset + spacings(i)
            
        End If
        
    Next i
    
    CommandState.StartDefaultCommand
    CadInputQueue.SendKeyin "ECHO All signs placed successfully"
    
End Sub

' ============================================================
' INTERACTIVE SIGN PLACEMENT
' Allows user to select sign and placement interactively
' ============================================================
Public Sub PlaceSignInteractive(SignNumber As String)
    Dim startPoint As Point3d
    Dim endPoint As Point3d
    Dim oMessage As CadInputMessage
    Dim upperPoint As Point3d
    Dim lowerPoint As Point3d
    
    ' Initialize library
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    ' Verify sign exists
    If Not SignExists(SignNumber) Then
        CadInputQueue.SendKeyin "ECHO ERROR: Sign " & SignNumber & " not found in library"
        Exit Sub
    End If
    
    ' --- Force active view to be unrotated ---
    Dim v As View
    Set v = ActiveDesignFile.Views(1)
    v.Rotation = Matrix3dIdentity
    v.Redraw
    
    ' --- Force world coordinates ---
    CadInputQueue.SendKeyin "ACS SET WORLD"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"
    CadInputQueue.SendKeyin "LOCK ROTATION OFF"
    
    ' Prompt user for first location
    CadInputQueue.SendKeyin "ECHO Select location for FIRST sign post (sign: " & SignNumber & ")"
    CadInputQueue.SendCommand "NULL"
    
    ' Get first data point
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
    
    ' Prompt user for second location
    CadInputQueue.SendKeyin "ECHO Select location for SECOND sign post (or reset for single sign)"
    CadInputQueue.SendCommand "NULL"
    
    ' Get second data point (or reset for single placement)
    Set oMessage = CadInputQueue.GetInput
    Do While oMessage.InputType <> msdCadInputTypeDataPoint And oMessage.InputType <> msdCadInputTypeReset
        Set oMessage = CadInputQueue.GetInput
    Loop
    
    If oMessage.InputType = msdCadInputTypeReset Then
        ' Single sign placement
        Call PlaceSignFromLibrary(SignNumber, startPoint, startPoint, False)
    Else
        ' Both sides placement
        endPoint = oMessage.point
        
        ' Determine which point has higher Y coordinate
        If startPoint.Y > endPoint.Y Then
            upperPoint = startPoint
            lowerPoint = endPoint
        Else
            upperPoint = endPoint
            lowerPoint = startPoint
        End If
        
        Call PlaceSignFromLibrary(SignNumber, upperPoint, lowerPoint, True)
    End If
    
    CommandState.StartDefaultCommand
    
End Sub

' ============================================================
' HELPER: Get Sign Description
' ============================================================
Public Function GetSignDescription(SignNumber As String) As String
    Dim signData As signData
    
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    signData = GetSignData(SignNumber)
    GetSignDescription = signData.Description
End Function

' ============================================================
' HELPER: Get Sign Dimensions
' ============================================================
Public Function GetSignDimensions(SignNumber As String) As String
    Dim signData As signData
    
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    signData = GetSignData(SignNumber)
    
    If signData.SignNumber <> "" Then
        GetSignDimensions = signData.WidthInches & """ x " & signData.HeightInches & """"
    Else
        GetSignDimensions = "N/A"
    End If
End Function


