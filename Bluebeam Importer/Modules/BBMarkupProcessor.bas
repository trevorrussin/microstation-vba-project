Option Explicit

' ============================================================
' BBMarkupProcessor.bas
' Executes MicroStation operations for each classified BBMarkup.
' Uses CadInputQueue patterns from the existing WZTC project.
'
' Public entry point:
'   ProcessMarkup(idx)  — dispatches to operation-specific Sub
'
' All changes to bbMarkups(idx).Status and .StatusNote must be
' written back to the array (VBA UDTs are value types, not refs).
' ============================================================

' Default cell library path — same as WZTC project
Private Const WZTC_CELL_LIB As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"

' Search radius in MicroStation design units (feet) for FindNearestElement
Private Const SEARCH_RADIUS As Double = 20#

' ============================================================
' ProcessMarkup — top-level dispatcher
' ============================================================
Public Sub ProcessMarkup(idx As Integer)
    If idx < 1 Or idx > bbMarkupCount Then Exit Sub

    Dim m As BBMarkup
    m = bbMarkups(idx)

    If m.Status = "Done" Or m.Status = "Skipped" Then Exit Sub

    On Error GoTo ProcessErr

    Select Case m.OpType
        Case "DELETE":          ExecuteDelete idx
        Case "MOVE":            ExecuteMove idx
        Case "ADD_CELL":        ExecuteAddCell idx
        Case "CHANGE_LEVEL":    ExecuteChangeLevel idx
        Case "EDIT_TEXT":       ExecuteEditText idx
        Case "ADD_DIMENSION":   ExecuteAddDimension idx
        Case "ADD_CALLOUT":     ExecuteAddCallout idx
        Case "DELETE_CALLOUT":  ExecuteDeleteCallout idx
        Case Else
            bbMarkups(idx).Status = "Skipped"
            bbMarkups(idx).StatusNote = "No matching operation for: " & m.OpType
    End Select
    Exit Sub

ProcessErr:
    bbMarkups(idx).Status = "Error"
    bbMarkups(idx).StatusNote = "Runtime error: " & Err.Description
End Sub

' ============================================================
' ExecuteDelete
' Finds the nearest element to the markup location and removes it.
' If no coordinate, prompts user to click the element.
' ============================================================
Private Sub ExecuteDelete(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    Dim el As Element
    If m.HasCoord And bbCalibrated Then
        Set el = FindNearestElement(m.MstnX, m.MstnY, SEARCH_RADIUS)
        If el Is Nothing Then
            ' Widen search once before falling back to user click
            Set el = FindNearestElement(m.MstnX, m.MstnY, SEARCH_RADIUS * 5)
        End If
    End If

    If el Is Nothing Then
        Set el = PromptUserClickElement(m.MarkupID, "Click the element to DELETE: " & m.RawText)
    End If

    If el Is Nothing Then
        bbMarkups(idx).Status = "Skipped"
        bbMarkups(idx).StatusNote = "User cancelled or no element found"
        Exit Sub
    End If

    ActiveModelReference.RemoveElement el
    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Element deleted"
End Sub

' ============================================================
' ExecuteMove
' Finds element, moves it by the direction + distance in Param1/Param2.
' Direction: N/S/E/W/LEFT/RIGHT/UP/DOWN
' Distance : Param2 in feet (MicroStation design units)
' ============================================================
Private Sub ExecuteMove(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    ' Parse distance
    Dim dist As Double
    dist = 0
    If Len(Trim(m.Param2)) > 0 Then
        On Error Resume Next
        dist = CDbl(m.Param2)
        On Error GoTo 0
    End If

    If dist <= 0 Then
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "Could not parse move distance from: " & m.RawText
        Exit Sub
    End If

    ' Compute direction vector
    Dim dx As Double, dy As Double
    dx = 0: dy = 0
    Select Case UCase(Trim(m.Param1))
        Case "N", "NORTH", "UP":    dy = dist
        Case "S", "SOUTH", "DOWN":  dy = -dist
        Case "E", "EAST", "RIGHT":  dx = dist
        Case "W", "WEST", "LEFT":   dx = -dist
        Case Else
            bbMarkups(idx).Status = "Error"
            bbMarkups(idx).StatusNote = "Unknown direction '" & m.Param1 & "'. Use N/S/E/W/Left/Right."
            Exit Sub
    End Select

    ' Find element
    Dim el As Element
    If m.HasCoord And bbCalibrated Then
        Set el = FindNearestElement(m.MstnX, m.MstnY, SEARCH_RADIUS)
    End If
    If el Is Nothing Then
        Set el = PromptUserClickElement(m.MarkupID, "Click the element to MOVE: " & m.RawText)
    End If
    If el Is Nothing Then
        bbMarkups(idx).Status = "Skipped"
        bbMarkups(idx).StatusNote = "User cancelled"
        Exit Sub
    End If

    ' Apply translation using MicroStation Move3d
    Dim delta As Point3d
    delta.X = dx
    delta.Y = dy
    delta.Z = 0
    el.Move delta
    el.Rewrite

    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Moved " & dist & " ft " & m.Param1
End Sub

' ============================================================
' ExecuteAddCell
' Places a cell from the WZTC cell library at the markup location.
' Param1 = cell name (e.g. "TWZAP_P")
' ============================================================
Private Sub ExecuteAddCell(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    If Len(Trim(m.Param1)) = 0 Then
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "No cell name found in markup text"
        Exit Sub
    End If

    ' Determine placement point
    Dim pt As Point3d
    If m.HasCoord And bbCalibrated Then
        pt.X = m.MstnX
        pt.Y = m.MstnY
        pt.Z = 0
    Else
        If Not PromptUserClickPoint(m.MarkupID, "Click placement point for cell " & m.Param1 & ": " & m.RawText, pt) Then
            bbMarkups(idx).Status = "Skipped"
            bbMarkups(idx).StatusNote = "User cancelled"
            Exit Sub
        End If
    End If

    ' Set active properties
    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE 0"

    ' Attach library and place cell
    CadInputQueue.SendCommand "ATTACH LIBRARY " & WZTC_CELL_LIB
    SetCExpressionValue "tcb->activeCellUtf16", m.Param1, ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendDataPoint pt, 1
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Cell placed: " & m.Param1
End Sub

' ============================================================
' ExecuteChangeLevel
' Finds the nearest element and changes its level.
' Param1 = level name (e.g. "TWZBT_P")
' ============================================================
Private Sub ExecuteChangeLevel(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    If Len(Trim(m.Param1)) = 0 Then
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "No level name found in markup text"
        Exit Sub
    End If

    ' Verify level exists
    Dim lvl As Level
    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels(m.Param1)
    On Error GoTo 0
    If lvl Is Nothing Then
        ' Try creating the level on-the-fly (MicroStation will auto-create unknown levels)
        ' Just proceed — the Level assignment below will raise an error if truly invalid
    End If

    ' Find element
    Dim el As Element
    If m.HasCoord And bbCalibrated Then
        Set el = FindNearestElement(m.MstnX, m.MstnY, SEARCH_RADIUS)
    End If
    If el Is Nothing Then
        Set el = PromptUserClickElement(m.MarkupID, "Click element to change level to " & m.Param1 & ": " & m.RawText)
    End If
    If el Is Nothing Then
        bbMarkups(idx).Status = "Skipped"
        bbMarkups(idx).StatusNote = "User cancelled"
        Exit Sub
    End If

    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels(m.Param1)
    If lvl Is Nothing Then
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "Level not found: " & m.Param1
        Exit Sub
    End If
    el.Level = lvl
    el.Rewrite
    On Error GoTo 0

    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Level changed to " & m.Param1
End Sub

' ============================================================
' ExecuteEditText
' Finds the nearest text element and sets its text content.
' Param1 = new text string
' ============================================================
Private Sub ExecuteEditText(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    If Len(Trim(m.Param1)) = 0 Then
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "No replacement text found in markup text"
        Exit Sub
    End If

    ' Find a text element at the markup location
    Dim el As Element
    If m.HasCoord And bbCalibrated Then
        Set el = FindNearestTextElement(m.MstnX, m.MstnY, SEARCH_RADIUS)
    End If
    If el Is Nothing Then
        Set el = PromptUserClickElement(m.MarkupID, "Click text element to edit (new text: " & m.Param1 & ")")
    End If
    If el Is Nothing Then
        bbMarkups(idx).Status = "Skipped"
        bbMarkups(idx).StatusNote = "User cancelled"
        Exit Sub
    End If

    ' Handle both TextElement and TextNodeElement
    On Error Resume Next
    If el.Type = msdElementTypeText Then
        Dim te As TextElement
        Set te = el
        te.Text = m.Param1
        te.Rewrite
    ElseIf el.Type = msdElementTypeTextNode Then
        Dim tn As TextNodeElement
        Set tn = el
        ' TextNodeElement: update first line
        Dim lineEl As TextElement
        Dim lineEnum As ElementEnumerator
        Set lineEnum = tn.GetSubElements
        If lineEnum.MoveNext Then
            Set lineEl = lineEnum.Current
            lineEl.Text = m.Param1
            lineEl.Rewrite
        End If
        tn.Rewrite
    Else
        bbMarkups(idx).Status = "Error"
        bbMarkups(idx).StatusNote = "Selected element is not a text element (type " & el.Type & ")"
        Exit Sub
    End If
    On Error GoTo 0

    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Text updated to: " & Left(m.Param1, 30)
End Sub

' ============================================================
' ExecuteAddDimension
' Activates the linear dimension command.
' User places the dimension manually (interactive).
' ============================================================
Private Sub ExecuteAddDimension(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 2"   ' yellow
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendCommand "DIMENSION LINEAR SIZE ARROW"

    ' Let user place the dimension; do not block here
    ' The user completes placement interactively
    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Dimension command activated — place in drawing"
End Sub

' ============================================================
' ExecuteAddCallout
' Activates the place note / callout command.
' Param1 = callout text (pre-populated if available)
' ============================================================
Private Sub ExecuteAddCallout(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"

    ' Pre-set text if we have it
    If Len(Trim(m.Param1)) > 0 Then
        CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & m.Param1 & """"
    End If

    CadInputQueue.SendCommand "TEXTEDITOR PLACENOTE"

    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Callout command activated — place in drawing"
End Sub

' ============================================================
' ExecuteDeleteCallout
' Finds the nearest text/note element and removes it.
' ============================================================
Private Sub ExecuteDeleteCallout(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    Dim el As Element
    If m.HasCoord And bbCalibrated Then
        Set el = FindNearestTextElement(m.MstnX, m.MstnY, SEARCH_RADIUS)
        If el Is Nothing Then
            Set el = FindNearestElement(m.MstnX, m.MstnY, SEARCH_RADIUS * 3)
        End If
    End If
    If el Is Nothing Then
        Set el = PromptUserClickElement(m.MarkupID, "Click the CALLOUT to delete: " & m.RawText)
    End If
    If el Is Nothing Then
        bbMarkups(idx).Status = "Skipped"
        bbMarkups(idx).StatusNote = "User cancelled"
        Exit Sub
    End If

    ActiveModelReference.RemoveElement el
    bbMarkups(idx).Status = "Done"
    bbMarkups(idx).StatusNote = "Callout deleted"
End Sub

' ============================================================
' FindNearestElement
' Scans all graphic elements; returns the one closest to (x, y)
' within the given radius (in design units). Returns Nothing if none found.
' ============================================================
Private Function FindNearestElement(x As Double, y As Double, _
                                     radius As Double) As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim bestEl As Element
    Dim bestDist As Double
    bestDist = 1E+30
    Set bestEl = Nothing

    Do While oEnum.MoveNext
        Dim el As Element
        Set el = oEnum.Current

        ' Use bounding box center as approximate element center
        On Error Resume Next
        Dim lo As Point3d, hi As Point3d
        lo = el.Range.Low
        hi = el.Range.High
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextEl
        End If
        On Error GoTo 0

        Dim cx As Double, cy As Double
        cx = (lo.X + hi.X) / 2
        cy = (lo.Y + hi.Y) / 2

        Dim dist As Double
        dist = Sqr((cx - x) ^ 2 + (cy - y) ^ 2)

        If dist < bestDist And dist <= radius Then
            bestDist = dist
            Set bestEl = el
        End If
NextEl:
    Loop

    Set FindNearestElement = bestEl
End Function

' ============================================================
' FindNearestTextElement
' Same as FindNearestElement but restricted to text element types.
' ============================================================
Private Function FindNearestTextElement(x As Double, y As Double, _
                                         radius As Double) As Element
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    oScan.ExcludeAllTypes
    oScan.IncludeType msdElementTypeText
    oScan.IncludeType msdElementTypeTextNode

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim bestEl As Element
    Dim bestDist As Double
    bestDist = 1E+30
    Set bestEl = Nothing

    Do While oEnum.MoveNext
        Dim el As Element
        Set el = oEnum.Current

        On Error Resume Next
        Dim lo As Point3d, hi As Point3d
        lo = el.Range.Low
        hi = el.Range.High
        If Err.Number <> 0 Then Err.Clear: GoTo NextTextEl
        On Error GoTo 0

        Dim cx As Double, cy As Double
        cx = (lo.X + hi.X) / 2
        cy = (lo.Y + hi.Y) / 2
        Dim dist As Double
        dist = Sqr((cx - x) ^ 2 + (cy - y) ^ 2)
        If dist < bestDist And dist <= radius Then
            bestDist = dist
            Set bestEl = el
        End If
NextTextEl:
    Loop

    Set FindNearestTextElement = bestEl
End Function

' ============================================================
' PromptUserClickElement
' Hides frmBBImport, shows an ECHO prompt in MicroStation,
' waits for user to click an element, returns it.
' Returns Nothing if user cancels (right-click).
' ============================================================
Private Function PromptUserClickElement(markupID As String, promptText As String) As Element
    Set PromptUserClickElement = Nothing

    frmBBImport.Hide
    CadInputQueue.SendKeyin "ECHO " & markupID & ": " & promptText & " (right-click to skip)"

    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput

    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CommandState.StartDefaultCommand
            frmBBImport.Show vbModeless
            Exit Function
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CommandState.StartDefaultCommand
    frmBBImport.Show vbModeless

    ' Find element nearest to the clicked point
    Dim clickPt As Point3d
    clickPt = oMsg.Point
    Set PromptUserClickElement = FindNearestElement(clickPt.X, clickPt.Y, 5#)
End Function

' ============================================================
' PromptUserClickPoint
' Hides frmBBImport, prompts user to click a point in MicroStation.
' Returns True and fills pt if user clicked; False if cancelled.
' ============================================================
Private Function PromptUserClickPoint(markupID As String, promptText As String, _
                                       ByRef pt As Point3d) As Boolean
    PromptUserClickPoint = False

    frmBBImport.Hide
    CadInputQueue.SendKeyin "ECHO " & markupID & ": " & promptText & " (right-click to skip)"

    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput

    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CommandState.StartDefaultCommand
            frmBBImport.Show vbModeless
            Exit Function
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    pt = oMsg.Point
    CommandState.StartDefaultCommand
    frmBBImport.Show vbModeless
    PromptUserClickPoint = True
End Function
