
Option Explicit

' ============================================================
' WZTC DRAWING ELEMENTS MODULE
' ------------------------------------------------------------
' Manages the interactive drawing of WZTC linear/shape elements
' after the sign placement step.
'
' Elements (in order):
'   1. Work Space                       - PLACE SHAPE CONSTRAINED (closed polygon)
'   2. Channelizing Devices             - level TWZCD_P,   PLACE LINE CONSTRAINED
'   3. Removal Striping                 - level TWZPMRC_P, PLACE LINE CONSTRAINED
'   4. Temporary Barrier                - level TWZBT_P,   PLACE LINE CONSTRAINED
'   5. Temp. Barrier w/Warning Lights   - level TWZBTWL_P, PLACE LINE CONSTRAINED
'
' The user can click "Start Drawing" multiple times to add multiple
' segments/runs for the same element type before advancing.
' ============================================================

Private Const ELEMENT_COUNT As Integer = 5

Public currentElementIdx As Integer   ' 1-based index

' ============================================================
' ELEMENT METADATA
' ============================================================
Private Function GetElementName(idx As Integer) As String
    Select Case idx
        Case 1: GetElementName = "Work Space"
        Case 2: GetElementName = "Channelizing Devices"
        Case 3: GetElementName = "Removal Striping"
        Case 4: GetElementName = "Temporary Barrier"
        Case 5: GetElementName = "Temp. Barrier w/Warning Lights"
        Case Else: GetElementName = ""
    End Select
End Function

Private Function GetElementLevel(idx As Integer) As String
    Select Case idx
        Case 1: GetElementLevel = "TWZWS2_P"
        Case 2: GetElementLevel = "TWZCD_P"
        Case 3: GetElementLevel = "TWZPMRC_P"
        Case 4: GetElementLevel = "TWZBT_P"
        Case 5: GetElementLevel = "TWZBTWL_P"
        Case Else: GetElementLevel = ""
    End Select
End Function

Private Function IsWorkSpace(idx As Integer) As Boolean
    IsWorkSpace = (idx = 1)
End Function

' ============================================================
' ENTRY POINT - called by PlaceSign btnWZTCElements_Click
' ============================================================
Public Sub StartWZTCElementsPlacement()
    currentElementIdx = 1
    PlaceElements.Show vbModeless
End Sub

' ============================================================
' STATE ACCESSORS (called by PlaceElements)
' ============================================================
Public Function GetCurrentElementDisplayName() As String
    GetCurrentElementDisplayName = GetElementName(currentElementIdx)
End Function

Public Function GetCurrentElementNumber() As Integer
    GetCurrentElementNumber = currentElementIdx
End Function

Public Function GetTotalElementCount() As Integer
    GetTotalElementCount = ELEMENT_COUNT
End Function

Public Function IsAllElementsDone() As Boolean
    IsAllElementsDone = (currentElementIdx > ELEMENT_COUNT)
End Function

Public Sub AdvanceElement()
    currentElementIdx = currentElementIdx + 1
End Sub

Public Function GetCurrentElementInstructions() As String
    If IsWorkSpace(currentElementIdx) Then
        GetCurrentElementInstructions = _
            "Click points to trace the work space boundary, then right-click to close the shape." & vbCrLf & _
            "After the shape closes, click inside the shape border to place the hatch pattern."
    Else
        GetCurrentElementInstructions = _
            "Click points to draw the " & GetElementName(currentElementIdx) & _
            " line, then right-click when done. Click 'Start Drawing' again to add another run."
    End If
End Function

' ============================================================
' DRAW ONE SEGMENT/RUN FOR THE CURRENT ELEMENT
' Called when the user clicks "Start Drawing".
' Sets the level, starts the appropriate command, then routes
' user clicks via GetInput until the user right-clicks (Reset).
' ============================================================
Public Sub DrawCurrentElementSegment()
    If currentElementIdx < 1 Or currentElementIdx > ELEMENT_COUNT Then Exit Sub

    Dim lvl As String
    lvl = GetElementLevel(currentElementIdx)

    ' Set level, color 6 and weight 2 for all WZTC elements
    If lvl <> "" Then
        CadInputQueue.SendCommand "ACTIVE LEVEL """ & lvl & """"
    End If
    CadInputQueue.SendCommand "ACTIVE COLOR 6"
    CadInputQueue.SendCommand "ACTIVE WEIGHT 2"

    If IsWorkSpace(currentElementIdx) Then
        ' ---- Work Space: draw a closed shape ----
        CadInputQueue.SendCommand "PLACE SHAPE CONSTRAINED"
    Else
        ' ---- Linear element ----
        CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"
    End If

    ' Route user clicks to the active command until right-click (Reset).
    Dim oMsg As CadInputMessage
    Dim nPts As Integer
    nPts = 0
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeReset
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
            nPts = nPts + 1
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    ' ---- WorkSpace: user-click hatch (matches LegacyElements.bas pattern) ----
    ' After the shape closes, start HATCH ICON and let the user click inside
    ' the shape boundary to place the hatch. This avoids centroid issues with
    ' non-convex shapes and gives the user control over hatch placement.
    If IsWorkSpace(currentElementIdx) And nPts >= 3 Then
        CadInputQueue.SendCommand "HATCH ICON"
        CadInputQueue.SendKeyin "ECHO Click inside the work space shape to place the hatch pattern."

        ' Wait for user to click inside the shape (data point)
        Set oMsg = CadInputQueue.GetInput
        Do While oMsg.InputType <> msdCadInputTypeDataPoint
            If oMsg.InputType = msdCadInputTypeReset Then
                ' User cancelled hatch
                CadInputQueue.SendReset
                CommandState.StartDefaultCommand
                GoTo HatchDone
            End If
            Set oMsg = CadInputQueue.GetInput
        Loop

        ' Send the click point to HATCH ICON (first point identifies the shape)
        CadInputQueue.SendDataPoint oMsg.Point, 1
        ' Send same point again to confirm (matches Legacy pattern)
        CadInputQueue.SendDataPoint oMsg.Point, 1
        CadInputQueue.SendReset
        CommandState.StartDefaultCommand
    End If
HatchDone:
End Sub
