
Option Explicit

' ============================================================
' SIGN PLACEMENT MODULE
' ------------------------------------------------------------
' Manages the sign drawing step that follows alignment placement.
' After the user clicks "Next: Draw Signs" in PlacePerp,
' this module shows PlaceSign and steps through each sign
' that had a perpendicular line placed during alignment placement.
'
' State is stored in the wztcPlacedSign* public arrays
' (ModuleWZTCData.bas) plus the currentSignIdx below.
' ============================================================

Public currentSignIdx As Integer   ' 0-based index into wztcPlacedSign* arrays

' ============================================================
' ENTRY POINT - called by PlacePerp btnNext_Click
' ============================================================
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

' ============================================================
' STATE ACCESSORS (called by PlaceSign)
' ============================================================
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

' ============================================================
' DRAW THE CURRENT SIGN
' Passes the stored perpendicular line geometry to PlaceWorkZoneSign
' (ModTest.bas) which handles click collection and drawing.
' ============================================================
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
