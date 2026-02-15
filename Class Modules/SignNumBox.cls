
Option Explicit

' Sinks Exit and KeyDown for dynamically created sign number textboxes.
' When the user leaves the field or presses Enter, spacing and size auto-fill from the sign library.

Public WithEvents Txt As MSForms.TextBox
Public RowIndex As Integer
Public ParentForm As Object

Private Sub Txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.ApplySignLibraryToRow RowIndex
    End If
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Enter key: auto-fill spacing and size, then move focus to spacing column
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        If Not ParentForm Is Nothing Then
            ParentForm.ApplySignLibraryToRow RowIndex
            ParentForm.MoveFocusToSpacing RowIndex
        End If
        KeyCode = 0   ' Consume Enter so it doesn't beep or add newline
    End If
End Sub
