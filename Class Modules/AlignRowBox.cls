
Option Explicit

' Sinks GotFocus, Exit, and KeyDown for dynamically created Label/Sign# textboxes
' in the combined alignment table.
' GotFocus  → selects the row (highlights it, sets activeAlignIdx/activeRowIdx)
' Exit      → triggers sign library auto-fill if row type is "Sign"
' KeyDown   → Enter key auto-fills and advances focus

Public WithEvents Txt As MSForms.TextBox
Public AlignIdx As Integer
Public RowIdx As Integer
Public ParentForm As Object

Private Sub Txt_GotFocus()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.SelectAlignRow AlignIdx, RowIdx
    End If
End Sub

Private Sub Txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.ApplySignLibraryToAlignRow AlignIdx, RowIdx
    End If
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        On Error Resume Next
        If Not ParentForm Is Nothing Then
            ParentForm.ApplySignLibraryToAlignRow AlignIdx, RowIdx
        End If
        KeyCode = 0   ' Consume Enter so it doesn't beep
    End If
End Sub
