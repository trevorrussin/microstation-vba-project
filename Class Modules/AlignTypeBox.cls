
Option Explicit

' Sinks the Change event for dynamically created Type ComboBoxes
' in the combined alignment table.
' Change → enables/disables Size and Side columns based on "Sign" vs "Non-Sign"

Public WithEvents Cbo As MSForms.ComboBox
Public AlignIdx As Integer
Public RowIdx As Integer
Public ParentForm As Object

Private Sub Cbo_Change()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.OnTypeChange AlignIdx, RowIdx
    End If
End Sub
