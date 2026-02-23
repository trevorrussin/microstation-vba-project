Option Explicit

' ============================================================
' AlignAddBtn — WithEvents handler for per-alignment "+ Add Row" buttons.
' One instance per alignment section. Clicking the button calls
' ParentForm.AddRowForAlign(AlignIdx).
' ============================================================

Public WithEvents Btn As MSForms.CommandButton
Public AlignIdx    As Integer
Public ParentForm  As Object

Private Sub Btn_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        ParentForm.AddRowForAlign AlignIdx
    End If
End Sub