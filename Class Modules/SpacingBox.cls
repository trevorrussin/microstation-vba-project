Option Explicit

' ============================================================
' SpacingBox — WithEvents handler for editable spacing textboxes.
' When user edits a value and presses Enter OR tabs away (Exit),
' syncs the new value to the matching row in the alignment table.
' Set AlignIdx = 0 to disable syncing (boxes with no alignment row).
' ============================================================

Public WithEvents Txt As MSForms.TextBox
Public AlignIdx   As Integer   ' 1 = Upstream, 2 = Downstream, 0 = no sync
Public RowLabel   As String    ' Label text to match in the alignment table row
Public ParentForm As Object

Private Sub Txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    SyncToAlignment
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SyncToAlignment
        KeyCode = 0
    End If
End Sub

Private Sub SyncToAlignment()
    On Error Resume Next
    If AlignIdx < 1 Or RowLabel = "" Then Exit Sub
    If Not ParentForm Is Nothing Then
        ParentForm.SyncSpacingToAlignment AlignIdx, RowLabel, Trim(Txt.Text)
    End If
End Sub
