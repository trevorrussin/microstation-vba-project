Option Explicit

' ============================================================
' AlignSectionBtns — WithEvents handler for the per-alignment
' row action footer: + Add Row, Del Row, Row# selector, Move Up, Move Down.
' One instance per alignment section.
' ============================================================

Public WithEvents BtnAdd   As MSForms.CommandButton
Public WithEvents BtnDel   As MSForms.CommandButton
Public WithEvents BtnUp    As MSForms.CommandButton
Public WithEvents BtnDown  As MSForms.CommandButton
Public WithEvents CboRowSel As MSForms.ComboBox

Public AlignIdx   As Integer
Public ParentForm As Object

Private Sub BtnAdd_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then ParentForm.AddRowForAlign AlignIdx
End Sub

Private Sub BtnDel_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then ParentForm.DelRowForAlign AlignIdx
End Sub

Private Sub BtnUp_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then ParentForm.MoveRowUpForAlign AlignIdx
End Sub

Private Sub BtnDown_Click()
    On Error Resume Next
    If Not ParentForm Is Nothing Then ParentForm.MoveRowDownForAlign AlignIdx
End Sub

Private Sub CboRowSel_Change()
    On Error Resume Next
    If Not ParentForm Is Nothing Then
        Dim rIdx As Integer: rIdx = CboRowSel.ListIndex + 1
        If rIdx >= 1 Then ParentForm.SelectAlignRowFromSel AlignIdx, rIdx
    End If
End Sub
