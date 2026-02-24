Option Explicit

' ============================================================
' SignSubColorEditor.bas
' Entry point for the Sign Attribute Editor form.
' ============================================================

Public Sub LaunchSignSubColorEditor()
    On Error GoTo Fail
    frmSignSubColors.Show vbModeless
    Exit Sub
Fail:
    MsgBox "Error launching editor: " & Err.Description, vbCritical
End Sub
