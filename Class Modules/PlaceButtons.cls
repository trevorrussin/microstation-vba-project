Attribute VB_Name = "PlacementButtonHandler"
Option Explicit

' ============================================================
' PLACEMENT BUTTON HANDLER
' Sinks Click events for the dynamically-created CommandButtons
' on frmAlignmentPlacement and dispatches to the form's public
' OnPlace / OnSkip / OnCancel methods.
'
' Usage (in frmAlignmentPlacement.UserForm_Initialize):
'   Dim h As New PlacementButtonHandler
'   h.Action = "Place"        ' or "Skip" / "Cancel"
'   Set h.Btn = <CommandButton reference>
'   Set h.HostForm = Me
'   btnHandlers.Add h          ' keep reference alive
' ============================================================

Public WithEvents Btn As MSForms.CommandButton
Public Action   As String    ' "Place" | "Skip" | "Cancel"
Public HostForm As Object    ' frmAlignmentPlacement instance

Private Sub Btn_Click()
    Select Case Action
        Case "Place":  HostForm.OnPlace
        Case "Skip":   HostForm.OnSkip
        Case "Cancel": HostForm.OnCancel
    End Select
End Sub
