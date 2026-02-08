Attribute VB_Name = "Module1"
' ============================================================
' MODULE - Launch NYSDOT Sheet Viewer & WZTC Designer
' ============================================================

Sub LaunchNYSDOTViewer()
    ' Launch the NYSDOT 619 Standard Sheets viewer
    UserForm1.Show
End Sub

Sub LaunchWZTC()
    ' Launch the Workzone Traffic Control Designer
    Call LaunchWZTCDesigner
End Sub
