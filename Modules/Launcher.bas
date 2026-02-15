
' ============================================================
' MODULE - Launch NYSDOT Sheet Viewer & WZTC Designer
' ============================================================

Sub LaunchNYSDOTViewer()
    ' Launch the NYSDOT 619 Standard Sheets viewer
    SheetViewer.Show
End Sub

Sub LaunchWZTC()
    ' Launch the Workzone Traffic Control Designer (modeless so sheet viewer can stay open too)
    WZTCDesigner.Show vbModeless
End Sub
