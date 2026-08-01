
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

Sub LaunchChatPanel()
    ' Launch the WZTC Agent Chat panel. Run this macro -- do not press F5
    ' on WZTCChatPanel's designer/code window in the VBA IDE. F5 on a
    ' UserForm runs an implicit Show with no argument, which defaults to
    ' vbModal and blocks clicking anywhere else (MicroStation, other
    ' forms) until it's closed. This explicit vbModeless call is what
    ' every other form in this project is launched through.
    WZTCChatPanel.Show vbModeless
End Sub
