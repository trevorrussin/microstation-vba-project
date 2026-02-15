Attribute VB_Name = "ModuleWZTCCells"
Option Explicit

' ============================================================
' WZTC CELL LIBRARY MODULE
' ------------------------------------------------------------
' Manages interactive placement of WZTC cell library symbols
' from ny_plan_wztc.cel after the WZTC elements drawing step.
'
' The user selects a cell from a dropdown and clicks
' "Place Cell". The form hides, MicroStation enters PLACE CELL
' ICON mode, and the user clicks to place one or more instances
' of the selected cell. Right-clicking ends placement and
' returns to the form.
' ============================================================

Private Const WZTC_CELL_LIB As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"

' ============================================================
' ENTRY POINT - called by frmWZTCElements btnGoCellLib_Click
' ============================================================
Public Sub StartWZTCCellPlacement()
    frmWZTCCells.Show vbModeless
End Sub

' ============================================================
' CELL CATALOGUE
' Returns a 1-based array of "CELL_NAME - Description" strings
' for the ComboBox.
' ============================================================
Public Function GetCellCatalogue() As String()
    Dim cats(1 To 16) As String
    cats(1)  = "TWZAP_P - Arrow Panel"
    cats(2)  = "TWZAPC_P - Arrow Panel (Closed)"
    cats(3)  = "TWZAPT_P - Arrow Panel (Type)"
    cats(4)  = "TWZBCD_P - Barricade"
    cats(5)  = "TWZCMS_P - Changeable Message Sign"
    cats(6)  = "TWZFLG_P - Flagger"
    cats(7)  = "TWZFT_P - Flagger in Traffic"
    cats(8)  = "TWZIA_P - Impact Attenuator"
    cats(9)  = "TWZLUM_P - Luminaire"
    cats(10) = "TWZSDT_P - Sequential Directional Taper"
    cats(11) = "TWZSDTD_P - Seq. Dir. Taper (Dynamic)"
    cats(12) = "TWZSGN_P - Sign Post"
    cats(13) = "TWZSIG_P - Signal"
    cats(14) = "TWZWL_P - Warning Lights"
    cats(15) = "TWZWV_P - Warning Vehicle"
    cats(16) = "TWZWVA_P - Warning Vehicle w/Attenuator"
    GetCellCatalogue = cats
End Function

' ============================================================
' EXTRACT CELL NAME from "CELL_NAME - Description" string
' ============================================================
Public Function ExtractCellName(catalogueEntry As String) As String
    Dim pos As Integer
    pos = InStr(catalogueEntry, " - ")
    If pos > 0 Then
        ExtractCellName = Left(catalogueEntry, pos - 1)
    Else
        ExtractCellName = Trim(catalogueEntry)
    End If
End Function

' ============================================================
' PLACE THE SELECTED CELL INTERACTIVELY
' Attaches the WZTC cell library, sets the active cell,
' enters PLACE CELL ICON mode, and routes user clicks until
' the user right-clicks (Reset).
' ============================================================
Public Sub PlaceSelectedCell(cellName As String)
    If Trim(cellName) = "" Then Exit Sub

    ' Attach cell library and set active cell
    CadInputQueue.SendCommand "ATTACH LIBRARY " & WZTC_CELL_LIB
    SetCExpressionValue "tcb->activeCellUtf16", cellName, ""
    CadInputQueue.SendCommand "PLACE CELL ICON"

    ' Route user clicks until right-click (Reset)
    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeReset
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
End Sub
