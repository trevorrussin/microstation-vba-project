Sub BmrWZTCOther()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long
    Dim oMessage As CadInputMessage

'   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->activeCellUtf16", "Default", ""

'   Start a command
    CadInputQueue.SendCommand "PLACE CELL ICON"

'   Coordinates are in master units
    startPoint.X = -5358.11232975071
    startPoint.Y = -8742.55813011295
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZAP_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 0.00530957537012
    point.Y = startPoint.Y - 1.67393942538183E-02
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 2.65478768687899E-03
    point.Y = startPoint.Y - 2.15220783265977E-02
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a reset to the current command
    CadInputQueue.SendReset

    CadInputQueue.SendCommand "PLACE CELL ICON"

    SetCExpressionValue "tcb->activeCellUtf16", "TWZAPC_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 2.00363730680874
    point.Y = startPoint.Y - 13.9097618514224
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZAPT_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 2.00363730680874
    point.Y = startPoint.Y - 30.4051400484841
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZBCD_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 2.00363730680874
    point.Y = startPoint.Y - 44.2152241204421
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZCMS_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 2.00363730680874
    point.Y = startPoint.Y - 58.6007283620638
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZFLG_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 2.00363730680874
    point.Y = startPoint.Y - 71.1657812931062
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZFT_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.20511290845025
    point.Y = startPoint.Y - 86.9487345182006
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZIA_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.20511290845025
    point.Y = startPoint.Y - 102.238470455013
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZLUM_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 10.2905140244729
    point.Y = startPoint.Y - 114.842912266718
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZSDT_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.4720451339208
    point.Y = startPoint.Y - 125.584088767131
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZSDTD_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.4720451339208
    point.Y = startPoint.Y - 138.897041127881
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZSGN_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.4720451339208
    point.Y = startPoint.Y - 150.186236837497
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZSIG_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 9.40220660014893
    point.Y = startPoint.Y - 162.19642031922
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZWL_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 6.9929721548715
    point.Y = startPoint.Y - 172.389577610427
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZWV_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.02178033600921
    point.Y = startPoint.Y - 183.021150269
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "tcb->activeCellUtf16", "TWZWVA_P", ""

    CadInputQueue.SendCommand "PLACE CELL ICON"

    point.X = startPoint.X - 5.02178033600921
    point.Y = startPoint.Y - 196.831234340956
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CommandState.StartDefaultCommand
End Sub
