Sub BmrWorkZoneElements()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long
    Dim oMessage As CadInputMessage

'   Start a command
    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZWS2_P"""
    CadInputQueue.SendCommand "PLACE SHAPE CONSTRAINED"

'   Coordinates are in master units
    startPoint.X = -161.526617901037
    startPoint.Y = -11053.8809769682
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X
    point.Y = startPoint.Y + 64.0274817954232
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 117.968643722089
    point.Y = startPoint.Y + 64.0274817954232
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 117.968643722089
    point.Y = startPoint.Y - 1.17481617973135
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 0.219021313206923
    point.Y = startPoint.Y - 1.17481617973135
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "HATCH ICON"

    point.X = startPoint.X + 44.3175938442207
    point.Y = startPoint.Y - 0.217495123650224
    point.Z = startPoint.Z - 4.805624485E-11
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 44.3175938442207
    point.Y = startPoint.Y - 0.217495123650224
    point.Z = startPoint.Z - 4.805624485E-11
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZBT_P"""

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZCD_P"""

    point.X = startPoint.X - 45.6471216595473
    point.Y = startPoint.Y + 90.6119792619465
    point.Z = startPoint.Z - 4.824250937E-11
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "PLACE LINE CONSTRAINED"

    point.X = startPoint.X - 42.1256696081417
    point.Y = startPoint.Y + 90.6119792619465
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 157.746678629294
    point.Y = startPoint.Y - 14.534068824114
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 192.96119914335
    point.Y = startPoint.Y - 116.155668370977
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 104.337989182975
    point.Y = startPoint.Y - 60.3518998336822
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a reset to the current command
    CadInputQueue.SendReset

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZPMRC_P"""

    point.X = startPoint.X - 85.8996775411165
    point.Y = startPoint.Y + 105.683958759186
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 189.195604382349
    point.Y = startPoint.Y + 21.6846019083114
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 217.954129468828
    point.Y = startPoint.Y - 50.5665931452386
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZBT_P"""

    point.X = startPoint.X - 116.418928653299
    point.Y = startPoint.Y + 127.418058084239
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 219.714855494531
    point.Y = startPoint.Y + 40.4816607840312
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 240.843567802965
    point.Y = startPoint.Y - 28.2450857303211
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "ACTIVE LEVEL ""TWZBTWL_P"""

    point.X = startPoint.X - 140.482184337904
    point.Y = startPoint.Y + 143.278076510627
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 232.974200463599
    point.Y = startPoint.Y + 69.1687538258684
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 260.37924227862
    point.Y = startPoint.Y - 24.6247338295107
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CommandState.StartDefaultCommand
End Sub