Attribute VB_Name = "ModuleWZTC"
' ============================================================
' WORKZONE TRAFFIC CONTROL (WZTC) - MAIN LAUNCHER MODULE
' Version 1.0
' ============================================================
' This module provides the entry point for the WZTC application
' and coordinates between the user interface and drawing functions.
' ============================================================

Option Explicit

' ============================================================
' LAUNCH WZTC USER FORM
' Main entry point - call this sub to start the application
' ============================================================
Public Sub LaunchWZTCDesigner()
    ' Show the main WZTC user form
    WZTCUserForm.Show
End Sub

' ============================================================
' GET SIGN COUNT (stub for library verification)
' ============================================================
Public Function GetSignCount() As Integer
    ' This would return count from Module3 sign library
    ' For now, check if library is initialized
    GetSignCount = 0  ' Will be overridden by actual library module
End Function

' ============================================================
' HELPER: CALCULATE TAPER LENGTH
' Uses MUTCD NY formula: 100 + (speed - 20) * 1.5
' ============================================================
Public Function CalculateTaperLength(speedMPH As Integer) As Double
    CalculateTaperLength = 100 + (speedMPH - 20) * 1.5
End Function

' ============================================================
' HELPER: CALCULATE BUFFER SPACE
' Uses MUTCD NY formula: speed * 1.0
' ============================================================
Public Function CalculateBufferSpace(speedMPH As Integer) As Double
    CalculateBufferSpace = speedMPH * 1.0
End Function

' ============================================================
' HELPER: CALCULATE VEHICLE SPACE
' Uses MUTCD NY formula: speed * 1.5
' ============================================================
Public Function CalculateVehicleSpace(speedMPH As Integer) As Double
    CalculateVehicleSpace = speedMPH * 1.5
End Function

' ============================================================
' HELPER: CALCULATE MERGING TAPER
' Uses MUTCD NY formula: speed * 2.5
' ============================================================
Public Function CalculateMergingTaper(speedMPH As Integer) As Double
    CalculateMergingTaper = speedMPH * 2.5
End Function

' ============================================================
' HELPER: CALCULATE SHIFTING TAPERS
' Uses MUTCD NY formula: speed * 1.2
' ============================================================
Public Function CalculateShiftingTapers(speedMPH As Integer) As Double
    CalculateShiftingTapers = speedMPH * 1.2
End Function

' ============================================================
' HELPER: CALCULATE SHOULDER TAPERS
' Uses MUTCD NY formula: speed * 0.8
' ============================================================
Public Function CalculateShoulderTapers(speedMPH As Integer) As Double
    CalculateShoulderTapers = speedMPH * 0.8
End Function

' ============================================================
' HELPER: CALCULATE ADVANCED WARNING SPACING
' Uses MUTCD NY formula: speed * 10
' ============================================================
Public Function CalculateAdvancedWarningSpacing(speedMPH As Integer) As Double
    CalculateAdvancedWarningSpacing = speedMPH * 10
End Function
