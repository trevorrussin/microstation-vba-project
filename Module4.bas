Attribute VB_Name = "Module4"
' ============================================================
' SIGN DATA TYPES MODULE
' This module must be a STANDARD MODULE (not a class module)
' Contains the public type definition for SignData
' ============================================================

Option Explicit

' Sign data structure - MUST be in a standard module to be public
Public Type signData
    SignNumber As String        ' e.g., "R02-10sNY", "NYR09-11"
    Description As String       ' e.g., "Road Work Ahead"
    CellName As String         ' Cell name in library: e.g., "R02-10sNY"
    CellLibraryPath As String  ' Full path to .cel file
    TextLabel As String        ' Associated text label (e.g., "NYR09-11")
    TextLine2 As String        ' Second line of text if needed (e.g., "48"" x 48""")
    WidthInches As Double      ' Sign width in inches
    HeightInches As Double     ' Sign height in inches
    PostType As String         ' e.g., "TWZSGN_P"
    PostLibraryPath As String  ' Path to post cell library
    DefaultSpacing As Double   ' Default spacing in feet
End Type

