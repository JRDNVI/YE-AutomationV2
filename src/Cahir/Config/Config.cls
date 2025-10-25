VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ConfigCahir"
Option Explicit

' ==============================================
' CLASS: ConfigCahir
' Configuration for Cahir automation
' ==============================================

Private Const MASTER_FILE As String = "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Yields and Efficiencies 2025 - Cahir.xlsb"
Private Const MASTER_DOWNTIME_FILE As String = "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Downtime Master File 2025 (New).xlsx"
Private Const SUPERVISOR_ROOT As String = "\\Oak-w2k22-dc01\Department\Supervisor\"
Private Const LOG_FOLDER As String = "C:\Users\coadyj\projects\YE-AutomationV2\Logs\"
Private Const TARGET_DATE As Date = #22/10/2025#

Private Shifts As Variant
Private Lines As Variant
Private CuringPrefixes As Variant

Private LineOne As Object
Private LineTwoThree As Object
Private LineFour As Object


Private Sub Class_Initialize()

    CuringPrefixes = Array("Smoking & Heat Treating Day By Hour Sheet", "Thermo Day By Hour Sheet.")
    Shifts = Array("Day Shift Yield & Efficiency", "Night Shift Yield & Efficiency")
    Lines = Array("Line 1", "Line 2 & 3", "Line 4")

    Set LineOne = CreateObject("Scripting.Dictionary")
    Set LineTwoThree = CreateObject("Scripting.Dictionary")
    Set LineFour = CreateObject("Scripting.Dictionary")

    With LineOne
        .Add "Product Code", "C"
        .Add "Total Min.Work", "N"
        .Add "Pack Produced", "Q"
        .Add "Nominal weight", "M"
        .Add "Weight to slicer", "I"
        .Add "Slicer Weight", "L"
        .Add "Over Runs Used KG", "J"
        .Add "Gripper End KGS", "T"
        .Add "Trim KGS", "U"
        .Add "Dropped Joint KGS", "S"
        .Add "Downtime", "O"
    End With

    With LineTwoThree
        .Add "Product Code", "C"
        .Add "Total Min.Work", "N"
        .Add "Pack Produced", "Q"
        .Add "Nominal weight", "M"
        .Add "Weight to slicer", "I"
        .Add "Slicer Weight", "L"
        .Add "Over Runs Used KG", "J"
        .Add "Gripper End KGS", "T"
        .Add "Trim KGS", "U"
        .Add "Dropped Joint KGS", "S"
        .Add "Downtime", "O"
    End With

    With LineFour
        .Add "Product Code", "C"
        .Add "Total Min.Work", "N"
        .Add "Pack Produced", "Q"
        .Add "Dropped Joint KGS", "S"
        .Add "Nominal weight", "M"
        .Add "Weight to slicer", "I"
        .Add "Gripper End KGS", "T"
        .Add "Trim KGS", "U"
        .Add "Downtime", "O"
    End With

End Sub

Public Function GetPath(ByVal key As String) As String
    Select Case LCase(key)
        Case "master_file": GetPath = MASTER_FILE
        Case "supervisor_root": GetPath = SUPERVISOR_ROOT
        Case "master_downtime_file": GetPath = MASTER_DOWNTIME_FILE
        Case "log_folder": GetPath = LOG_FOLDER
        Case Else: GetPath = ""
    End Select
End Function

Public Property Get GetShifts() As Variant
    GetShifts = Shifts
End Property

Public Property Get GetLines() As Variant
    GetLines = Lines
End Property

Public Property Get GetCuringPrefixes() As Variant
    GetCuringPrefixes = CuringPrefixes
End Property

Public Property Get GetTargetDate() As Date
    GetTargetDate = TARGET_DATE
End Property

Public Property Get GetLineMap(ByVal lineName As String) As Object
    Select Case LCase(lineName)
        Case "line 1": Set GetLineMap = LineOne
        Case "line 2 & 3": Set GetLineMap = LineTwoThree
        Case "line 4": Set GetLineMap = LineFour
        Case Else: Set GetLineMap = Nothing
    End Select
End Property
