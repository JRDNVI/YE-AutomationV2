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


' "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Yields and Efficiencies 2025 - Cahir.xlsb"
' "\\Oak-w2k22-dc01\Department\Production\Production\Yields\2025\Downtime Master File 2025 (New).xlsx"
' "\\Oak-w2k22-dc01\Department\Supervisor\"
' C:\Users\coadyj\projects\YE-AutomationV2\Logs\

Private Const MASTER_FILE As String = "D:\project\OakPark-Data-YE\Master-Data\Yields and Efficiencies 2025 - Cahir.xlsb"
Private Const MASTER_DOWNTIME_FILE As String = "D:\project\OakPark-Data-YE\Master-Data\Downtime Master File 2025 (New).xlsx"
Private Const SUPERVISOR_ROOT As String = "D:\project\OakPark-Data-YE\supervisor\"
Private Const LOG_FOLDER As String = "D:\project\YE-AutomationV2\Logs" 
Private Const TARGET_DAY As Long = 9
Private Const TARGET_MONTH As Long = 10
Private Const TARGET_YEAR As Long = 2025

Private Shifts As Variant
Private Lines As Variant
Private CuringPrefixes As Variant

Private ThermoColumns As Variant
private SmokingColumns As Variant

Private LineOne As Object
Private LineTwoThree As Object
Private LineFour As Object

Private LineOneProduction As Object
Private LineTwoThreeProduction As Object
Private LineFourProduction As Object

private LineOneBreakdown As Object
private LineTwoThreeBreakdown As Object
private LineFourBreakdown As Object

Private Sub Class_Initialize()

    CuringPrefixes = Array("Smoking & Heat Treating Day By Hour Sheet", "Thermo Day By Hour Sheet.")
    Shifts = Array("Day Shift", "Night Shift")
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

    Set LineOneProduction = CreateObject("Scripting.Dictionary")
    Set LineTwoThreeProduction = CreateObject("Scripting.Dictionary")
    Set LineFourProduction = CreateObject("Scripting.Dictionary")

    With LineOneProduction
        .Add "Product Code", "C"
        .Add "Product", "D"
        .Add "Time", "G"
        .Add "Web Change", "H"
        .Add "Label Change", "I"
        .Add "Raw material", "J"
        .Add "Packaging", "K"
        .Add "Change over", "L"
        .Add "People", "M"
        .Add "Training", "N"
        .Add "Other", "P"
        .Add "Description of Issue", "S"
    End With

    With LineTwoThreeProduction
        .Add "Product Code", "C"
        .Add "Product", "D"
        .Add "Time", "G"
        .Add "Web Change", "H"
        .Add "Label Change", "I"
        .Add "Raw material", "J"
        .Add "Packaging", "K"
        .Add "Change over", "L"
        .Add "People", "M"
        .Add "Training", "N"
        .Add "Other", "P"
        .Add "Description of Issue", "S"
    End With

    With LineFourProduction
        .Add "Product Code", "C"
        .Add "Product", "D"
        .Add "Time", "G"
        .Add "Web Change", "H"
        .Add "Label Change", "I"
        .Add "Raw material", "J"
        .Add "Packaging", "K"
        .Add "Change over", "L"
        .Add "People", "M"
        .Add "Training", "N"
        .Add "Other", "P"
        .Add "Description of Issue", "S"
    End With

    set LineOneBreakdown = CreateObject("Scripting.Dictionary")
    set LineTwoThreeBreakdown = CreateObject("Scripting.Dictionary")
    set LineFourBreakdown = CreateObject("Scripting.Dictionary")

    with LineOneBreakdown
        .Add "Product Code", "C"
        .Add "total", "G"
        .Add "Product", "D"
        .Add "Machine", "R"
        .Add "Duration of Downtime", "O"
        .Add "Description of Downtime", "S"
    end with

    with LineTwoThreeBreakdown
        .Add "Product Code", "C"
        .Add "total", "G"
        .Add "Product", "D"
        .Add "Machine", "R"
        .Add "Duration of Downtime", "O"
        .Add "Description of Downtime", "S"
    end with

    with LineFourBreakdown
        .Add "Product Code", "C"
        .Add "total", "G"
        .Add "Product", "D"
        .Add "Machine", "R"
        .Add "Duration of Downtime", "O"
        .Add "Description of Downtime", "S"
    end with

    ThermoColumns = Array( _ 
        Array("C", "H"), _
        Array("G", "G"), _
        Array("H", "D") _
    )

    SmokingColumns = Array( _
        Array("B", "G"), _
        Array("C", "F"), _
        Array("E", "E"), _
        Array("A", "C"), _
        Array("I", "H"), _
        Array("J", "I") _
    )

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

Public Property Get GetLineMap(ByVal lineName As String) As Object
    Select Case LCase(lineName)
        Case "line 1": Set GetLineMap = LineOne
        Case "line 2 & 3": Set GetLineMap = LineTwoThree
        Case "line 4": Set GetLineMap = LineFour
        Case Else: Set GetLineMap = Nothing
    End Select
End Property

Public Property Get GetLineProductionMap(ByVal lineName As String) As Object
    Select Case LCase(lineName)
        Case "line 1": Set GetLineProductionMap = LineOneProduction
        Case "line 2 & 3": Set GetLineProductionMap = LineTwoThreeProduction
        Case "line 4": Set GetLineProductionMap = LineFourProduction
        Case Else: Set GetLineProductionMap = Nothing
    End Select
End Property

Public Property Get GetLineBreakdownMap(ByVal lineName As String) As Object
    Select Case LCase(lineName)
        Case "line 1": Set GetLineBreakdownMap = LineOneBreakdown
        Case "line 2 & 3": Set GetLineBreakdownMap = LineTwoThreeBreakdown
        Case "line 4": Set GetLineBreakdownMap = LineFourBreakdown
        Case Else: Set GetLineBreakdownMap = Nothing
    End Select
End Property

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
    GetTargetDate = DateSerial(TARGET_YEAR, TARGET_MONTH, TARGET_DAY)
End Property

Public Property Get GetThermoColumns() As Variant
    GetThermoColumns = ThermoColumns
End Function

Public Property Get GetSmokingColumns() As Variant
    GetSmokingColumns = SmokingColumns
End Function


 