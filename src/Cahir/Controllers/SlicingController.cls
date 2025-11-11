VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "SlicingImportController"
Option Explicit

' ===========================================
' CLASS: SlicingImportController
' ===========================================
' Handles import of all slicing line data
' (Line 1, Line 2 & 3, Line 4) into the
' Master Yield & Efficiency workbook using
' the ReadYESheet and ReadYEDowntime helpers.
' ===========================================

Private config As ConfigCahir
Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
Private yeHelper As ReadYESheet
Private downtimeHelper As ReadYEDowntime

Private masterPath As String
Private downtimePath As String
Private supervisorRoot As String
Private targetDate As Date

Public Sub Init(ByRef cfg As ConfigCahir, _
                ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)

    Set config = cfg
    Set logger = lg
    Set errHandler = eh
    Set state = st

    masterPath = config.GetPath("master_file")
    downtimePath = config.GetPath("master_downtime_file")
    supervisorRoot = config.GetPath("supervisor_root")
    
    targetDate = config.GetTargetDate

    Set yeHelper = New ReadYESheet
    Set downtimeHelper = New ReadYEDowntime

    yeHelper.Init logger, errHandler, state
    downtimeHelper.Init logger, errHandler, state
End Sub


Public Sub Run()
    On Error GoTo ERR_HANDLER
    Dim startTime As Single: startTime = Timer

    logger.Info "Slicing Import Started..."

    Dim allYERows As New Collection
    Dim allDowntimeRows As New Collection

    Dim lineName As Variant, shift As Variant
    Dim filePath As String
    Dim colMap As Object, downtimeMap As Object, breakdownMap As Object
    Dim rows As Collection, downtimeRows As Collection
    Dim row As Variant, downtimeRow As Variant

    For Each lineName In config.GetLines
        logger.Info "Processing " & lineName

        Set colMap = config.GetLineMap(lineName)
        Set downtimeMap = config.GetLineProductionMap(lineName)
        set breakdownMap = config.GetLineBreakdownMap(lineName)

        If colMap Is Nothing Then
            logger.Warn "No column mapping found for " & lineName
            GoTo NEXT_LINE
        End If

        For Each shift In config.GetShifts
            filePath = BuildSupervisorPath(lineName, shift)
            logger.Info "Checking file: " & filePath

            If Dir(filePath, vbNormal) = "" Then
                logger.Warn "File not found: " & filePath
                GoTo CONTINUE_SHIFT
            End If

            Set rows = yeHelper.ReadSupervisorData(filePath, lineName, shift, targetDate, colMap)
            Set downtimeRows = downtimeHelper.ReadDowntime(filePath, lineName, shift, targetDate, downtimeMap, breakdownMap)

            If Not rows Is Nothing And rows.Count > 0 Then
                logger.Success "Collected " & rows.Count & " yield rows from " & Mid(filePath, InStrRev(filePath, "\") + 1)

                Dim r As Variant
                For Each r In rows
                    allYERows.Add r
                Next r
            End If

            If Not downtimeRows Is Nothing And downtimeRows.Count > 0 Then
                Dim dr As Variant
                For Each dr In downtimeRows
                    allDowntimeRows.Add dr
                Next dr
                logger.Success "Collected " & downtimeRows.Count & " downtime rows."
            End If

CONTINUE_SHIFT:
        Next shift

NEXT_LINE:
    Next lineName

    ' Write yield data
    'If allYERows.Count > 0 Then
        'yeHelper.WriteToMaster allYERows, masterPath, config.GetLineMap("Line 1")
        'logger.Success "Wrote " & allYERows.Count & " yield rows to master."
    'Else
        'logger.Warn "No yield data collected from any line or shift."
    'End If

    ' Write downtime data
    If allDowntimeRows.Count > 0 Then
        downtimeHelper.WriteToMaster allDowntimeRows, downtimePath, config.GetLineProductionMap("Line 1"), config.GetLineBreakdownMap("Line 1")
        logger.Success "Wrote " & allDowntimeRows.Count & " downtime rows to master."
    Else
        logger.Warn "No downtime data collected from any line or shift."
    End If

    logger.Success "Slicing Import Completed: " & allYERows.Count & " total yield rows, " & allDowntimeRows.Count & " downtime rows."
    Exit Sub


ERR_HANDLER:
    errHandler.HandleError "SlicingImportController", "Run", Erl
    On Error Resume Next
    Resume Next
End Sub


Public Function BuildSupervisorPath(ByVal lineName As String, ByVal shift As String) As String
    Dim yearStr As String, monthStr As String
    Dim folderPath As String, fileName As String

    yearStr = Format(Year(targetDate), "0000")
    monthStr = Format(targetDate, "mmmm")

    ' Folder structure: <root>\<shift>\<line>\<year>\<month>\
    folderPath = supervisorRoot & shift & "\" & lineName & "\" & yearStr & "\" & monthStr & "\"
    fileName = Format(targetDate, "dd.mm.yyyy") & ".xlsx"

    BuildSupervisorPath = folderPath & fileName
End Function
