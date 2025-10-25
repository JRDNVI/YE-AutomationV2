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
' the ReadYESheet helper.
' ===========================================

' Dependencies 
Private config As ConfigCahir
Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
Private yeHelper As ReadYESheet  


Private masterPath As String
Private supervisorRoot As String
Private targetDate As Date


' Initialise Dependencies
Public Sub Init(ByRef cfg As ConfigCahir, _
                ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)
    Set config = cfg
    Set logger = lg
    Set errHandler = eh
    Set state = st

    masterPath = config.GetPath("master_file")
    supervisorRoot = config.GetPath("supervisor_root")
    targetDate = config.GetTargetDate

    Set yeHelper = New ReadYESheet
    yeHelper.Init logger, errHandler, state
End Sub

Public Sub Run()
    On Error GoTo ERR_HANDLER

    logger.Info "Slicing Import Started..."

    Dim allRows As New Collection     ' Collects everything across all lines
    Dim lineName As Variant
    Dim shift As Variant
    Dim filePath As String
    Dim colMap As Object
    Dim rows As Collection
    Dim r As Variant

    ' --- Process all lines and shifts ---
    For Each lineName In config.GetLines
        logger.Info "Processing " & lineName

        Set colMap = config.GetLineMap(lineName)
        If colMap Is Nothing Then
            logger.Warn "No column mapping found for " & lineName
            GoTo NEXT_LINE
        End If

        For Each shift In config.GetShifts
            filePath = BuildSupervisorPath(lineName, shift)
            logger.Info "Checking file: " & filePath

            If Dir(filePath) = "" Then
                logger.Warn "File not found: " & filePath
                GoTo CONTINUE_SHIFT
            End If

            ' --- Read data for this shift ---
            Set rows = yeHelper.ReadSupervisorData(filePath, lineName, shift, targetDate, colMap)

            If Not rows Is Nothing And rows.Count > 0 Then
                logger.Info "Collected " & rows.Count & " rows for " & lineName & " - " & shift

                ' Append all shift data to main list
                For Each r In rows
                    allRows.Add r
                Next r
            Else
                logger.Warn lineName & " - " & shift & ": no valid data found."
            End If

CONTINUE_SHIFT:
        Next shift

NEXT_LINE:
    Next lineName

    ' --- One single write at the end ---
    If allRows.Count > 0 Then
        yeHelper.WriteToMaster allRows, masterPath, config.GetLineMap("Line 1")
        logger.Success "Slicing Import Completed: " & allRows.Count & " total rows written."
    Else
        logger.Warn "No data collected from any line or shift."
    End If

    Exit Sub

ERR_HANDLER:
    errHandler.HandleError "SlicingImportController", "Run", Erl
    On Error Resume Next
    Resume Next
End Sub




Private Function BuildSupervisorPath(ByVal lineName As String, _
                                     ByVal shift As String) As String
    Dim yearStr As String
    Dim monthStr As String
    Dim folderPath As String
    Dim fileName As String

    yearStr = Format(Year(targetDate), "0000")
    monthStr = Format(targetDate, "mmmm")

    ' Folder structure: <root>\<shift>\<line>\<year>\<month>\
    folderPath = supervisorRoot & "\" & shift & "\" & lineName & "\" & yearStr & "\" & monthStr & "\"

    fileName = Format(targetDate, "dd.mm.yyyy") & ".xlsx"

    BuildSupervisorPath = folderPath & fileName
End Function
