VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "SummaryImportController"
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

Private masterPath As String
Private downtimePath As String
Private targetDate As Date

Public Sub Init(ByRef cfg As ConfigCahir, _
                ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)

    Set config = cfg
    Set logger = lg
    Set errHandler = eh
    Set state = st

End Sub

Public Sub Run()
    On Error GoTo ERR_HANDLER

    logger.Info "Updating Pivot Tables..."

    Dim targetDate As Date: targetDate = config.GetTargetDate
    Dim masterPath As String: masterPath = config.GetPath("master_file")
    Dim downtimePath As String: downtimePath = config.GetPath("master_downtime_file")

    Dim pivotUpdater As New UpdatePivotTables
    Dim pivots As Object: Set pivots = config.GetPivotTableName

    pivotUpdater.UpdateAllPivotTables masterPath, downtimePath, pivots, targetDate, logger

    Exit Sub

ERR_HANDLER:
    errHandler.HandleError "SlicingImportController", "Run", Erl
End Sub
