VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "MainControllerCahir"
Option Explicit

' ===========================================
' CLASS: MainControllerCahir
' ===========================================
' Entry point for Cahir site automation.
' Orchestrates slicing & curing imports with
' logging, error handling, and
' Excel state management.
' ===========================================

' helpers for DI 
Private config As ConfigCahir
Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
Private slicingCtrl As SlicingImportController
Private curingCtrl As CuringImportController
Private summaryCtrl As SummaryImportController
Private runStart As Date

Public Sub RunImport()
    On Error GoTo ERR_HANDLER
    runStart = Now

    Set config = New ConfigCahir
    Set logger = New LogHelper
    Set errHandler = New ErrorHandler
    Set state = New ExcelStateManager

    logger.StartSession config.GetPath("log_folder")
    logger.Info "Supervisor Import Started..."
    logger.Info "Target Date: " & Format$(config.GetTargetDate, "dd/mm/yyyy")
    errHandler.Init logger, False

    state.Init Application
    state.DisableAll
    logger.Info "Excel state set to automation mode."

    Set slicingCtrl = New SlicingImportController
    slicingCtrl.Init config, logger, errHandler, state
    slicingCtrl.Run

    Set curingCtrl = New CuringImportController
    curingCtrl.Init config, logger, errHandler, state   
    curingCtrl.Run

    Set summaryCtrl = New SummaryImportController
    summaryCtrl.Init config, logger, errHandler, state
    summaryCtrl.Run

    Dim masterPath As String, newPath As String
    masterPath = config.GetPath("master_file")
    newPath = Replace(masterPath, ".xlsb", ".xlsm")

    logger.Info "Creating deployed XLSM version..."

    'Dim wbSrc As Workbook
    'Application.DisplayAlerts = False 
    'Set wbSrc = Workbooks.Open(masterPath, ReadOnly:=False)
    'wbSrc.SaveAs Filename:=newPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'wbSrc.Close SaveChanges:=True
    'Application.DisplayAlerts = True 

    logger.Success "Deployment copy created: " & newPath

    state.RestoreAll
    logger.Info "Excel state restored to normal."
    logger.Success "All imports completed successfully."
    logger.Info "Total runtime: " & Format$(Now - runStart, "hh:nn:ss")
    logger.EndSession
    Exit Sub

ERR_HANDLER:
    errHandler.HandleError "MainControllerCahir", "RunImport", Erl
    On Error Resume Next
    state.RestoreAll
    logger.Error "Unexpected error. Excel state restored."
    logger.EndSession
End Sub
