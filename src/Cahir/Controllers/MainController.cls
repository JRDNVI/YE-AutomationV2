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
'Private curingCtrl As CuringImportController

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
    logger.Info "Target Date: " & Format(config.GetTargetDate, "dd/mm/yyyy")
    errHandler.Init logger, False

    state.Init Application
    state.DisableAll
    logger.Info "Excel state set to automation mode."

    Set slicingCtrl = New SlicingImportController
    slicingCtrl.Init config, logger, errHandler, state
    logger.Info "Running Slicing Import Controller..."
    slicingCtrl.Run
    logger.Success "Slicing import complete."

    ' --- Run Curing Controller 
    'Set curingCtrl = New CuringImportController
    'curingCtrl.Init config, logger, errHandler, state
    'logger.Info "Running Curing Import Controller..."
    'curingCtrl.Run
    'logger.Success "Curing import complete."

    state.RestoreAll
    logger.Info "Excel state restored to normal."
    logger.Success "All imports completed successfully."
    logger.Info "Total runtime: " & Format(Now - runStart, "hh:nn:ss")
    logger.EndSession

    Exit Sub

' Global error Catch
ERR_HANDLER:
    errHandler.HandleError "MainControllerCahir", "RunImport", Erl

    On Error Resume Next
    state.RestoreAll
    logger.Error "Unexpected error. Excel state restored."
    logger.EndSession
End Sub
