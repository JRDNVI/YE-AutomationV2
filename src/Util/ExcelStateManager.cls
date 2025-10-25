VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ExcelStateManager"
Option Explicit

' =============================================
' CLASS: ExcelStateManager
' Manages Excel state safely (alerts, updating, events, links)
' =============================================

Private app As Excel.Application
Private prevScreenUpdating As Boolean
Private prevDisplayAlerts As Boolean
Private prevEnableEvents As Boolean
Private prevAskToUpdateLinks As Boolean
Private stateActive As Boolean


' Initialise with current application
Public Sub Init(ByRef xlApp As Excel.Application)
    Set app = xlApp
End Sub


' Disable volatile features for automation
Public Sub DisableAll(Optional ByVal showInfo As Boolean = True)
    On Error Resume Next

    If app Is Nothing Then Set app = Application

    ' Store current state
    prevScreenUpdating = app.ScreenUpdating
    prevDisplayAlerts = app.DisplayAlerts
    prevEnableEvents = app.EnableEvents
    prevAskToUpdateLinks = app.AskToUpdateLinks

    ' Apply performance/quiet mode
    app.ScreenUpdating = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AskToUpdateLinks = False

    stateActive = True

    If showInfo Then
        Debug.Print "[INFO] Excel state disabled: alerts, events, screen updating."
    End If
End Sub



' Restore previous state
Public Sub RestoreAll(Optional ByVal showInfo As Boolean = True)
    On Error Resume Next

    If Not stateActive Then Exit Sub
    If app Is Nothing Then Set app = Application

    app.ScreenUpdating = prevScreenUpdating
    app.DisplayAlerts = prevDisplayAlerts
    app.EnableEvents = prevEnableEvents
    app.AskToUpdateLinks = prevAskToUpdateLinks

    stateActive = False

    If showInfo Then
        Debug.Print "[INFO] Excel state restored to defaults."
    End If
End Sub

' Safe cleanup (for use in Finally blocks)
Public Sub Dispose()
    On Error Resume Next
    RestoreAll False
    Set app = Nothing
End Sub
