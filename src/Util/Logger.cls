VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "LogHelper"

Option Explicit

' =============================================
' CLASS: LogHelper
' Centralised logging to file and terminal
' =============================================

Private logFilePath As String
Private logSessionActive As Boolean

Public Sub StartSession(ByVal logFolder As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(logFolder) Then fso.CreateFolder logFolder

    Dim ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")
    logFilePath = fso.BuildPath(logFolder, "ImportLog_" & ts & ".txt")
    logSessionActive = True
    Info "Log session started: " & logFilePath
End Sub

Private Sub WriteLine(ByVal level As String, ByVal message As String)
    Dim ts As String
    ts = Format(Now, "yyyy-MM-dd hh:nn:ss")

    Dim fullMessage As String
    fullMessage = "[" & ts & "] [" & UCase(level) & "] " & message

    ' write to PowerShell terminal
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").OpenTextFile("CON", 8, True).WriteLine fullMessage
    On Error GoTo 0

    ' Log file
    If logSessionActive Then
        Dim fso As Object, file As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(logFilePath, 8, True)
        file.WriteLine fullMessage
        file.Close
    End If
End Sub


' Public methods
Public Sub Info(ByVal message As String)
    WriteLine "INFO", message
End Sub

Public Sub Warn(ByVal message As String)
    WriteLine "WARN", message
End Sub

Public Sub Error(ByVal message As String)
    WriteLine "ERROR", message
End Sub

Public Sub Success(ByVal message As String)
    WriteLine "SUCCESS", message
End Sub

Public Sub EndSession()
    Info "Log session ended."
    logSessionActive = False
End Sub
