VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ErrorHandler"

Option Explicit

Private logger As LogHelper
Private showMsgBox As Boolean

' Initialise
Public Sub Init(ByRef logRef As LogHelper, Optional ByVal displayMsgBox As Boolean = False)
    Set logger = logRef
    showMsgBox = displayMsgBox
End Sub



' HandleError
' Called from any controller in an ERR_HANDLER block.
' Logs error details, source, and optionally shows a message.
Public Sub HandleError(ByVal moduleName As String, _
                       ByVal procedureName As String, _
                       Optional ByVal lineNumber As Long = 0)
    On Error Resume Next

    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim msg As String
    Dim ts As String

    ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source

    msg = "[" & ts & "] " & _
          "Error " & errNum & " in " & moduleName & "." & procedureName & _
          IIf(lineNumber > 0, " (Line " & lineNumber & ")", "") & ": " & errDesc


    If Not logger Is Nothing Then
        logger.Error msg
    Else
        Debug.Print msg
    End If

    ' Optional on-screen alert
    If showMsgBox Then
        MsgBox msg, vbCritical + vbOKOnly, "Error"
    End If

    Err.Clear
End Sub



' RaiseError
' Use this to raise custom/controlled exceptions
Public Sub RaiseError(ByVal moduleName As String, _
                      ByVal procedureName As String, _
                      ByVal description As String, _
                      Optional ByVal errNum As Long = 9999)
    On Error Resume Next
    Err.Raise errNum, moduleName & "." & procedureName, description
End Sub
