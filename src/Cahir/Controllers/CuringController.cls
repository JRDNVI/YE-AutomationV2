VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CuringImportController"
Option Explicit

' ===========================================
' CLASS: CuringImportController
' ===========================================

Private config As ConfigCahir
Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
Private curingHelper As ReadWriteCuringHallYE

Public Sub Init(ByRef cfg As ConfigCahir, _
                ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)
                
    Set config = cfg
    Set logger = lg
    Set errHandler = eh
    Set state = st
    Set curingHelper = New ReadWriteCuringHallYE
End Sub

Public Sub Run()
    On Error GoTo ERR_HANDLER

    logger.Info "Curing Import Started"

    Dim targetDate As Date: targetDate = config.GetTargetDate
    Dim masterPath As String: masterPath = config.GetPath("master_file")
    Dim rootPath As String: rootPath = config.GetPath("supervisor_root")

    Dim shifts As Variant: shifts = config.GetShifts
    Dim curingPrefixes As Variant: curingPrefixes = config.GetCuringPrefixes
    Dim thermoColMap As Variant: thermoColMap = config.GetThermoColumns
    Dim smokingColMap As Variant: smokingColMap = config.GetSmokingColumns

    Dim shift As Variant, prefix As Variant
    Dim folder As String, fileName As String
    Dim yr As String, mn As String
    yr = CStr(Year(targetDate))
    mn = Format$(targetDate, "mmmm")

    For Each shift In shifts
        For Each prefix In curingPrefixes
            folder = rootPath & shift & "\Curing Hall\" & yr & "\" & mn & "\"
            fileName = GetLatestCuringFile(folder, targetDate, CStr(prefix))

            If Len(fileName) = 0 Then
                logger.Warn "No file found for " & CStr(prefix) & " in " & folder
            Else
                logger.Info "Found file: " & fileName
                If InStr(1, UCase$(prefix), "THERMO", vbTextCompare) > 0 Then
                    curingHelper.ReadThermoData fileName, masterPath, CStr(shift), thermoColMap, targetDate, logger
                Else
                    curingHelper.ReadSmokingData fileName, masterPath, CStr(shift), smokingColMap, targetDate, logger
                End If
            End If
        Next prefix
    Next shift

    Exit Sub

ERR_HANDLER:
    If Not logger Is Nothing Then
        logger.Error "CuringImportController.Run error " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetLatestCuringFile(ByVal folder As String, _
                                     ByVal targetDate As Date, _
                                     ByVal prefix As String) As String
    On Error GoTo EH

    Dim f As String, fullPath As String
    Dim expectedLong As String, expectedShort As String
    Dim fallbackNameLong As String, fallbackNameShort As String
    Dim clean As String

    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    expectedLong = Format$(targetDate, "dd.mm.yyyy")
    expectedShort = Format$(targetDate, "dd.mm.yy")

    fallbackNameLong = folder & expectedLong & ".xlsx"
    fallbackNameShort = folder & expectedShort & ".xlsx"

    f = Dir(folder & "*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And InStr(1, f, "backup", vbTextCompare) = 0 Then
            
            clean = UCase$(Replace$(Replace$(f, "_", " "), "  ", " "))
            clean = Trim$(Replace$(clean, "  ", " "))
            
            If Left$(clean, Len(prefix)) = UCase$(prefix) _
               And (InStr(1, clean, expectedLong, vbTextCompare) > 0 _
                Or InStr(1, clean, expectedShort, vbTextCompare) > 0) Then
                    GetLatestCuringFile = folder & f
                    Exit Function
            End If
        End If

        f = Dir
    Loop

    If Len(Dir(fallbackNameLong)) > 0 Then
        GetLatestCuringFile = fallbackNameLong
        Exit Function
    End If

    If Len(Dir(fallbackNameShort)) > 0 Then
        GetLatestCuringFile = fallbackNameShort
        Exit Function
    End If

    GetLatestCuringFile = ""
    Exit Function

EH:
    logger.Error "GetLatestCuringFile error " & Err.Number & ": " & Err.Description
    GetLatestCuringFile = ""
End Function
