VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ReadYESheet"

Option Explicit

' ===========================================
' CLASS: ReadYESheet
' ===========================================
' Reads supervisor Y&E sheets into memory and
' writes collected data to the master workbook
' in one bulk operation.
' ===========================================

Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
private config As ConfigCahir
Private utils As Utils

Public Sub Init(ByRef lg As LogHelper, ByRef eh As ErrorHandler, ByRef st As ExcelStateManager)
    Set logger = lg
    Set errHandler = eh
    Set state = st
    Set utils = New Utils
    set config = New ConfigCahir
End Sub

Public Function ReadSupervisorData(ByVal filePath As String, _
                                   ByVal lineName As String, _
                                   ByVal shiftName As String, _
                                   ByVal targetDate As Date, _
                                   ByRef colMap As Object) As Collection
    On Error GoTo ERR_HANDLER

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, key As Variant, idx As Long
    Dim tempRow() As Variant
    Dim dataList As New Collection
    Dim supervisor As String

    Dim qaLogPath As String
        qaLogPath = config.GetPath("log_folder") & "DataQuality_" & Format(Now(), "yyyymmdd_hhnnss") & ".txt"


    Const DATA_START_ROW As Long = 9
    Const SUPERVISOR_CELL As String = "J7"
    Const END_MARKER_COL As String = "I"

    If Dir(filePath) = "" Then
        logger.Warn "File not found: " & filePath
        Exit Function
    End If

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    If wb Is Nothing Then
        logger.[Error] "Could not open workbook: " & filePath
        Exit Function
    End If

    On Error Resume Next
    Set ws = wb.Sheets("Yields and Efficiency")
    On Error GoTo ERR_HANDLER
    If ws Is Nothing Then
        logger.Warn "Missing 'Yields and Efficiency' sheet in " & wb.Name
        wb.Close False
        Exit Function
    End If

    supervisor = Trim(ws.Range(SUPERVISOR_CELL).Text)
    logger.Info "Supervisor: " & IIf(supervisor <> "", supervisor, "Unknown")

    i = DATA_START_ROW
    Do While Not IsEmpty(ws.Range(END_MARKER_COL & i).Value) _
         And ws.Range(END_MARKER_COL & i).Value <> 0

        ReDim tempRow(1 To colMap.Count + 4)
        idx = 1

        For Each key In colMap.Keys
            Dim srcCol As String
            Dim rawVal As Variant

            srcCol = Utils.FindColumnByHeader(ws, key)

            If srcCol <> "" Then
                rawVal = ws.Range(srcCol & i).Value
                tempRow(idx) = rawVal

                If IsError(rawVal) Or Trim$(CStr(rawVal & "")) = "" Then
                    Call LogDataQualityIssue(qaLogPath, filePath, lineName, shiftName, key, srcCol, i, rawVal)
                End If
            Else
                tempRow(idx) = ""
                Call LogDataQualityIssue(qaLogPath, filePath, lineName, shiftName, key, "(No Column Found)", i, "(Missing Column)")
            End If

            idx = idx + 1
        Next key

     
        Call Utils.NormaliseLineAndShift(lineName, shiftName)
   
        tempRow(idx) = targetDate: idx = idx + 1
        tempRow(idx) = lineName: idx = idx + 1       
        tempRow(idx) = shiftName: idx = idx + 1       
        tempRow(idx) = supervisor

        dataList.Add tempRow
        i = i + 1
        If i > 2000 Then
            logger.Warn "Safety stop triggered (row > 2000) in " & wb.Name
            Exit Do
        End If
    Loop

    Set ReadSupervisorData = dataList
    logger.Success "Collected " & dataList.Count & " valid rows from " & wb.Name
    wb.Close SaveChanges:=False
    Exit Function

ERR_HANDLER:
    logger.Error "Unexpected error in ReadSupervisorData: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If ReadSupervisorData Is Nothing Then Set ReadSupervisorData = New Collection
    Resume Next
End Function

Public Sub WriteToMaster(ByRef rows As Collection, _
                         ByVal masterPath As String, _
                         ByRef defaultColMap As Object)

    On Error GoTo ERR_HANDLER

    Dim masterWB As Workbook, masterWS As Worksheet
    Dim lastRowDst As Long, i As Long, idx As Long
    Dim rowArr As Variant, key As Variant, dstCol As String
    Dim sheetName As String
    Dim activeMap As Object
    Dim lineName As String

    Const DEST_DATE_COL As String = "A"
    Const DEST_LINE_COL As String = "D"
    Const DEST_SHIFT_COL As String = "E"
    Const DEST_SUPERVISOR_COL As String = "R"
    sheetName = "Yield & Efficiencies Lines"

    If rows Is Nothing Or rows.Count = 0 Then
        logger.Warn "No rows to write for this line."
        Exit Sub
    End If

    Set masterWB = Workbooks.Open(masterPath, UpdateLinks:=0, ReadOnly:=False, Notify:=False)
    Set masterWS = masterWB.Sheets(sheetName)

    lastRowDst = masterWS.Cells(masterWS.Rows.Count, "A").End(xlUp).Row + 1

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For i = 1 To rows.Count
        rowArr = rows(i)
        idx = 1

        lineName = Trim$(CStr(rowArr(UBound(rowArr) - 2)))

        If LCase$(lineName) Like "*jointing*" Or LCase$(lineName) = "line 4" Then
            Set activeMap = config.GetLineMap("Line 4")
        Else
            Set activeMap = defaultColMap
        End If

        For Each key In activeMap.Keys
            dstCol = CStr(activeMap(key))
            masterWS.Cells(lastRowDst, Utils.ColLetterToNum(dstCol)).Value = rowArr(idx)'
            idx = idx + 1
        Next key

        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_DATE_COL)).Value = rowArr(idx): idx = idx + 1'
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_LINE_COL)).Value = rowArr(idx): idx = idx + 1'
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_SHIFT_COL)).Value = rowArr(idx): idx = idx + 1'
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_SUPERVISOR_COL)).Value = rowArr(idx)'

        lastRowDst = lastRowDst + 1
    Next i

CLEANUP:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Not masterWB Is Nothing Then masterWB.Close SaveChanges:=True
    Set masterWS = Nothing
    Set masterWB = Nothing
    Exit Sub

ERR_HANDLER:
    errHandler.HandleError "ReadYESheet", "WriteToMaster", Erl
    Resume CLEANUP
End Sub

Private Sub LogDataQualityIssue(ByVal logPath As String, _
                                ByVal filePath As String, _
                                ByVal lineName As String, _
                                ByVal shiftName As String, _
                                ByVal fieldName As String, _
                                ByVal srcCol As String, _
                                ByVal rowNum As Long, _
                                ByVal value As Variant)

    Dim f As Integer: f = FreeFile
    Open logPath For Append As #f
    Print #f, "[" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "] Missing/Invalid Field:"
    Print #f, "  File: " & filePath
    Print #f, "  Line: " & lineName & " | Shift: " & shiftName
    Print #f, "  Field: " & fieldName & " (Column " & srcCol & ")"
    Print #f, "  Row: " & rowNum
    Print #f, "  Value: " & CStr(value)
    Print #f, String(60, "-")
    Close #f
End Sub
