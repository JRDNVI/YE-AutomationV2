VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ReadYEDowntime"
Option Explicit

' ===========================================
' CLASS: ReadYEDowntime
' ===========================================
' Reads downtime/production section from supervisor sheets.
' Second "Product" in column C marks header row.
' The column immediately to the left is "Product Code".
' Stops reading when "Total" appears in column C.
' Logs header, last row, and each downtime row collected.
' ===========================================

Private logger As LogHelper
Private errHandler As ErrorHandler
Private state As ExcelStateManager
Private config As ConfigCahir
Private Utils As Utils

' -------------------------------------------
' INITIALISE
' -------------------------------------------
Public Sub Init(ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)
    Set logger = lg
    Set errHandler = eh
    Set state = st
    Set utils = New Utils
    Set config = New ConfigCahir
End Sub


' -------------------------------------------
' READ DOWNTIME SECTION
' -------------------------------------------
Public Function ReadDowntime(ByVal filePath As String, _
                             ByVal lineName As String, _
                             ByVal shift As String, _
                             ByVal targetDate As Date, _
                             ByRef colMap As Object) As Collection
    On Error GoTo ERR_HANDLER

    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim firstProductRow As Long, secondProductRow As Long
    Dim startRow As Long, headerRow As Long, endRow As Long
    Dim rows As New Collection
    Dim valCheck As String

    ' --- Open workbook ---
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Sheets("Yields and Efficiency")
    logger.Info "In function"

    ' --- Find two "Product" rows in column C (robust to casing / spaces) ---
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    logger.Info "Scanning column C (1-" & lastRow & ") for Product headers..."

    For r = 1 To lastRow
        On Error Resume Next
        valCheck = Trim(LCase(ws.Cells(r, "C").MergeArea.Cells(1, 1).Value))
        On Error GoTo ERR_HANDLER
        If valCheck = "product" Then
            If firstProductRow = 0 Then
                firstProductRow = r
            Else
                secondProductRow = r
                Exit For
            End If
        End If
    Next r

    If secondProductRow = 0 Then
        logger.Warn "Downtime header not found in " & filePath
        GoTo CLEANUP
    End If

    headerRow = secondProductRow
    startRow = headerRow + 1


    endRow = lastRow
    Dim valB As Variant
    For r = startRow To lastRow
        On Error Resume Next
        valB = ws.Cells(r, "B").MergeArea.Cells(1, 1).Value
        On Error GoTo 0

        If Trim(valB) = "" Or Trim(LCase(valB)) = "0" Then
            endRow = r - 1
            logger.Info "Stopped at first blank/0 in column B at row " & r
            Exit For
        End If
    Next r

    logger.Info "Downtime section: header=" & headerRow & _
                ", first=" & startRow & ", last=" & endRow

    ' --- Dynamic column mapping ---
    Dim prodCol As Long: prodCol = ws.Range("C" & headerRow).Column
    Dim productCodeCol As Long: productCodeCol = prodCol - 1
    If colMap.Exists("Product Code") Then colMap("Product Code") = Split(ws.Cells(1, productCodeCol).Address(True, False), "$")(0)
    If colMap.Exists("Product") Then colMap("Product") = Split(ws.Cells(1, prodCol).Address(True, False), "$")(0)

    ' --- Extract rows ---
    For r = startRow To endRow
        Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
        Dim key As Variant, colLetter As String, val As Variant
        For Each key In colMap.Keys
            colLetter = colMap(key)
            On Error Resume Next
            val = ws.Range(colLetter & r).Value
            On Error GoTo ERR_HANDLER
            entry(key) = Trim(CStr(val))
        Next key
        entry("Line") = lineName
        entry("Shift") = shift
        entry("SupervisorFile") = Mid(filePath, InStrRev(filePath, "\") + 1)
        entry("Date") = targetDate
        rows.Add entry

        logger.Info " [ROW " & r & "] " & lineName & " | " & shift & _
                    " | Code: " & entry("Product Code") & _
                    " | Product: " & entry("Product") & _
                    " | Time: " & entry("Time") & _
                    " | Desc: " & Left(entry("Description of Issue"), 60)
    Next r

    If rows.Count > 0 Then
        logger.Success "Collected " & rows.Count & _
                       " downtime rows from " & Mid(filePath, InStrRev(filePath, "\") + 1)
    Else
        logger.Warn "No downtime rows found in " & filePath
    End If
    Set ReadDowntime = rows

CLEANUP:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    Exit Function

ERR_HANDLER:
    If Err.Number <> 0 Then
        logger.Error "Runtime error " & Err.Number & " (" & Err.Description & _
                     ") at row " & r & " in ReadDowntime"
        errHandler.HandleError "ReadYEDowntime", "ReadDowntime", Erl
    End If
    Resume CLEANUP
End Function




' -------------------------------------------
' WRITE TO MASTER (unchanged)
' -------------------------------------------
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
    Const DEST_LINE_COL As String = "E"
    Const DEST_SHIFT_COL As String = "F"
    sheetName = "Master Downtime"

    logger.Info "WriteToMaster() called — " & rows.Count & " rows pending for write."

    If rows Is Nothing Or rows.Count = 0 Then
        logger.Warn "No downtime rows to write."
        Exit Sub
    End If

    logger.Info "Opening master workbook: " & masterPath
    Set masterWB = Workbooks.Open(masterPath, UpdateLinks:=0, ReadOnly:=False, Notify:=False)

    logger.Info "Activating sheet: " & sheetName
    Set masterWS = masterWB.Sheets(sheetName)

    lastRowDst = masterWS.Cells(masterWS.Rows.Count, "A").End(xlUp).Row + 1
    logger.Info "Writing starts at master row: " & lastRowDst

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim writtenCount As Long
    writtenCount = 0

    ' --- Begin main write loop ---
    For i = 1 To rows.Count
        logger.Info "Processing row " & i & " of " & rows.Count
        Set rowArr = rows(i)
        idx = 1

        On Error Resume Next
        lineName = Trim$(CStr(rowArr(UBound(rowArr) - 2)))
        On Error GoTo ERR_HANDLER

        logger.Info "Line detected: " & lineName

        If LCase$(lineName) Like "*jointing*" Or LCase$(lineName) = "line 4" Then
            Set activeMap = config.GetLineProductionMap("Line 4")
            logger.Info "Using Line 4 mapping."
        Else
            Set activeMap = defaultColMap
            logger.Info "Using default mapping."
        End If

        ' --- Write mapped fields ---
        For Each key In activeMap.Keys
            dstCol = CStr(activeMap(key))
            logger.Info "Writing field '" & key & "' to column " & dstCol & _
                        " (value=" & CStr(rowArr(idx)) & ")"
            masterWS.Cells(lastRowDst, Utils.ColLetterToNum(dstCol)).Value = rowArr(idx)
            idx = idx + 1
        Next key

        ' --- Write metadata ---
        logger.Info "Writing metadata fields..."
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_DATE_COL)).Value = rowArr(idx): idx = idx + 1
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_LINE_COL)).Value = rowArr(idx): idx = idx + 1
        masterWS.Cells(lastRowDst, Utils.ColLetterToNum(DEST_SHIFT_COL)).Value = rowArr(idx)

        logger.Info "Completed writing row " & i & " to master (targetRow=" & lastRowDst & ")"

        lastRowDst = lastRowDst + 1
        writtenCount = writtenCount + 1
    Next i

    logger.Success "Wrote " & writtenCount & " downtime rows to master successfully."

CLEANUP:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Not masterWB Is Nothing Then
        logger.Info "Saving and closing master workbook..."
        masterWB.Close SaveChanges:=True
    End If

    Set masterWS = Nothing
    Set masterWB = Nothing

    logger.Info "WriteToMaster() finished normally."
    Exit Sub

ERR_HANDLER:
    logger.Error "Runtime error " & Err.Number & " (" & Err.Description & _
                 ") occurred in ReadYEDowntime.WriteToMaster at line " & Erl & _
                 ". LastRowDst=" & lastRowDst & ", idx=" & idx & ", key=" & key
    errHandler.HandleError "ReadYEDowntime", "WriteToMaster", Erl
    Resume CLEANUP
End Sub

