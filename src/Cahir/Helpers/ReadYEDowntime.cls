VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ReadYEDowntime"
Option Explicit

' ===========================================
' CLASS: ReadYEDowntime
' ===========================================
' Reads downtime/breakdown sections from supervisor sheets.
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


Public Sub Init(ByRef lg As LogHelper, _
                ByRef eh As ErrorHandler, _
                ByRef st As ExcelStateManager)
    Set logger = lg
    Set errHandler = eh
    Set state = st
    Set utils = New Utils
    Set config = New ConfigCahir
End Sub


Public Function ReadDowntime(ByVal filePath As String, _
                             ByVal lineName As String, _
                             ByVal shift As String, _
                             ByVal targetDate As Date, _
                             ByRef colMap As Object, _
                             ByRef breakdownCol As Object) As Collection
    On Error GoTo ERR_HANDLER

    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, r As Long, lastRow2 As Long, r2 As Long
    Dim firstProductRow As Long, secondProductRow As Long, thirdProductRow As Long
    Dim startRow As Long, downtimeHeaderRow As Long, breakdownHeaderRow As Long, endRow As Long, thirdStartRow As Long, thirdEndRow As Long
    Dim rows As New Collection
    Dim valCheck As String

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Sheets("Yields and Efficiency")

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    logger.Info "Scanning column C (1-" & lastRow & ") for Product headers..."

    Dim productRows(1 To 3) As Long
    Dim count As Long: count = 0

    ' Find first, second, third "Product" header in Column C
    For r = 1 To lastRow
        valCheck = CStr(ws.Cells(r, "C").MergeArea.Cells(1, 1).Value & "")
        valCheck = Trim(LCase(valCheck))

        If valCheck = "product" Then
            count = count + 1
            If count <= 3 Then productRows(count) = r
            If count = 3 Then Exit For
        End If
    Next r

    firstProductRow = productRows(1)
    secondProductRow = productRows(2)
    thirdProductRow = productRows(3)

    If secondProductRow = 0 Then
        logger.Warn "Downtime header not found in " & filePath
        GoTo CLEANUP
    End If

   ' Downtime section
    downtimeHeaderRow = secondProductRow
    startRow = downtimeHeaderRow + 1
    endRow = FindSectionEnd(ws, startRow, "B", lastRow)

    logger.Info "Downtime section: header=" & downtimeHeaderRow & ", first=" & startRow & ", last=" & endRow

    '  Breakdown section
    If thirdProductRow > 0 Then
        breakdownHeaderRow = thirdProductRow
        thirdStartRow = thirdProductRow + 1
        thirdEndRow = FindSectionEnd(ws, thirdStartRow, "I", lastRow)

        If thirdStartRow <= thirdEndRow Then
            logger.Info "Breakdown Section: header=" & breakdownHeaderRow & ", first=" & thirdStartRow & ", last=" & thirdEndRow
        Else
            logger.Info "Breakdown Section present but contains **no rows** — skipped."
        End If
    End If


    Dim fixedMap As Object: Set fixedMap = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In colMap.Keys
        fixedMap(Trim(LCase(k))) = colMap(k)
    Next k
    Set colMap = fixedMap

    Dim c As Long, hdr As String, colAddr As String, colLetter As String
    Dim cell As Range

    For c = 1 To ws.Cells(downtimeHeaderRow, ws.Columns.Count).End(xlToLeft).Column
        Set cell = ws.Cells(downtimeHeaderRow, c).MergeArea.Cells(1, 1)

    ' Only map if this cell is the top-left cell of its merge area!!!!!!!!!!!!!!!!!!!!!!!!!!
        If cell.Address = ws.Cells(downtimeHeaderRow, c).Address Then
        
            hdr = Trim(LCase(cell.Value))

            If colMap.Exists(hdr) Then
                colAddr = cell.Address(False, False)
                Do While IsNumeric(Right(colAddr, 1))
                    colAddr = Left(colAddr, Len(colAddr) - 1)
                Loop
            
                colLetter = colAddr
                colMap(hdr) = colLetter
            End If
        End If
    Next c

    For r = startRow To endRow
        Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
        Dim key As Variant
    
        For Each key In colMap.Keys
            colLetter = colMap(key)
            entry(key) = Trim(CStr(ws.Range(colLetter & r).MergeArea.Cells(1, 1).Value))
        Next key

        entry("Line") = lineName
        entry("Shift") = shift
        entry("SupervisorFile") = Mid(filePath, InStrRev(filePath, "\") + 1)
        entry("Date") = targetDate
        rows.Add entry

    Next r

    ' Read Breakdown Section
    If thirdProductRow > 0 And thirdStartRow <= thirdEndRow Then

        ' Normalise breakdown column map
        Dim fixedBreak As Object: Set fixedBreak = CreateObject("Scripting.Dictionary")
        For Each k In breakdownCol.Keys
            fixedBreak(Trim(LCase(k))) = breakdownCol(k)
        Next k
        Set breakdownCol = fixedBreak

        ' Map headers for breakdown section
        For c = 1 To ws.Cells(breakdownHeaderRow, ws.Columns.Count).End(xlToLeft).Column
            Set cell = ws.Cells(breakdownHeaderRow, c).MergeArea.Cells(1, 1)

            If cell.Address = ws.Cells(breakdownHeaderRow, c).Address Then
                hdr = Trim(LCase(cell.Value))
                If breakdownCol.Exists(hdr) Then
                    colAddr = cell.Address(False, False)
                    Do While IsNumeric(Right(colAddr, 1))
                        colAddr = Left(colAddr, Len(colAddr) - 1)
                    Loop
                    breakdownCol(hdr) = colAddr
                End If
            End If
        Next c

        ' Extract breakdown rows
        For r = thirdStartRow To thirdEndRow
            Dim entryB As Object: Set entryB = CreateObject("Scripting.Dictionary")
        
            For Each key In breakdownCol.Keys
                colLetter = breakdownCol(key)
                entryB(key) = Trim(CStr(ws.Range(colLetter & r).MergeArea.Cells(1, 1).Value))
            Next key

            entryB("Line") = lineName
            entryB("Shift") = shift
            entryB("SupervisorFile") = Mid(filePath, InStrRev(filePath, "\") + 1)
            entryB("Date") = targetDate
            entryB("SectionType") = "Breakdown"

            rows.Add entryB
        Next r
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

Public Sub WriteToMaster(ByRef rows As Collection, _
                         ByVal masterPath As String, _
                         ByRef defaultColMap As Object, _ 
                         ByRef breakdownColMap As Object)

    On Error GoTo ERR_HANDLER

    Dim masterWB As Workbook, masterWS As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim i As Long, key As Variant, dstCol As String
    Dim rowArr As Object, activeMap As Object, lineName As String, dowwntimeType As String
    Dim machineCol As Long, maintCol As Long

    Const DEST_DATE_COL As String = "A"
    Const DEST_LINE_COL As String = "E"
    Const DEST_SHIFT_COL As String = "F"

    logger.Info "WriteToMaster() called — " & rows.Count & " rows pending for write."


    If rows Is Nothing Or rows.Count = 0 Then
        logger.Warn "No downtime rows to write."
        Exit Sub
    End If

    Set masterWB = Workbooks.Open(masterPath, UpdateLinks:=0, ReadOnly:=False, Notify:=False)
    Set masterWS = masterWB.Sheets("Master Downtime")
    Set tbl = masterWS.ListObjects("Table1") 

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim writtenCount As Long: writtenCount = 0

    For i = 1 To rows.Count
        Set rowArr = rows(i)
        dowwntimeType = LCase$(rowArr("SectionType"))
        lineName = LCase$(rowArr("Line"))

        If dowwntimeType = "breakdown" Then
            Set activeMap = breakdownColMap
            rowArr("total") = rowArr("duration of downtime")

            If LCase$(Left$(rowArr("product"), 5)) = "error" Then
                rowArr("product") = ""
            End If

        Else
            ' Downtime
            If (lineName Like "*jointing*") Or (lineName = "line 4") Then
                Set activeMap = config.GetLineProductionMap("Line 4")
            Else
                Set activeMap = defaultColMap
            End If

        End If


        Set newRow = tbl.ListRows.Add

        For Each key In activeMap.Keys
            dstCol = CStr(activeMap(key))
            masterWS.Cells(newRow.Range.Row, Utils.ColLetterToNum(dstCol)).Value = rowArr(LCase$(key))
        Next key

      
        With masterWS.Cells(newRow.Range.Row, Utils.ColLetterToNum(DEST_DATE_COL))
            .Value = CDate(rowArr("Date"))
            .NumberFormat = "dd/mm/yyyy"
        End With
        
        masterWS.Cells(newRow.Range.Row, Utils.ColLetterToNum(DEST_LINE_COL)).Value = rowArr("Line")
        masterWS.Cells(newRow.Range.Row, Utils.ColLetterToNum(DEST_SHIFT_COL)).Value = rowArr("Shift")

        machineCol = Utils.ColLetterToNum("R")
        maintCol = Utils.ColLetterToNum("Q")    

        If Trim(masterWS.Cells(newRow.Range.Row, machineCol).Value) = "" Then
            masterWS.Cells(newRow.Range.Row, maintCol).Value = "Production"
        Else
            masterWS.Cells(newRow.Range.Row, maintCol).Value = "Maintenance"
        End If

        writtenCount = writtenCount + 1
    Next i


CLEANUP:
    On Error Resume Next
    masterWB.Close SaveChanges:=True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Set masterWS = Nothing: Set masterWB = Nothing
    Exit Sub

ERR_HANDLER:
    logger.Error "Runtime error " & Err.Number & " (" & Err.Description & ") in WriteToMaster"
    errHandler.HandleError "ReadYEDowntime", "WriteToMaster", Erl
    Resume CLEANUP

End Sub


Private Function FindSectionEnd(ws As Worksheet, startRow As Long, col As String, lastRow As Long) As Long
    Dim r As Long, v As Variant
    FindSectionEnd = lastRow
    For r = startRow To lastRow
        v = Trim(LCase(ws.Cells(r, col).MergeArea.Cells(1, 1).Value))
        If v = "" Or v = "0" Then
            FindSectionEnd = r - 1
            Exit For
        End If
    Next r
End Function

Private Function DictToString(ByVal d As Object) As String
    Dim k As Variant, output As String
    For Each k In d.Keys
        output = output & k & "=" & CStr(d(k)) & "; "
    Next k
    If Len(output) > 2 Then output = Left(output, Len(output) - 2)
    DictToString = output
End Function




