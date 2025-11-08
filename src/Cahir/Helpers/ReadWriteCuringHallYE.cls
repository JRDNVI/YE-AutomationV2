VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = "ReadWriteCuringHallYE"
Option Explicit

Private Const START_ROW As Long = 11
Private Const TARGET_SHEET As String = "Thermoforming"

Private Const START_ROW_SMOKE As Long = 6
Private Const TARGET_SHEET_SMOKE As String = "Smoking Yield & Efficiency"


Public Sub ReadThermoData(ByVal sourcePath As String, _
                          ByVal masterPath As String, _
                          ByVal shiftName As String, _
                          ByVal colMap As Variant, _
                          ByVal targetDate As Date, _
                          ByRef logger As LogHelper)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim tempRow() As Variant
    Dim dataList As Collection
    Dim cVal As Variant

    On Error GoTo EH
    logger.Info "Opening thermo file: " & sourcePath

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(sourcePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    Set dataList = New Collection
    i = START_ROW

    ' Read rows until column C is 0 or blank
    Do While Not IsEmpty(ws.Range("C" & i).Value) And ws.Range("C" & i).Value <> 0
        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value
        Next c
        dataList.Add tempRow
        i = i + 1
        If i > 2000 Then Exit Do  
    Loop

    wb.Close False
    logger.Info dataList.Count & " rows collected from " & shiftName

    WriteThermoToMaster dataList, shiftName, colMap, masterPath, logger, targetDate

    Exit Sub

EH:
    logger.Error "CuringHelper.ReadThermoData error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub


Private Sub WriteThermoToMaster(ByVal dataList As Collection, _
                                ByVal shiftName As String, _
                                ByVal colMap As Variant, _
                                ByVal masterPath As String, _
                                ByRef logger As LogHelper, _
                                ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long, count As Long

    On Error GoTo EH

    Set wb = Workbooks.Open(masterPath)
    Set ws = wb.Sheets(TARGET_SHEET)

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For Each dataRow In dataList
        With ws
            .Range("A" & nextRow).Value = targetDate
            .Range("C" & nextRow).Value = "Days"

            ' Map each value to its target column
            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c)
            Next c
        End With
        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    wb.Close SaveChanges:=True
    logger.Info count & " thermo rows written to master."
    Exit Sub

EH:
    logger.Error "CuringHelper.WriteThermoToMaster error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub


Public Sub ReadSmokingData(ByVal sourcePath As String, _
                           ByVal masterPath As String, _
                           ByVal shiftName As String, _
                           ByVal colMap As Variant, _
                           ByVal targetDate As Date, _
                           ByRef logger As LogHelper)

    Dim wb As Workbook, ws As Worksheet
    Dim i As Long, c As Long
    Dim tempRow() As Variant
    Dim dataList As Collection
    Dim cVal As Variant

    On Error GoTo EH

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(sourcePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True

    Set dataList = New Collection
    i = START_ROW_SMOKE

    Do While True
        cVal = ws.Range("C" & i).Value

        If IsError(cVal) Then Exit Do
        If IsEmpty(cVal) Then Exit Do
        If Trim(cVal & "") = "" Then Exit Do
        If IsNumeric(cVal) And cVal = 0 Then Exit Do

        ReDim tempRow(LBound(colMap) To UBound(colMap))
        For c = LBound(colMap) To UBound(colMap)
            tempRow(c) = ws.Range(colMap(c)(0) & i).Value
        Next c
        dataList.Add tempRow

        i = i + 1
        If i > 2000 Then Exit Do 
    Loop

    wb.Close False
    logger.Info dataList.Count & " smoking rows collected from " & shiftName

    WriteSmokingToMaster dataList, shiftName, colMap, masterPath, logger, targetDate
    Exit Sub

EH:
    logger.Error "SmokingHelper.ReadSmokingData error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub


Private Sub WriteSmokingToMaster(ByVal dataList As Collection, _
                                 ByVal shiftName As String, _
                                 ByVal colMap As Variant, _
                                 ByVal masterPath As String, _
                                 ByRef logger As LogHelper, _
                                 ByVal targetDate As Date)

    Dim wb As Workbook, ws As Worksheet
    Dim nextRow As Long, dataRow As Variant, c As Long, count As Long
    Dim prevCycle As Variant, currentCycle As Variant

    On Error GoTo EH

    Set wb = Workbooks.Open(masterPath)
    Set ws = wb.Sheets(TARGET_SHEET_SMOKE)

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For Each dataRow In dataList

        With ws
            .Range("A" & nextRow).Value = targetDate

            Select Case LCase$(shiftName)
                Case "day shift":   .Range("D" & nextRow).Value = "Days"
                Case "night shift": .Range("D" & nextRow).Value = "Evenings"
                Case Else:          .Range("D" & nextRow).Value = shiftName
            End Select

            For c = LBound(colMap) To UBound(colMap)
                .Range(colMap(c)(1) & nextRow).Value = dataRow(c)
            Next c
        End With

        
        currentCycle = ws.Range("C" & nextRow).Value  

        If count > 0 Then
            If currentCycle <> prevCycle Then
                ws.Range("K" & (nextRow - 1)).Value = 1996
            End If
        End If

        prevCycle = currentCycle
        nextRow = nextRow + 1
        count = count + 1
    Next dataRow

    ws.Range("K" & (nextRow - 1)).Value = 1996

    wb.Close SaveChanges:=True
    logger.Info count & " smoking rows written to master (with Target Weight applied)."
    Exit Sub

EH:
    logger.Error "SmokingHelper.WriteSmokingToMaster error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub
