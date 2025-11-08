VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 
END
Attribute VB_Name = "UpdatePivotTables"
Option Explicit


Public Sub UpdateAllPivotTables(ByVal masterSlicingPath As String, _
                                ByVal masterDowntimePath As String, _
                                ByVal pivotTableNames As Object, _
                                ByVal targetDate As Date, _
                                ByRef logger As LogHelper)

    Dim wb As Workbook, ws As Worksheet
    Dim sheetName As Variant, pivts As Variant, i As Long
    Dim pt As PivotTable

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Open(masterSlicingPath, UpdateLinks:=0, ReadOnly:=False)

    For Each sheetName In pivotTableNames.Keys
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(CStr(sheetName))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet

        ' Print all Pivot fields, to see why thermo isn't updating 

        pivts = pivotTableNames(sheetName) ' this is an array
        For i = LBound(pivts) To UBound(pivts)
            Set pt = Nothing
            On Error Resume Next
            Set pt = ws.PivotTables(CStr(pivts(i)))
            On Error GoTo 0
            If Not pt Is Nothing Then
                logger.Info "Refreshing Pivot: " & pt.Name & " (Sheet: " & ws.Name & ")"
                pt.RefreshTable
                On Error Resume Next
                With pt.PivotFields("Date")
                    .ClearAllFilters
                    ' Works when Date is a Page field
                    .CurrentPage = targetDate
                End With
                On Error GoTo 0
            End If
        Next i
NextSheet:
    Next sheetName

    wb.Close SaveChanges:=True

    ' --- Downtime workbook
    Set wb = Workbooks.Open(masterDowntimePath, UpdateLinks:=0, ReadOnly:=False)
    On Error Resume Next
    Set ws = wb.Sheets("Report Summary")
    If Not ws Is Nothing Then
        Set pt = ws.PivotTables("PivotTable1")
        If Not pt Is Nothing Then
            logger.Info "Refreshing Downtime PivotTable1"
            pt.RefreshTable
            With pt.PivotFields("Date")
                .ClearAllFilters
                .CurrentPage = targetDate
            End With
        End If
    End If
    On Error GoTo 0
    wb.Close SaveChanges:=True

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    logger.Error "Error in UpdateAllPivotTables: " & Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub

