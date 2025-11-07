VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Utils"

Option Explicit


Public Sub NormaliseLineAndShift(ByRef lineName As String, ByRef shiftName As String)

    If InStr(1, shiftName, "Night", vbTextCompare) > 0 Then
        shiftName = "Evenings"
    ElseIf InStr(1, shiftName, "Day", vbTextCompare) > 0 Then
        shiftName = "Days"
    End If

   
    If shiftName = "Days" And lineName = "Line 1" Then
        lineName = "Line 1 Days"

    ElseIf shiftName = "Days" And lineName = "Line 2 & 3" Then
        lineName = "Line 2/3 Days"

    ElseIf lineName = "Line 4" Then
        lineName = "Jointing Line"

    ElseIf shiftName = "Evenings" And lineName = "Line 1" Then
        lineName = "Line 1 Nights"

    ElseIf shiftName = "Evenings" And lineName = "Line 2 & 3" Then
        lineName = "Line 2/3 Nights"
    End If

End Sub



Public Function FindColumnByHeader(ws As Worksheet, headerName As Variant) As String
    Dim headerRow As Long: headerRow = 8
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(headerRow, c).Value)) = headerName Then
            FindColumnByHeader = ColumnLetterFromNumber(c)
            Exit Function
        End If
    Next c
End Function



Public Function ColumnLetterFromNumber(ByVal colNum As Long) As String
    ' Converts column number (e.g., 2) → letter (B)
    Dim colLetter As String
    colLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
    ColumnLetterFromNumber = colLetter
End Function



Public Function ColLetterToNum(ByVal colLetter As String) As Long
    Dim i As Long, n As Long
    colLetter = UCase$(Trim$(colLetter))
    For i = 1 To Len(colLetter)
        n = n * 26 + (Asc(Mid$(colLetter, i, 1)) - 64)
    Next
    ColLetterToNum = n
End Function

