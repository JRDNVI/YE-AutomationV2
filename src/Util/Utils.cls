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

