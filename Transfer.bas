Attribute VB_Name = "Transfer"
Option Explicit

Sub Transfer_Click()
Dim last As Long
Dim first As Long
Dim i As Integer
Dim areastart As Long
Dim areastop As Long
Dim a As Long


first = Worksheets("Menu").Range("H7").Value

last = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row
areastop = Worksheets("Area Break").Range("A" & Rows.count).End(xlUp).Row

For a = 1 To 2

For i = first To last

For areastart = 2 To areastop

If Worksheets("Database").Range("B" & i).Value = Worksheets("Area Break").Range("A" & areastart).Value And Worksheets("Area Break").Range("B" & areastart).Value = "Pri" Then
Worksheets("Database").Range("V" & i).Value = Worksheets("Area Break").Range("D" & areastart).Value
Worksheets("Database").Range("W" & i).Value = Worksheets("Area Break").Range("E" & areastart).Value
Worksheets("Database").Range("X" & i).Value = Worksheets("Area Break").Range("F" & areastart).Value


ElseIf Worksheets("Database").Range("B" & i).Value = Worksheets("Area Break").Range("A" & areastart).Value And Worksheets("Area Break").Range("B" & areastart).Value = "Sec" Then
Worksheets("Database").Range("Z" & i).Value = Worksheets("Area Break").Range("D" & areastart).Value
Worksheets("Database").Range("AA" & i).Value = Worksheets("Area Break").Range("E" & areastart).Value
Worksheets("Database").Range("AB" & i).Value = Worksheets("Area Break").Range("F" & areastart).Value
Exit For
Else

End If



Next areastart


Next i







Next a


End Sub
