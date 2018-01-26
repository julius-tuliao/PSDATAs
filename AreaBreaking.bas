Attribute VB_Name = "AreaBreaking"
Sub area()
Dim startendo As Integer
Dim lastendo As Integer
Dim area As Integer
Dim i As Long
Dim a As Long
Dim listarea As Long

Worksheets("Area Break").Rows("2:" & Rows.count).ClearContents

startendo = Worksheets("Menu").Range("H7").Value
lastendo = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row

area = 1


For i = startendo To lastendo

area = area + 1
If Len(Worksheets("Database").Range("O" & startendo)) > 4 Then
Worksheets("Area Break").Range("A" & area).Value = Worksheets("Database").Range("B" & i).Value
Worksheets("Area Break").Range("B" & area).Value = "Pri"
Worksheets("Area Break").Range("C" & area).Value = Worksheets("Database").Range("o" & i).Value
End If

Next i

area = Worksheets("Area Break").Range("A" & Rows.count).End(xlUp).Row


For i = startendo To lastendo

area = area + 1
If Len(Worksheets("Database").Range("O" & startendo)) > 4 Then
Worksheets("Area Break").Range("A" & area).Value = Worksheets("Database").Range("B" & i).Value
Worksheets("Area Break").Range("B" & area).Value = "Sec"
Worksheets("Area Break").Range("C" & area).Value = Worksheets("Database").Range("p" & i).Value
End If

Next i

area = Worksheets("Area Break").Range("A" & Rows.count).End(xlUp).Row + 1
listarea = Worksheets("List Of Areas").Range("A" & Rows.count).End(xlUp).Row + 1


For a = 2 To area

For i = 2 To listarea

If Len(Worksheets("List Of Areas").Range("C" & i).Value) > 3 Then

If (InStr(1, Worksheets("Area Break").Range("C" & a).Value, Worksheets("List Of Areas").Range("A" & i).Value) And InStr(1, Worksheets("Area Break").Range("C" & a).Value, Worksheets("List Of Areas").Range("B" & i).Value) And InStr(1, Worksheets("Area Break").Range("C" & a).Value, Worksheets("List Of Areas").Range("C" & i).Value)) Then
 
 Worksheets("Area Break").Range("D" & a).Value = Worksheets("List Of Areas").Range("E" & i).Value
 Worksheets("Area Break").Range("E" & a).Value = Worksheets("List Of Areas").Range("F" & i).Value
  Worksheets("Area Break").Range("F" & a).Value = Worksheets("List Of Areas").Range("G" & i).Value
 Exit For
 


 
End If

Else

If (InStr(1, Worksheets("Area Break").Range("C" & a).Value, Worksheets("List Of Areas").Range("A" & i).Value) And InStr(1, Worksheets("Area Break").Range("C" & a).Value, Worksheets("List Of Areas").Range("B" & i).Value)) Then
 
 Worksheets("Area Break").Range("D" & a).Value = Worksheets("List Of Areas").Range("E" & i).Value
 Worksheets("Area Break").Range("E" & a).Value = Worksheets("List Of Areas").Range("F" & i).Value
 Exit For
 
 
End If

End If

Next i

Next a

'With Worksheets("Area Break").UsedRange
'.Value = .Value
'End With
'
Worksheets("Area Break").Activate
Rows("1:1").Select
    Selection.AutoFilter
Worksheets("Area Break").Range("$A$1:$G$" & Rows.count).AutoFilter field:=4, Criteria1:=""
    Worksheets("Area Break").Range("A1").Select



End Sub




'Formulas
'Province =IFERROR(VLOOKUP(INDEX(INDIRECT("'List Of Areas'!$A$3:$A$" & 'List Of Areas'!$E$3),MATCH(TRUE,ISNUMBER(SEARCH(INDIRECT("'List Of Areas'!$A$3:$A$" & 'List Of Areas'!$E$3),C2)),0)),'List Of Areas'!J:K,2,0),"NO DATA")
'City =IFERROR(VLOOKUP(INDEX(INDIRECT("'List Of Areas'!$B$3:$B$" & 'List Of Areas'!$F$3),MATCH(TRUE,ISNUMBER(SEARCH(INDIRECT("'List Of Areas'!$B$3:$B$" & 'List Of Areas'!$F$3),C2)),0)),'List Of Areas'!L:M,2,0),"NO DATA")
'brgy =IFERROR(VLOOKUP(INDEX(INDIRECT("'List Of Areas'!$C$3:$C$" & 'List Of Areas'!$G$3),MATCH(TRUE,ISNUMBER(SEARCH(INDIRECT("'List Of Areas'!$C$3:$C$" & 'List Of Areas'!$G$3),C2)),0)),'List Of Areas'!N:O,2,0),"NO DATA")


