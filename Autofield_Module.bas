Attribute VB_Name = "Autofield_Module"
Sub autofield_report()
    Dim startendo As Integer
Dim lastendo As Integer
Dim area As Integer
Dim i As Long
Dim a As Long
Dim field As Long
Dim listarea As Long
Dim areaforfield As Long

Worksheets("AutoField").Rows("1:" & Rows.count).ClearContents

startendo = 2
lastendo = Worksheets("Area Break").Range("B" & Rows.count).End(xlUp).Row

areaforfield = Worksheets("Menu").Range("L" & Rows.count).End(xlUp).Row

area = 1


For i = startendo To lastendo

For field = 7 To areaforfield

If Worksheets("Area Break").Range("D" & i).Value = Worksheets("Menu").Range("L" & field).Value Then

If Worksheets("Area Break").Range("B" & i).Value = "Pri" Then
If Worksheets("Area Break").Range("D" & i).Value <> "NCR" Then
Worksheets("AutoField").Range("D" & area).Value = Worksheets("Area Break").Range("D" & i).Value
Else
Worksheets("AutoField").Range("D" & area).Value = Worksheets("Area Break").Range("E" & i).Value
End If





Worksheets("AutoField").Range("B" & area).Value = "primary address"

Worksheets("AutoField").Range("A" & area).Value = Worksheets("Area Break").Range("A" & i).Value
Worksheets("AutoField").Range("C" & area).Value = "DL2"
Worksheets("AutoField").Range("E" & area).Value = "Demand_letter"

Else

End If

Else

End If

Next field


area = area + 1


Next i
Worksheets("AutoField").Activate
  ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    Range("A1").Select
    
 Range("A1").EntireRow.Insert
Range("A1:E1000").Sort _
Key1:=Range("D1"), Order1:=xlAscending

Dim wb As Workbook
Set wb = Workbooks.Add
ThisWorkbook.Sheets("AutoField").Copy Before:=wb.Sheets(1)
wb.SaveAs "\\192.168.15.252\admins\JAM\PSB AUTO LOAN\AUTO FIELD\" & "PSB AUTO LOAN - AUTO FIELD " & Format(Now(), "yyyyMMMDD") & "  (PRIMARY).xls", FileFormat:=56

wb.Close
ThisWorkbook.Activate

Worksheets("AutoField").Rows("1:" & Rows.count).ClearContents

startendo = 2
lastendo = Worksheets("Area Break").Range("B" & Rows.count).End(xlUp).Row

area = 1


For i = startendo To lastendo


For field = 7 To areaforfield


If Worksheets("Area Break").Range("D" & i).Value = Worksheets("Menu").Range("L" & field).Value Then

If Worksheets("Area Break").Range("B" & i).Value = "Sec" Then
If Worksheets("Area Break").Range("D" & i).Value <> "NCR" Then
Worksheets("AutoField").Range("D" & area).Value = Worksheets("Area Break").Range("D" & i).Value
Else
Worksheets("AutoField").Range("D" & area).Value = Worksheets("Area Break").Range("E" & i).Value
End If





Worksheets("AutoField").Range("B" & area).Value = "secondary address"

Worksheets("AutoField").Range("A" & area).Value = Worksheets("Area Break").Range("A" & i).Value
Worksheets("AutoField").Range("C" & area).Value = "DL2"
Worksheets("AutoField").Range("E" & area).Value = "Demand_letter"

Else
area = area - 1
End If


Else

End If

Next field




area = area + 1


Next i
Worksheets("AutoField").Activate
  ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
  
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("A1").Select
    
 Range("A1").EntireRow.Insert
Range("A1:E1000").Sort _
Key1:=Range("D1"), Order1:=xlAscending
'Range("A1").EntireRow.Delete

Set wb = Workbooks.Add
ThisWorkbook.Sheets("AutoField").Copy Before:=wb.Sheets(1)
wb.SaveAs "\\192.168.15.252\admins\JAM\PSB AUTO LOAN\AUTO FIELD\" & "PSB AUTO LOAN - AUTO FIELD " & Format(Now(), "yyyyMMMDD") & "  (SECONDARY).xls", FileFormat:=56

wb.Close
ThisWorkbook.Activate

End Sub
