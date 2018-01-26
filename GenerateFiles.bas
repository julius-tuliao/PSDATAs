Attribute VB_Name = "GenerateFiles"
Sub genfiles()
Dim LastRow As Long
Dim Start As Long
Dim LastCol As Long
Dim i As Long
Dim colname As String

Worksheets("BCRM").Rows("2:" & Rows.count).Clear

Start = Worksheets("Menu").Range("H" & 7).Value
LastRow = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row + 2

Worksheets("BCRM").Range("A2:N" & LastRow - Start).Value = Worksheets("Database").Range("D" & Start & ":Q" & LastRow).Value
Worksheets("BCRM").Range("S2:T" & LastRow - Start).Value = Worksheets("Database").Range("B" & Start & ":C" & LastRow).Value
Worksheets("BCRM").Range("R2:R" & LastRow - Start).Value = Worksheets("Database").Range("T" & Start & ":T" & LastRow).Value
'Worksheets("BCRM").Range("B2:B" & lastrow - Start).Value = Worksheets("Database").Range("A" & Start & ":A" & lastrow).Value
With Worksheets("BCRM")
 .Range("C2", "C" & LastRow - Start).NumberFormat = "#,##.00"
 .Range("R2", "R" & LastRow - Start).NumberFormat = "mm/dd/yyyy"
End With

Worksheets("BCRM").Range("R2:R" & LastRow - Start).Value = Worksheets("BCRM").Range("R2:R" & LastRow - Start).Value

With Worksheets("Database").Range("A" & Start & ":A" & LastRow)
 .Copy

End With

With Worksheets("BCRM").Range("B2:B" & LastRow - Start)

  .PasteSpecial xlPasteValues
End With


With Worksheets("LEADS")
        LastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
    End With

For i = 1 To LastCol
colname = Worksheets("LEADS").Cells(1, i).Value

If Len(colname) > 0 Then

Worksheets("Database").Range(colname & Start & ":" & colname & LastRow).Copy Destination:=Worksheets("LEADS").Cells(3, i)

'Range(Cells(1, a), Cells(1, z))
End If

Next i

With Worksheets("Database").Range("A" & Start & ":A" & LastRow)
 .Copy

End With

With Worksheets("LEADS").Range("E3")
  .PasteSpecial xlPasteValues
End With


LastRow = Worksheets("LEADS").Range("A" & Rows.count).End(xlUp).Row
Worksheets("LEADS").Range("AL3:AL" & LastRow).Formula = "=VLOOKUP(A3,Database!$B:$Q,16,0)"

With Worksheets("LEADS").Range("AL3:AL" & LastRow)
.Copy
  .PasteSpecial xlPasteValues
End With


Dim wb As Workbook
Set wb = Workbooks.Add
ThisWorkbook.Sheets("BCRM").UsedRange.Copy

'ThisWorkbook.Sheets("BCRM").Copy Before:=wb.Sheets(1)
'
'wb.Sheets("BCRM").Columns("R:R").NumberFormat = "mm/dd/yyyy"
'wb.Sheets("BCRM").Columns("C:C").NumberFormat = "#,##.00"
wb.Sheets(1).Range("A1").Select
wb.Sheets(1).PasteSpecial xlPasteValuesAndNumberFormats

wb.SaveAs "\\192.168.15.252\admins\JAM\PSB AUTO LOAN\BCRM FILES\" & "BCRM" & Format(Now(), "mm-dd-yyyy") & ".xls", FileFormat:=56
wb.Close


End Sub
