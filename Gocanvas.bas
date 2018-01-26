Attribute VB_Name = "Gocanvas"
Sub Part1()

Dim newwb As Workbook
Dim LastRow As Long
Dim newwblast As Long

Worksheets("Database").Activate
 With ActiveSheet
        LastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    End With
Set newwb = Workbooks.Add
ThisWorkbook.ActiveSheet.Range("A1:O" & LastRow).Copy Destination:=newwb.Sheets(1).Range("A1:O" & LastRow)
Set newwb = ActiveWorkbook
 With ActiveSheet
        newwblast = .Cells(.Rows.count, "A").End(xlUp).Row + 1
 End With
 MsgBox newwblast
ThisWorkbook.ActiveSheet.Range("A1:N" & LastRow).Copy Destination:=newwb.Sheets(1).Range("A" & newwblast & ":N" & LastRow + newwblast)

ThisWorkbook.ActiveSheet.Range("P1:P" & LastRow).Copy Destination:=newwb.Sheets(1).Range("O" & newwblast & ":O" & LastRow + newwblast)

newwb.Sheets(1).Columns(5).EntireColumn.Delete

 On Error Resume Next
   newwb.Sheets(1).Columns("N").SpecialCells(xlBlanks).EntireRow.Delete
    
With ActiveSheet
        LastRow = .Cells(.Rows.count, "A").End(xlUp).Row
End With

Dim i As Long
Dim newval As String
For i = 2 To LastRow
newval = " '00" & newwb.Sheets(1).Range("A" & i).Value

newwb.Sheets(1).Range("A" & i).Value = newval
newwb.Sheets(1).Range("E" & i).Value = newval
Next i
newwb.Sheets(1).Columns("A").NumberFormat = "@"

newwb.SaveAs Filename:="\\192.168.15.252\admins\JAM\PSB AUTO LOAN\Go Canvas\psb auto part1.csv", FileFormat:=xlCSV
newwb.Close


End Sub

Sub part2()
Dim newwb As Workbook
Dim LastRow As Long
Dim newwblast As Long

 With ActiveSheet
        LastRow = .Cells(.Rows.count, "A").End(xlUp).Row
    End With
Set newwb = Workbooks.Add
ThisWorkbook.ActiveSheet.Range("B1:B" & LastRow).Copy Destination:=newwb.Sheets(1).Range("A1:A" & LastRow)
ThisWorkbook.ActiveSheet.Range("O1:O" & LastRow).Copy Destination:=newwb.Sheets(1).Range("B1:B" & LastRow)
ThisWorkbook.ActiveSheet.Range("Q1:Q" & LastRow).Copy Destination:=newwb.Sheets(1).Range("C1:C" & LastRow)
ThisWorkbook.ActiveSheet.Range("T1:Y" & LastRow).Copy Destination:=newwb.Sheets(1).Range("D1:I" & LastRow)

 With ActiveSheet
        newwblast = .Cells(.Rows.count, "A").End(xlUp).Row + 1
 End With

ThisWorkbook.ActiveSheet.Range("B1:B" & LastRow).Copy Destination:=newwb.Sheets(1).Range("A" & newwblast & ":A" & LastRow + newwblast)
ThisWorkbook.ActiveSheet.Range("P1:P" & LastRow).Copy Destination:=newwb.Sheets(1).Range("B" & newwblast & ":B" & LastRow + newwblast)
ThisWorkbook.ActiveSheet.Range("Q1:Q" & LastRow).Copy Destination:=newwb.Sheets(1).Range("C" & newwblast & ":C" & LastRow + newwblast)
ThisWorkbook.ActiveSheet.Range("T1:U" & LastRow).Copy Destination:=newwb.Sheets(1).Range("D" & newwblast & ":E" & LastRow + newwblast)
ThisWorkbook.ActiveSheet.Range("Z1:AB" & LastRow).Copy Destination:=newwb.Sheets(1).Range("F" & newwblast & ":H" & LastRow + newwblast)

 With ActiveSheet
        newwblast = .Cells(.Rows.count, "A").End(xlUp).Row + 1
 End With

newwb.Sheets(1).Range("I1").Value = "PRI AREA COUNT"
newwb.Sheets(1).Range("J1").Value = "PRI MUNICIPALITY COUNT"
newwb.Sheets(1).Range("K1").Value = "PRI BARANGAY COUNT"
newwb.Sheets(1).Range("L1").Value = "COUNT OF Address"

newwb.Sheets(1).Range("I2:I" & newwblast).Formula = "=COUNTIF(F:F,F2)"
newwb.Sheets(1).Range("J2:J" & newwblast).Formula = "=COUNTIF(G:G,G2)"
newwb.Sheets(1).Range("K2:K" & newwblast).Formula = "=COUNTIF(H:H,H2)"

 On Error Resume Next
    newwb.Sheets(1).Columns("B").SpecialCells(xlBlanks).EntireRow.Delete
newwb.Sheets(1).Range("L2:L" & newwblast).Formula = "=COUNTIF(A:A,A2)"


With newwb.Sheets(1).UsedRange
        .Value = .Value
    End With
'With ActiveSheet
'        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
'End With
'
'Dim i As Long
'Dim newval As String
'For i = 1 To lastrow
'newval = "'" & newwb.Sheets(1).Range("A" & i).Value
'newwb.Sheets(1).Range("A" & i).Value = newval
'Next i

newwb.SaveAs Filename:="\\192.168.15.252\admins\JAM\PSB AUTO LOAN\Go Canvas\psb auto part2.csv", FileFormat:=xlCSV
newwb.Close

End Sub
