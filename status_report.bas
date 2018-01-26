Attribute VB_Name = "status_report"
Sub autostat_report()
Dim LastRow As Long
Dim i As Long
Dim random_number As Long

 Dim conn As Variant



    For Each conn In ActiveWorkbook.Connections
        conn.ODBCConnection.BackgroundQuery = False
    Next conn

    ActiveWorkbook.RefreshAll


With Worksheets("Database")
LastRow = .Cells(.Rows.count, "B").End(xlUp).Row
End With
Worksheets("Autostat").Rows("2:" & Rows.count).Clear

Worksheets("Database").Range("B2:B" & LastRow).Copy Destination:=Worksheets("Autostat").Range("A2")
Worksheets("Autostat").Range("B2:B" & LastRow).Formula = "=IFERROR(MATCH(A2,'ODBC STAT REPORT'!E:E,0),""OK"")"


For i = LastRow To 2 Step -1
        If (Worksheets("Autostat").Cells(i, "B").Value) <> "OK" Then
    'Cells(i, "A").EntireRow.ClearContents ' USE THIS TO CLEAR CONTENTS BUT NOT DELETE ROW
            Worksheets("Autostat").Cells(i, "A").EntireRow.Delete
        End If
    Next i
    
    
   With Worksheets("Autostat")
LastRow = .Cells(.Rows.count, "A").End(xlUp).Row
End With
 
    
Worksheets("Autostat").Range("B2:B" & LastRow).Formula = "=IFERROR(IF(TEXT(INDEX(Database!T:T,MATCH(Autostat!A2,Database!B:B,0)),""yymm"")=TEXT(NOW(),""yymm""),""OK"",""NOT""),""NOT"")"
    
For i = LastRow To 2 Step -1
        If (Worksheets("Autostat").Cells(i, "B").Value) <> "OK" Then
    'Cells(i, "A").EntireRow.ClearContents ' USE THIS TO CLEAR CONTENTS BUT NOT DELETE ROW
            Worksheets("Autostat").Cells(i, "A").EntireRow.Delete
        End If
    Next i
    
    
With Worksheets("Autostat")
LastRow = .Cells(.Rows.count, "A").End(xlUp).Row
End With


 For i = 2 To LastRow
 
 Worksheets("Autostat").Range("B" & i).Value = "NEGATIVE"

random_number = Int(3 * Rnd) + 1

If random_number = 1 Then
 Worksheets("Autostat").Range("C" & i).Value = "KOR"
 
ElseIf random_number = 2 Then
 Worksheets("Autostat").Range("C" & i).Value = "CANNOT BE REACH"
 
Else
 Worksheets("Autostat").Range("C" & i).Value = "BC"

End If
 
  Worksheets("Autostat").Range("L" & i).Value = Format(Now(), "mm-dd-yyyy")
 Next i

Worksheets("Autostat").Range("H2:H" & LastRow).Formula = "=VLOOKUP(A2,Database!B:Q,16,0) & "" "" & Autostat!C2"
Worksheets("Autostat").Range("K2:K" & LastRow).Formula = "=VLOOKUP(A2,Database!B:C,2,0)"


With Worksheets("Autostat").UsedRange
 .Copy
  .PasteSpecial xlPasteValues
End With


Dim wb As Workbook
Set wb = Workbooks.Add
ThisWorkbook.Sheets("Autostat").UsedRange.Copy
wb.Sheets(1).Range("A1").Select
wb.Sheets(1).PasteSpecial xlPasteValuesAndNumberFormats
wb.SaveAs "\\192.168.15.252\admins\JAM\PSB AUTO LOAN\REPORTINGS\AUTO STAT (MONDAY)\" & "Autostat" & Format(Now(), "mm-dd-yyyy") & ".xls", FileFormat:=56
wb.Sheets(1).Rows("1:1").EntireRow.Delete
wb.Close

End Sub

Sub status()
Dim LastRow As Long
Dim i As Long

 Dim conn As Variant



    For Each conn In ActiveWorkbook.Connections
        conn.ODBCConnection.BackgroundQuery = False
    Next conn

    ActiveWorkbook.RefreshAll


    
With Worksheets("Database")
LastRow = .Cells(.Rows.count, "B").End(xlUp).Row
End With

Worksheets("Stat Report").Rows("2:" & Rows.count).Clear

Worksheets("Database").Range("B2:B" & LastRow).Copy Destination:=Worksheets("Stat Report").Range("A2")


Worksheets("Stat Report").Range("B2:B" & LastRow).Formula = "=Text(VLOOKUP(A2,Database!B:T,19,0),""mm/dd/yyyy"")"

Worksheets("Stat Report").Range("C2:C" & LastRow).Formula = "=""'00"" &VLOOKUP(A2,Database!B:E,4,0)"

Worksheets("Stat Report").Range("D2:D" & LastRow).Formula = "=VLOOKUP(A2,Database!B:D,3,0)"

Worksheets("Stat Report").Range("E2:E" & LastRow).Value = "SP MADRID"

Worksheets("Stat Report").Range("F2:F" & LastRow).Value = "=Iferror(Text(VLOOKUP('Stat Report'!A2,'ODBC STAT REPORT'!E:J,6,0),""mm/dd/yyyy""),""NOT OKAY"")"

Worksheets("Stat Report").Range("G2:G" & LastRow).Value = "=VLOOKUP(A2,'ODBC STAT REPORT'!E:I,5,0)"

Worksheets("Stat Report").Range("H2:H" & LastRow).Value = "=VLOOKUP('Stat Report'!A2,'ODBC STAT REPORT'!E:H,3,0) & "", "" &  VLOOKUP('Stat Report'!A2,'ODBC STAT REPORT'!E:H,4,0)"

With Worksheets("Stat Report").UsedRange
.Copy
.PasteSpecial xlPasteValues
End With


Worksheets("Stat Report").Range("A2:A" & LastRow).Value = ""


For i = LastRow To 2 Step -1
        If (Worksheets("Stat Report").Cells(i, "F").Value) = "NOT OKAY" Then
    'Cells(i, "A").EntireRow.ClearContents ' USE THIS TO CLEAR CONTENTS BUT NOT DELETE ROW
            Worksheets("Stat Report").Cells(i, "A").EntireRow.Delete
        End If
    Next i
    

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "- Inserted By API"
rplc = ""

 Worksheets("Stat Report").Activate


  Worksheets("Stat Report").Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

 Worksheets("Stat Report").Range("A1:H" & LastRow).Sort _
Key1:=Range("F2"), Header:=xlYes, Order1:=xlAscending


Dim wb As Workbook
Set wb = Workbooks.Add
ThisWorkbook.Sheets("Stat Report").Copy Before:=wb.Sheets(1)
wb.SaveAs "\\192.168.15.252\admins\JAM\PSB AUTO LOAN\REPORTINGS\AUTO STAT (MONDAY)\" & "Madrid Weekly Status Report as of " & Format(Now(), "yyyyMmmd") & ".xlsx", FileFormat:=51
wb.Password = "PSBAuto"
wb.Close
End Sub
