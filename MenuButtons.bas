Attribute VB_Name = "MenuButtons"
Sub copycolumns()
Dim lngLastRow As Long 'declare a variable for the last row
Dim lngLastRowAddinfo As Long
Dim sCellVal As String
Dim count As Integer
Dim LastCol As Integer
Dim agentcount As Integer
Dim agent As String


agentcount = Worksheets("Menu").Range("D" & Rows.count).End(xlUp).Row - 6
  lngLastRow = Worksheets("New Endo").Range("B" & Rows.count).End(xlUp).Row
  lngLastRowAddinfo = Worksheets("Database").Range("D" & Rows.count).End(xlUp).Row
  
  'filter
  If Worksheets("Database").AutoFilterMode = True Then
Worksheets("Database").Range("A1").AutoFilter
Worksheets("Database").Range("A1").AutoFilter
Else
Worksheets("Database").Range("A1").AutoFilter
End If
  
  
  Worksheets("Database").Range("D" & lngLastRowAddinfo + 1 & ":S" & lngLastRow + lngLastRowAddinfo - 1).Formula = "=IFERROR(INDEX('New Endo'!$A2:$DA2,MATCH(Database!D$1,'New Endo'!$A$1:$DA$1,0)),"""")"

With Worksheets("Database").Range("D2:AB" & lngLastRow + lngLastRowAddinfo - 1)
    .Value = .Value
End With

  Worksheets("Database").Range("C2:C" & lngLastRow + lngLastRowAddinfo - 1).Copy

With Worksheets("Database")
  .Range("C2").PasteSpecial xlPasteValues
End With




With Worksheets("New Endo")
        LastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
    End With


For i = 1 To LastCol

sCellVal = Worksheets("New Endo").Cells(1, i).Value
    
If sCellVal Like "*TELEPHONE*" Or _
    sCellVal Like "Telephone*" Then
    
   Worksheets("New Endo").Cells(1, i).ClearContents

   Worksheets("New Endo").Range(Cells(2, i), Cells(lngLastRow, i)).Copy

   Worksheets("Database").Range("Q" & lngLastRowAddinfo + 1).PasteSpecial xlPasteValues
End If

If sCellVal Like "*DESCRIPTION*" Or _
    sCellVal Like "Description*" Then
    
   Worksheets("New Endo").Cells(1, i).ClearContents

   Worksheets("New Endo").Range(Cells(2, i), Cells(lngLastRow, i)).Copy

   Worksheets("Database").Range("I" & lngLastRowAddinfo + 1).PasteSpecial xlPasteValues
End If

If sCellVal Like "*COURT*" Or _
    sCellVal Like "Court*" Then
    
   Worksheets("New Endo").Cells(1, i).ClearContents

   Worksheets("New Endo").Range(Cells(2, i), Cells(lngLastRow, i)).Copy

   Worksheets("Database").Range("L" & lngLastRowAddinfo + 1).PasteSpecial xlPasteValues
End If

Next i

Worksheets("Database").Range("T" & lngLastRowAddinfo + 1 & ":T" & (lngLastRow - 1) + lngLastRowAddinfo).Formula = "=Today()"
  Worksheets("Database").Range("T2:T" & lngLastRow + lngLastRowAddinfo - 1).Copy

With Worksheets("Database")
  .Range("T2").PasteSpecial xlPasteValues
End With


Worksheets("Database").Columns("A").NumberFormat = "0"
Worksheets("Database").Columns("E").NumberFormat = "0"



Worksheets("Database").Range("C1").AutoFilter field:=3, Criteria1:=""

'sort
Worksheets("Database").AutoFilter.Sort.SortFields.Clear
Worksheets("Database").AutoFilter.Sort.SortFields.Add Key:= _
        Range("F1:F" & lngLastRow + lngLastRowAddinfo), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With Worksheets("Database").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'add agent count if additonal agent is needed
For i = lngLastRowAddinfo + 1 To lngLastRowAddinfo + lngLastRow - 1
If agentcount = 1 Then
agentcount = agentcount + 1
ElseIf agentcount = 2 Then
agentcount = agentcount + 1
Else
agentcount = 1
End If


Worksheets("Database").Range("C" & i).Value = Worksheets("Menu").Range("D" & agentcount + 6).Value
Next i

  'filter
  If Worksheets("Database").AutoFilterMode = True Then
Worksheets("Database").Range("A1").AutoFilter
Worksheets("Database").Range("A1").AutoFilter
Else
Worksheets("Database").Range("A1").AutoFilter
End If
  Worksheets("Database").Activate

End Sub

Sub SearchPhone()

Dim sCellVal As String
Dim count As Integer
Dim LastCol As Integer

With Worksheets("New Endo")
        LastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
    End With


For i = 1 To LastCol

sCellVal = Worksheets("New Endo").Cells(1, LastCol).Value
    
If sCellVal Like "*TELEPHONE*" Or _
    sCellVal Like "Telephone*" Then
    
   Worksheets("New Endo").Cells(1, LastCol).ClearContents

End If

Next i
End Sub

Sub RemoveChar()
Dim rng As Range
Dim WorkRng As Range
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each rng In WorkRng
    xOut = ""
    For i = 1 To Len(rng.Value)
        xTemp = Mid(rng.Value, i, 1)
        If xTemp Like "[0-9]" Then
            xStr = xTemp
        Else
            xStr = ""
        End If
        xOut = xOut & xStr
    Next i
    rng.Value = xOut
Next
End Sub


Sub ChCodegen()
Dim lastrowdatabase As Long
Dim lastch As Long



lastrowdatabase = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row

lastch = Worksheets("Ch Code Used").Range("A" & Rows.count).End(xlUp).Row

Worksheets("Ch Code Used").Rows("2:" & Rows.count).EntireRow.Delete


Worksheets("Database").Range("B2:B" & lastrowdatabase).Copy

With Worksheets("Ch Code Used")
  .Range("A" & lastch + 1).PasteSpecial xlPasteValues
End With

lastch = Worksheets("Ch Code Used").Range("A" & Rows.count).End(xlUp).Row

Worksheets("Ch Code Used").Range("B2:D" & lastch + 1).Clear

Worksheets("Ch Code Used").Range("A1:D" & lastch + lastrowdatabase).RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes

lastch = Worksheets("Ch Code Used").Range("A" & Rows.count).End(xlUp).Row

Worksheets("Ch Code used").Range("B2:B" & lastch).Formula = "=MID(A2,FIND(""A"",A2)+1,FIND(""-"",A2)-FIND(""A"",A2)-1)"
Worksheets("Ch Code used").Range("C2:C" & lastch).Formula = "=RIGHT(A2,LEN(A2) - FIND(""-"",A2,1))"
Worksheets("Ch Code used").Range("I1").Formula = "=Text(Today(),""yymm"")"




With Worksheets("Ch Code used").UsedRange
.Value = .Value
End With
Worksheets("Ch Code used").Range("H2").Value = "Last Ch Code"
Worksheets("Ch Code used").Range("I2").FormulaArray = "=MAX((B3:B9000=I1)*C3:C9000)"


With Worksheets("Ch Code used").UsedRange
.Value = .Value
End With

Worksheets("Menu").Range("F7").Value = Worksheets("Ch Code used").Range("I2").Value

  ThisWorkbook.Worksheets("Database").Cells.WrapText = False

End Sub

Sub Chcodenumber()
Dim lastchcode As Long
Dim lastdata As Long
Dim i As Integer
Dim ch As Integer

lastchcode = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row + 1
lastdata = Worksheets("Database").Range("C" & Rows.count).End(xlUp).Row

Worksheets("Ch Code Used").Rows("2:2").EntireRow.Delete

Worksheets("Ch Code used").Range("H2").Value = "Last Ch Code"
Worksheets("Ch Code used").Range("I2").FormulaArray = "=MAX((B3:B9000=I1)*C3:C9000)"


With Worksheets("Ch Code used").UsedRange
.Value = .Value
End With


Worksheets("Menu").Range("H7").Value = lastchcode


ch = Worksheets("Menu").Range("F7").Value

For i = lastchcode To lastdata

ch = ch + 1

Worksheets("Database").Range("B" & i).Value = "01PA" & Format(Now(), "yymm") & "-" & ch

Worksheets("Database").Range("A" & i).NumberFormat = "@"
Worksheets("Database").Range("A" & i).Value = "00" & Worksheets("Database").Range("E" & i).Value
Worksheets("Database").Range("E" & i).Value = Worksheets("Database").Range("A" & i).Value
Next i
lastchcode = Worksheets("Database").Range("B" & Rows.count).End(xlUp).Row + 1
Worksheets("Database").Range("E2:E" & lastchcode).NumberFormat = "@"
Worksheets("Database").Range("E2:E" & lastchcode).Value = Worksheets("Database").Range("A2:A" & lastchcode).Value

End Sub

