Attribute VB_Name = "Clean_Excel"
Sub ClearExcessRowsAndColumns()
       
Dim myLastRow As Long
Dim myLastCol As Long
Dim wks As Worksheet
Dim dummyRng As Range
For Each wks In ActiveWorkbook.Worksheets
  With wks
    myLastRow = 0
    myLastCol = 0
    Set dummyRng = .UsedRange
    On Error Resume Next
    myLastRow = _
      .Cells.Find("*", after:=.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlWhole, _
        searchdirection:=xlPrevious, _
        SearchOrder:=xlByRows).Row
    myLastCol = _
      .Cells.Find("*", after:=.Cells(1), _
        LookIn:=xlFormulas, LookAt:=xlWhole, _
        searchdirection:=xlPrevious, _
        SearchOrder:=xlByColumns).Column
    On Error GoTo 0
    If myLastRow * myLastCol = 0 Then
        .Columns.Delete
    Else
        .Range(.Cells(myLastRow + 1, 1), _
          .Cells(.Rows.count, 1)).EntireRow.Delete
        .Range(.Cells(1, myLastCol + 1), _
          .Cells(1, .Columns.count)).EntireColumn.Delete
    End If
  End With
Next wks
     
End Sub

Sub count()


MsgBox (ActiveSheet.UsedRange.Rows.count & " |" & ActiveSheet.UsedRange.Columns.count)
End Sub


