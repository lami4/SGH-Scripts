Sub Gluer()
LastPopulated = ActiveWorkbook.Worksheets(2).Cells(ActiveWorkbook.Worksheets(2).Rows.Count, "B").End(xlUp).Row
MsgBox (LastPopulated)
For i = 1 To LastPopulated
    ActiveWorkbook.Worksheets("BrokenSource").Columns("C").Replace What:=ActiveWorkbook.Worksheets("Substrings").Cells(i, 2), Replacement:=ActiveWorkbook.Worksheets("Substrings").Cells(i, 1), LookAt:=xlPart
Next i
End Sub
