Sub Gluer()
LastPopulated = ActiveWorkbook.Worksheets("Substrings").Cells(ActiveWorkbook.Worksheets("Substrings").Rows.Count, "A").End(xlUp).Row
MsgBox (LastPopulated)
Dim ID As String
Dim ReplaceWith As String
For i = 2 To LastPopulated
    ID = ActiveWorkbook.Worksheets("Substrings").Cells(i, 1)
    ReplaceWith = ActiveWorkbook.Worksheets("Substrings").Cells(i, 3)
    StringLength = Len(ReplaceWith)
    If StringLength >= 255 Then
    RowNumber = ActiveWorkbook.Worksheets("BrokenSource").Columns("C").Find(ID).Row
    TemplateCell = ActiveWorkbook.Worksheets("BrokenSource").Cells(RowNumber, 3)
    TemplateCell = Replace(TemplateCell, ID, ReplaceWith, 1, 1)
    ActiveWorkbook.Worksheets("BrokenSource").Cells(RowNumber, 3) = TemplateCell
    Else
    ActiveWorkbook.Worksheets("BrokenSource").Columns("C").Replace What:=ID, Replacement:=ReplaceWith, LookAt:=xlPart
    End If
Next i
End Sub
