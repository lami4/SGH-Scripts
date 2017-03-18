Sub Test()
Dim Column As Range
Dim RowsCount As Integer
RowsCount = ActiveDocument.Tables(1).Rows.Count
MsgBox (RowsCount)
Set Column = ActiveDocument.Range(Start:=ActiveDocument.Tables(1).Cell(1, 1).Range.Start, _
End:=ActiveDocument.Tables(1).Cell(RowsCount, 1).Range.End)
End Sub

$application = New-Object -ComObject word.application
$application.Visible = $true
$document = $application.documents.open("C:\Users\Светлана\Downloads\Спецификация_final.docx")
$table = $document.Tables.Item(1)
$column = $document.Range([ref]$table.Cell(1, 1).Range.Start, [ref]$table.Cell(22, 1).Range.End)
$column.Select()
$column.Bold = $true
