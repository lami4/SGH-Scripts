$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Open("C:\Users\Светлана\Desktop\test\13.docx")
[string]$valueInCell = $document.Tables.Item(1).Cell(4,2).Range.Text
$newvalue = $valueInCell -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е'
Write-Host $newvalue.ToLower()
$document.Close()
$word.Quit()
