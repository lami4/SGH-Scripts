$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Open("C:\Users\Светлана\Desktop\test\13.docx")
[string]$valueInCell = $document.Tables.Item(1).Cell(4,2).Range.Text
$newvalue = $valueInCell -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е'
Write-Host $newvalue.ToLower()
$document.Close()
$word.Quit()

#Newest
$documentTitles = @()
$documentNames = @()
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Open("C:\Users\Светлана\Desktop\test\13.docx")
[int]$rowCount = $document.Tables.Item(1).Rows.Count
for ($i = 0; $i -lt $rowCount; $i++) {
#takes value for document title
[string]$valueInDocumentTitleCell = $document.Tables.Item(1).Cell($i,2).Range.Text
$parsedDocumentTitleValue = $valueInDocumentTitleCell -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е'
    if ($valueInDocumentTitleCell.Length -le 2) {
    Write-Host "Empty Cell"
    } else {
    $documentTitles += $parsedDocumentTitleValue.ToLower()
    }
#takes value for document title
[string]$valueInDocumentNameCell = $document.Tables.Item(1).Cell($i,1).Range.Text
$parsedDocumentNameValue = $valueInDocumentNameCell -replace '\s+', ' '
    if ($parsedDocumentNameValue.Length -le 2) {
    Write-Host "Empty Cell"
    } else {
    $documentNames += $valueInDocumentNameCell
    }
}
$document.Close()
$word.Quit()
Write-Host $documentTitles
Write-Host $documentNames

#get under value
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Open("C:\Users\Светлана\Desktop\test\2.docx")
$underValue = $document.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1).Cell(5, 8).Range.Text
Write-Host $underValue
$document.Close()
$word.Quit()
