#Global arrays and variables
$script:documentTitles = @()
$script:documentNames = @()

#Functions
Function Select-Folder ($description)
{
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
$objForm = New-Object System.Windows.Forms.FolderBrowserDialog
$objForm.Rootfolder = "Desktop"
$objForm.Description = $description
$Show = $objForm.ShowDialog()
If ($Show -eq "OK")
    {
    Return $objForm.SelectedPath
    } Else {
    Exit
    }
}

Function Get-InformationFromSPC ($selectedFolder, $currentSPCName)
{
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$document = $word.Documents.Open("$selectedFolder\$currentSPCName")
[int]$rowCount = $document.Tables.Item(1).Rows.Count
for ($i = 2; $i -lt $rowCount; $i++) {
#takes value from SPC for document title
[string]$valueInDocumentTitleCell = $document.Tables.Item(1).Cell($i,2).Range.Text
$parsedDocumentTitleValue = $valueInDocumentTitleCell -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е' -replace '\s', ' '
    if ($parsedDocumentTitleValue.Length -le 2) {
    Write-Host "Empty Cell"
    } else {
    $script:documentTitles += $parsedDocumentTitleValue.ToLower()
    }
#takes value from SPC for document name
[string]$valueInDocumentNameCell = $document.Tables.Item(1).Cell($i,1).Range.Text
$parsedDocumentNameValue = $valueInDocumentNameCell -replace '\s+', ' '
    if ($parsedDocumentNameValue.Length -le 2) {
    Write-Host $parsedDocumentNameValue.Length
    Write-Host "Empty Cell"
    } else {
    $script:documentNames += $valueInDocumentNameCell
    }
#takes value from SPC for notification number
[string]$script:valueInNotificationNoCell = $document.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1).Cell(2, 3).Range.Text
}
$document.Close()
$word.Quit()
}

#Script code
$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно проверить входимость."
Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Get-InformationFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name
Write-Host $script:valueInNotificationNoCell
$SPCdata = $script:documentTitles, $script:documentNames
Write-Host $SPCdata

$script:documentTitles = @()
$script:documentNames = @()
}
