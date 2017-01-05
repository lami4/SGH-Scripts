#1) DONE. Выяснить почему не добавляется строка в массив (апдейт - последняя строка в таблице не добавляется. Фор луп работает некорректно.)
#2) DONE. Добавить фильтр для файлов и добавлять их в отдельный массив
#3) DONE. Сделать проверку на существование файла
#4) Царский парсинг значений из ворда.
#5) Написать функцию сравнения сравнения
clear
#Global arrays and variables
$script:documentTitles = @()
$script:documentNames = @()
$script:fileNames = @()

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
[int]$rowCount = $document.Tables.Item(1).Rows.Count + 1
for ($i = 2; $i -lt $rowCount; $i++) {
#gets document title, document name and notification number from SPC
[string]$valueInDocumentTitleCell = $document.Tables.Item(1).Cell($i,2).Range.Text
[string]$valueInDocumentNameCell = $document.Tables.Item(1).Cell($i,1).Range.Text
[string]$script:valueInNotificationNoCell = ((($document.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1).Cell(2, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ') 
#checks if cells are empty
$parsedDocumentTitleValue = $valueInDocumentTitleCell -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е'
$parsedDocumentNameValue = $valueInDocumentNameCell -replace '\s+', ' '
#if document name and document title are filled out, puts their value in corresponding arrays
    if ($parsedDocumentTitleValue.Length -gt 2 -and $parsedDocumentNameValue.Length -gt 2) {
    $script:documentTitles += $parsedDocumentTitleValue.ToLower().Trim(' ')
    $script:documentNames += $parsedDocumentNameValue.Trim([char]0x0007, ' ')
    } else {
        if ($parsedDocumentTitleValue.Length -le 2 -and $parsedDocumentNameValue.Length -gt 2) {
        $script:fileNames += $parsedDocumentNameValue
        } else {
        Write-Host "Cells are empty"
        }
    }
}
$document.Close()
$word.Quit()
}

#Script code
$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно проверить входимость."
Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Get-InformationFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name
$SPCdata = $script:documentNames, $script:documentTitles
Write-Host $script:valueInNotificationNoCell
Write-Host $script:valueInNotificationNoCell.Length
for ($i = 0; $i -lt $SPCdata[0].Length; $i++) {
#checks if the document exist
$currentDocumentName = $SPCdata[0][$i]
$documentExistence = Test-Path -Path "$pathToFolder\$currentDocumentName.*" -Exclude "*.pdf"
    if ($documentExistence -eq $true) {
    Write-Host "Document exists"
    } else {
    Write-Host "Document does not exist"
    }
}
#clears variables and arrays
Clear-Variable -Name "valueInNotificationNoCell" -Scope Script
$script:documentTitles = @()
$script:documentNames = @()
$script:fileNames = @()
}
