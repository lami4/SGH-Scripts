#Возникшие проблемы
#1) Колонтитул "Особый колонтитул для первой страницы" Вкл.\Выкл
#2) Грамматические ошибки

#1) DONE. Выяснить почему не добавляется строка в массив (апдейт - последняя строка в таблице не добавляется. Фор луп работает некорректно.)
#2) DONE. Добавить фильтр для файлов и добавлять их в отдельный массив.
#3) DONE. Сделать проверку на существование файла.
#4) DONE. Царский парсинг значений из ворда.
#5) DONE. Выяснить как забирать значения из ячеек и очищать их от мусора (тэги, перенос строки, символ конца строки и т.п.)
#6) DONE. Написать функцию сравнения
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
        If ($Show -eq "OK") {
        Return $objForm.SelectedPath
        } Else {
        Exit
        }
}

Function Get-DataFromSPC ($selectedFolder, $currentSPCName)
{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open("$selectedFolder\$currentSPCName")
    [int]$rowCount = $document.Tables.Item(1).Rows.Count + 1
        for ($i = 2; $i -lt $rowCount; $i++) {
        #gets document title, document name and notification number from SPC
        [string]$valueInDocumentTitleCell = (((($document.Tables.Item(1).Cell($i,2).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
        [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,1).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
        [string]$script:valueInNotificationNoCell = ((($document.Sections.Item(1).Footers.Item(1).Range.Tables.Item(1).Cell(2, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ') 
            #if document name and document title cells are filled out (cell contains more than 0 characters), puts their value in corresponding arrays.
            #Otherwise, checks if document name cell is filled out - if it is, puts a value to array for files.
            if ($valueInDocumentTitleCell.Length -gt 0 -and $valueInDocumentNameCell.Length -gt 0) {
            $script:documentTitles += $valueInDocumentTitleCell
            $script:documentNames += $valueInDocumentNameCell
            } else {
                if ($valueInDocumentTitleCell.Length -eq 0 -and $valueInDocumentNameCell.Length -gt 0) {
                $script:fileNames += $valueInDocumentNameCell
                } else {
                Write-Host "Cells are empty"
                }
            }
        }
    $document.Close()
    $word.Quit()
}

Function Compare-Strings ($SPCvalue, $valueFromDocument, $message) 
{
    if ($valueFromDocument -eq $SPCvalue) {
    Write-Host "Hit for $message"
    } else {
    Write-Host "No hit for $message"
    }
}

Function Compare-DataFromSPCAgainstDocuments ($selectedFolder, $dataFromSPC) 
{
    for ($i = 0; $i -lt $dataFromSPC[0].Length; $i++) {
        Start-Sleep -Seconds 2
        #checks if the document exist
        $currentDocumentBaseName = $dataFromSPC[0][$i]
        $documentExistence = Test-Path -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            if ($documentExistence -eq $true) {
            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            Write-Host "$currentDocumentFullName exists"
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $document = $word.Documents.Open("$currentDocumentFullName")
            $valueForName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(5, 8).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            $valueForTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(8, 8).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
            $valueForNotificationNo = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(6, 5).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            Compare-Strings -SPCvalue $dataFromSPC[0][$i] -valueFromDocument $valueForName -message "name"
            Compare-Strings -SPCvalue $dataFromSPC[1][$i] -valueFromDocument $valueForTitle -message "title"
            Compare-Strings -SPCvalue $script:valueInNotificationNoCell -valueFromDocument $valueForNotificationNo -message "notification no."
            $document.Close()
            $word.Quit()
            } else {
            Write-Host "$currentDocumentFullName does not exist"
            }
    }
}


#Script code
$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно проверить входимость."
Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Get-DataFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name
$SPCdata = $script:documentNames, $script:documentTitles
Compare-DataFromSPCAgainstDocuments -selectedFolder $pathToFolder -dataFromSPC $SPCdata
#clears variables and arrays
Clear-Variable -Name "valueInNotificationNoCell" -Scope Script
$script:documentTitles = @()
$script:documentNames = @()
$script:fileNames = @()
}
