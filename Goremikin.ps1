#Возникшие проблемы
#1) Колонтитул "Особый колонтитул для первой страницы" Вкл.\Выкл
#2) Грамматические ошибки

#1) DONE. Выяснить почему не добавляется строка в массив (апдейт - последняя строка в таблице не добавляется. Фор луп работает некорректно.)
#2) DONE. Добавить фильтр для файлов и добавлять их в отдельный массив.
#3) DONE. Сделать проверку на существование файла.
#4) DONE. Царский парсинг значений из ворда.
#5) DONE. Выяснить как забирать значения из ячеек и очищать их от мусора (тэги, перенос строки, символ конца строки и т.п.)
#6) DONE. Написать функцию сравнения
#7) DONE. HTML статистика
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

Function Compare-Strings ($SPCvalue, $valueFromDocument, $message, $positive, $negative) 
{
    if ($valueFromDocument -eq $SPCvalue) {
    Write-Host "Hit for $message"

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green""><b>$positive</b></font></td>" -Encoding UTF8
#========Statistics========

    } else {
    Write-Host "No hit for $message"

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""red""><b>$negative</b></font></td>" -Encoding UTF8
#========Statistics========

    }
}

Function Compare-DataFromSPCAgainstDocuments ($selectedFolder, $dataFromSPC) 
{
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<tr>
<th>Документ</th>
<th>*.doc/*.docx</th> 
<th>Обозначение</th>
<th>Название</th>
<th>Номер извещения</th>
</tr>" -Encoding UTF8
#========Statistics========

    for ($i = 0; $i -lt $dataFromSPC[0].Length; $i++) {
    $currentDocumentBaseName = $dataFromSPC[0][$i]

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<tr>
<td>$currentDocumentBaseName</td>" -Encoding UTF8
#========Statistics========

        Start-Sleep -Seconds 2
        #checks if the document exist
        $documentExistence = Test-Path -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            if ($documentExistence -eq $true) {

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green""><b>Найден</b></font></td>" -Encoding UTF8
#========Statistics========

            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            Write-Host "$currentDocumentFullName exists"
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $document = $word.Documents.Open("$currentDocumentFullName")
            $valueForName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(5, 8).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            $valueForTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(8, 8).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
            $valueForNotificationNo = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(6, 5).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            Compare-Strings -SPCvalue $dataFromSPC[0][$i] -valueFromDocument $valueForName -message "name" -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $dataFromSPC[1][$i] -valueFromDocument $valueForTitle -message "title" -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $script:valueInNotificationNoCell -valueFromDocument $valueForNotificationNo -message "notification no." -positive "Совпадает" -negative "Не совпадает"
            $document.Close()
            $word.Quit()
            Add-Content "$PSScriptRoot\Test Report.html" "</tr>" -Encoding UTF8
            } else {
            Write-Host "$currentDocumentFullName does not exist"
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "
<td><font color=""red""><b>Не найден</b></font></td>
<td>---</td>
<td>---</td>
<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
    }
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</table>
<br>" -Encoding UTF8
#========Statistics========
}

#Script code
$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно проверить входимость."

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>LiveDoc Report</title>
<style type=""text/css"">
   div {
    font-family: Verdana, Arial, Helvetica, sans-serif;
   }
table {
    border-collapse: collapse;
}
table, td, th {
    border: 1px solid black;
    padding: 3px;
}
td {
    text-align:center;
    background-color: #FFC;
}
</style>
</head>
<body>
<div>
<h3>Hello.</h3>" -Encoding UTF8
#========Statistics========

Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Start-Sleep -Seconds 2
Get-DataFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name

#========Statistics========
$curSpc = $_.Name
Add-Content "$PSScriptRoot\Test Report.html" "
<table style=""width:100%"">
<tr>
<td colspan=""5"">$curSpc</td>
</tr>" -Encoding UTF8
#========Statistics========

$SPCdata = $script:documentNames, $script:documentTitles
Compare-DataFromSPCAgainstDocuments -selectedFolder $pathToFolder -dataFromSPC $SPCdata
#clears variables and arrays
Clear-Variable -Name "valueInNotificationNoCell" -Scope Script
$script:documentTitles = @()
$script:documentNames = @()
$script:fileNames = @()
}

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "
</div>
</body>
</html>" -Encoding UTF8
#========Statistics========
