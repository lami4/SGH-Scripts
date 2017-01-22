#Возникшие проблемы
#1) Колонтитул "Особый колонтитул для первой страницы" Вкл.\Выкл?
#2) Грамматические ошибки?
#3) Как различать документ и файлы?

#1) DONE. Выяснить почему не добавляется строка в массив (апдейт - последняя строка в таблице не добавляется. Фор луп работает некорректно.)
#2) DONE. Добавить фильтр для файлов и добавлять их в отдельный массив.
#3) DONE. Сделать проверку на существование файла.
#4) DONE. Царский парсинг значений из ворда.
#5) DONE. Выяснить как забирать значения из ячеек и очищать их от мусора (тэги, перенос строки, символ конца строки и т.п.)
#6) DONE. Написать функцию сравнения
#7) DONE. HTML статистика
#8) DONE. Просмотр значений из специйикации и документа
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
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green"" id=""text""><b>$positive</b></font>
<div id=""hide"">
<table>
<tr>
<td id=""indication"">В спецификации:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">В документе:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
#========Statistics========

    } else {
    Write-Host "No hit for $message"

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""red"" id=""text""><b>$negative</b></font>
<div id=""hide"">
<table>
<tr>
<td id=""indication"">В спецификации:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">В документе:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
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
<br>
<hr>" -Encoding UTF8
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
th {
padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #bfbfbf;
}
td {
	padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #FFC;
}
#tableHeader {
background-color: white;
text-align: left;
border: none;
padding: 0px;
}
hr {
	border-top: 1px solid #8c8b8b;
	border-bottom: 1px solid #fff;
    width: 80%;
}
#hide {
    display: none;
	position: absolute;
	background-color: white;
	text-align: left;
}
#text:hover + #hide {
display: block;
}
#indication {
text-align: left;
border: 0px;
background-color: #bfbfbf;
}
</style>
</head>
<body>
<div>
<h3>Результаты сравнения</h3>" -Encoding UTF8
#========Statistics========

Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Start-Sleep -Seconds 2
Get-DataFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name

#========Statistics========
$curSpc = $_.Name
Add-Content "$PSScriptRoot\Test Report.html" "
<table style=""width:90%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
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
Invoke-Item "$PSScriptRoot\Test Report.html"


ONCLICK INFO
#Возникшие проблемы
#1) Колонтитул "Особый колонтитул для первой страницы" Вкл.\Выкл
#2) Грамматические ошибки
#3) Как различать документ и файлы?

#1) DONE. Выяснить почему не добавляется строка в массив (апдейт - последняя строка в таблице не добавляется. Фор луп работает некорректно.)
#2) DONE. Добавить фильтр для файлов и добавлять их в отдельный массив.
#3) DONE. Сделать проверку на существование файла.
#4) DONE. Царский парсинг значений из ворда.
#5) DONE. Выяснить как забирать значения из ячеек и очищать их от мусора (тэги, перенос строки, символ конца строки и т.п.)
#6) DONE. Написать функцию сравнения
#7) DONE. HTML статистика
#8) DONE. Просмотр значений из специйикации и документа
clear
#Global arrays and variables
$script:documentTitles = @()
$script:documentNames = @()
$script:fileNames = @()
$script:JSvariable = 1

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
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>$positive</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
<table>
<tr>
<td id=""indication"">Спецификация:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">Документ:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========

    } else {
    Write-Host "No hit for $message"

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>$negative</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
<table>
<tr>
<td id=""indication"">Спецификация:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">Документ:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
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
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</tr>" -Encoding UTF8
#========Statistics========
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
<br>
<hr>" -Encoding UTF8
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
th {
padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #bfbfbf;
}
td {
	padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #FFC;
}
#tableHeader {
background-color: white;
text-align: left;
border: none;
padding: 0px;
}
hr {
	border-top: 1px solid #8c8b8b;
	border-bottom: 1px solid #fff;
    width: 80%;
}
.hide {
    display: none;
	position: absolute;
	background-color: white;
	text-align: left;
	border: solid 1px black;
}
#indication {
text-align: left;
border: 0px;
background-color: #bfbfbf;
}
</style>
<script>
function my_f(objName) {
var object = document.getElementById(objName);
object.style.display == 'block' ? object.style.display = 'none' : object.style.display = 'block'
}
</script>
</head>
<body>
<div>
<h3>Результаты сравнения</h3>" -Encoding UTF8
#========Statistics========

Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
Start-Sleep -Seconds 2
Get-DataFromSPC -selectedFolder $pathToFolder -currentSPCName $_.Name

#========Statistics========
$curSpc = $_.Name
Add-Content "$PSScriptRoot\Test Report.html" "
<table style=""width:80%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
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
Invoke-Item "$PSScriptRoot\Test Report.html"
Release version ++++++++++++
clear
#Script arrays and variables
$script:JSvariable = 0

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

Function Compare-Strings ($SPCvalue, $valueFromDocument, $message, $positive, $negative) 
{
    if ($valueFromDocument -eq $SPCvalue) {
    Write-Host "Hit for $message"
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>$positive</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
<table>
<tr>
<td id=""indication"">Спецификация:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">Документ:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
    } else {
    Write-Host "No hit for $message"
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>$negative</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
<table>
<tr>
<td id=""indication"">Спецификация:</td>
<td id=""indication"">$SPCvalue</td>
</tr>
<tr>
<td id=""indication"">Документ:</td>
<td id=""indication"">$valueFromDocument</td>
</tr>
</table>
</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
    }
}

Function Get-DataFromSpecification ($selectedFolder, $currentSPCName) {
    $documentNames = @()
    $documentTitles = @()
    $fileNames = @()
    $fileMd5s = @()
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open("$selectedFolder\$currentSPCName")
    [int]$rowCount = $document.Tables.Item(1).Rows.Count + 1
    for ($i = 1; $i -lt $rowCount; $i++) {
        [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
        if ($valueInDocumentNameCell.length -ne 0) {
        if ($valueInDocumentNameCell -match '\b([A-Z]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d([^\s]*)') {
            [string]$valueInDocumentTitleCell = (((($document.Tables.Item(1).Cell($i,5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
            $documentNames += $valueInDocumentNameCell
            $documentTitles += $valueInDocumentTitleCell
            } else {
            [string]$valueInFileMd5Cell = (((($document.Tables.Item(1).Cell($i,7).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')).ToLower()
            if ($valueInFileMd5Cell -match '([m,M]\s*[d,D]\s*5)\s*:') {
            [string]$valueInFileNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            $fileMd5s += $valueInFileMd5Cell
            $fileNames += $valueInFileNameCell
            }
            }
        }
        }
    $document.Close()
    $documentData = $documentNames, $documentTitles
    $fileData = $fileNames, $fileMd5s
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<tr>
<th>Название документа/Файла</th>
<th>Документ/Файл</th> 
<th>Обозначение</th>
<th>Наименование</th>
<th>MD5</th>
</tr>" -Encoding UTF8
#========Statistics========
    for ($i = 0; $i -lt $documentData[0].Length; $i++) {
    $currentDocumentBaseName = $documentData[0][$i]
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<tr>
<td>$currentDocumentBaseName</td>" -Encoding UTF8
#========Statistics========
    $documentExistence = Test-Path -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
        if ($documentExistence -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green""><b>Найден</b></font></td>" -Encoding UTF8
#========Statistics========
            if ($currentDocumentBaseName -match 'SPC') {
            #FOR SPC
            Write-Host "***** PROCESSING SPECIFICAGTION ******"
            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            Write-Host "$currentDocumentFullName найден."
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            Write-Host $valueForDocTitle
            Write-Host $valueForDocName
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "document name" -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "document title" -positive "Совпадает" -negative "Не совпадает"
            $document.Close()
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            } else {
            #FOR REST
            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            Write-Host "$currentDocumentFullName найден."
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace '\s+', ' ' -replace 'ё', 'е').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
            Write-Host $valueForDocTitle
            Write-Host $valueForDocName
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "document name"  -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "document title"  -positive "Совпадает" -negative "Не совпадает"
            $document.Close()
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
        } else {
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
    for ($i = 0; $i -lt $fileData[0].Length; $i++) {
        $currentFileName = $fileData[0][$i]
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<tr>
<td>$currentFileName</td>" -Encoding UTF8
#========Statistics========
        $fileExistence = Test-Path -Path "$selectedFolder\$currentFileName"
        if ($fileExistence -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "<td><font color=""green""><b>Найден</b></font></td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========
            Write-Host "File found"
            $fileHash = Get-FileHash -Path "$selectedFolder\$currentFileName" -Algorithm MD5
            $fileHashFromSpc = $fileData[1][$i] -split (":")
            Compare-Strings -SPCvalue $fileHashFromSpc[1].Trim(' ') -valueFromDocument $fileHash.Hash.ToLower() -message "md5" -positive "Совпадает" -negative "Не совпадает"
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</tr>" -Encoding UTF8
#========Statistics========
           # if ($fileHashFromSpc[1].Trim(' ') -eq $fileHash.Hash.ToLower()) {
           # Write-Host "Hash sum matches"
            #} else {
            #Write-Host "Hash sum does not match"
           #}
        } else {
        Write-Host "File not found"
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
    Write-Host "-------------------------------"
    $word.Quit()
    Write-Host $documentNames
    Write-Host $documentTitles
    Write-Host $fileNames
    Write-Host $fileMd5s
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</table>
<br>
<hr>" -Encoding UTF8
#========Statistics========
}


#script codee
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
th {
padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #bfbfbf;
}
td {
	padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #FFC;
}
#tableHeader {
background-color: white;
text-align: left;
border: none;
padding: 0px;
}
hr {
	border-top: 1px solid #8c8b8b;
	border-bottom: 1px solid #fff;
    width: 80%;
}
.hide {
    display: none;
	position: absolute;
	background-color: white;
	text-align: left;
	border: solid 1px black;
}
#indication {
text-align: left;
border: 0px;
background-color: #bfbfbf;
}
</style>
<script>
function my_f(objName) {
var object = document.getElementById(objName);
object.style.display == 'block' ? object.style.display = 'none' : object.style.display = 'block'
}
</script>
</head>
<body>
<div>
<h3>Результаты сравнения</h3>" -Encoding UTF8
#========Statistics========
Measure-Command {
Get-ChildItem "$pathToFolder\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
#========Statistics========
$curSpc = $_.Name
Add-Content "$PSScriptRoot\Test Report.html" "
<table style=""width:80%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
</tr>" -Encoding UTF8
#========Statistics========
Get-DataFromSpecification -selectedFolder $pathToFolder -currentSPCName $_.Name
}
}
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "
</div>
</body>
</html>" -Encoding UTF8
#========Statistics========
Invoke-Item "$PSScriptRoot\Test Report.html"
