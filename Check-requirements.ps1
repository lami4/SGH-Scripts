#Global variables
$script:JSvariable = 0
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

Function Check-Specification ($selectedFolder, $currentSpecification) {
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "
<tr>
<td class=""filename"">$currentSpecification</td>" -Encoding UTF8
#========Statistics========
    #function variable
    $no_match_count = 0
    $referenceToFiles = 0
    $referenceToDocs = 0
    #фильтр для excel файла
    $fileExtension = [IO.Path]::GetExtension($currentSpecification)
    if ($fileExtension -eq ".xlsx" -or $fileExtension -eq ".xls") {
    Write-Host "Excel файл. Требуется ручная проверка."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td colspan=""4"">
Excel файл. Требуется ручная проверка.
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
    } else {
        #Write-Host "$selectedFOlder\$currentSpecification"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open("$selectedFolder\$currentSpecification")
        #написать проверку на количество таблиц в спецификации
        [int]$tableCount = $document.Tables.Count
        if ($tableCount -eq 0) {
        Write-Host "$currentSpecification не содержит таблиц."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>
<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">Файл не содержит таблиц: отсутствуют данные для считывания.</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
        } elseif ($tableCount -gt 1) {
        Write-Host "$currentSpecification содержит несколько таблиц."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>
<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">Файл содержит несколько таблиц: невозможно корректно считать данные.</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
        } else {
        [int]$rowCount = try {$document.Tables.Item(1).Rows.Count + 1} catch {""}
        Write-Host "$currentSpecification : $rowCount"
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>" -Encoding UTF8
            for ($i = 1; $i -lt $rowCount; $i++) {
            [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                #добавить подсчет совпадений и вывод полученного значения в статистику (Ссылается на <количество документов>)
                if ($valueInDocumentNameCell -match '\b(.{13})\d\d\.\d\d\.\d\d\.(.{4})\.\d\d\.\d\d([^\s]*)' -or $valueInDocumentNameCell -match '[Rr][Ff]([^a-zA-Zа-яА-я\d])[Gg][Ll]' -or $valueInDocumentNameCell -match '\d\d[^a-zA-Zа-яА-я\d\s:\-_[\]]\d\d[^a-zA-Zа-яА-я\d\s:\-_[\]]\d\d[^a-zA-Zа-яА-я\d\s:\-_[\]]') {
                    $referenceToDocs += 1
                    if ($valueInDocumentNameCell -notmatch '\b([A-Z]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d([^\s]*)') {
                    Write-Host "Обозначение содержит русские буквы или недопустимые символы."
#========Statistics========
if ($no_match_count -eq 0) {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
Обозначение ""$valueInDocumentNameCell"" содержит русские буквы/недопустимые символы или не соответствует маске." -Encoding UTF8
$script:JSvariable += 1
$no_match_count =+ 1
} else {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<br>
Обозначение ""$valueInDocumentNameCell"" содержит русские буквы/недопустимые символы или не соответствует маске." -Encoding UTF8
$no_match_count =+ 1
}
#========Statistics========
                    } else {
                    Write-Host "Обозначение соответствует маске"
                    }
                } else {
                [string]$valueInMd5Cell = ((($document.Tables.Item(1).Cell($i,7).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                [string]$valueIFileNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                    if ($valueInMd5Cell -match '([m,M]\s*[d,D]\s*5)') {
                    $referenceToFiles +=1
                        if ($valueInMd5Cell -notmatch '([m,M]\s*[d,D]\s*5)\s*:') {
                        Write-Host "Ячейка с MD5 оформлена некорректно. Отсутствует разделитель."
#========Statistics========
if ($no_match_count -eq 0) {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
Ячейка с MD5 для файла $valueIFileNameCell оформлена некорректно: отсутствует разделитель "":""." -Encoding UTF8
$script:JSvariable += 1
$no_match_count =+ 1
} else {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<br>
Ячейка с MD5 для файла $valueIFileNameCell оформлена некорректно: отсутствует разделитель "":""." -Encoding UTF8
$no_match_count =+ 1
}
#========Statistics========
                        } else {
                        Write-Host "Ячейка с MD5 суммой формлена корректно."
                    }
                    #сделать еще проверку - если есть нет мд5, но есть маска самой суммы неправильно оформлена ячейка
                    #добавить подсчет файлов и вывод полученного значения в статистику (Ссылается на <количество документов>)
                    } else {
                        if ($valueInMd5Cell -match '[a-zA-Z0-9]{32}') {
                        $referenceToFiles +=1
                        Write-Host "Ячейка с MD5 для файла $valueIFileNameCell оформлена некорректно: не соответствует маске."
#========Statistics========
if ($no_match_count -eq 0) {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
Ячейка с MD5 для файла $valueIFileNameCell оформлена некорректно: не соответствует маске." -Encoding UTF8
$script:JSvariable += 1
$no_match_count =+ 1
} else {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<br>
Ячейка с MD5 для файла $valueIFileNameCell оформлена некорректно: не соответствует маске." -Encoding UTF8
$no_match_count =+ 1
}
#========Statistics========
                        }
                    }
                }
            }
#========Statistics========
if ($no_match_count -eq 0) {
Add-Content -Path "$PSScriptRoot\Test Report.html" "<font color=""green""><b>+</b></font>
</td>" -Encoding UTF8
} else {
Add-Content -Path "$PSScriptRoot\Test Report.html" "</div>
</td>" -Encoding UTF8
}
#========Statistics========
        }
        #checks values in cells
        [string]$valueInDocVersionCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text} catch {""}
        [string]$valueInNotificationNoCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 3).Range.Text} catch {""}
        [string]$valueInDocNameCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text} catch {""}
        [string]$valueInDocTitleCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text} catch {""}
            if ($valueInDocVersionCell.Length -eq 0 -or $valueInNotificationNoCell.Length -eq 0 -or $valueInDocNameCell.Length -eq 0 -or $valueInDocTitleCell.Length -eq 0) {
            Write-Host "Невозможно получить значения из штампа. Штамп либо отсутствует, либо неверно заверстан."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>
<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
Невозможно получить значения из штампа. Штамп либо отсутствует, либо неверно заверстан.
</div>
</td>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
            } else {
            Write-Host "Значения из штампа получены."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>
<font color=""green""><b>+</b></font>
</td>" -Encoding UTF8
#========Statistics========
            }
        $document.Close()
        $word.Quit()
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<td>
<b>$referenceToDocs</b>
</td>
<td>
<b>$referenceToFiles</b>
</td>
</tr>" -Encoding UTF8
#========Statistics========
    Write-Host "-------End of document-------"   
    }
}


Function Check-Rest ($selectedFolder, $currentDocument) {
    $fileExtension = [IO.Path]::GetExtension($currentDocument)
    if ($fileExtension -eq ".xlsx" -or $fileExtension -eq ".xls") {
    Write-Host "Excel файл. Требуется ручная проверка."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<tr>
<td class=""filename"">$currentDocument
</td>
<td>
Excel файл. Требуется ручная проверка.
</td>
</tr>" -Encoding UTF8
#========Statistics========
    } else {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open("$selectedFolder\$currentDocument")
    Write-Host "$currentDocument"
    #checks values in cells
    [string]$valueInDocVersionCell = try {$document.Tables.Item(1).Cell(7, 3).Range.Text} catch {""}
    [string]$valueInNotificationNoCell = try {$document.Tables.Item(1).Cell(7, 5).Range.Text} catch {""}
    [string]$valueInDocNameCell = try {$document.Tables.Item(1).Cell(6, 8).Range.Text} catch {""}
    [string]$valueInDocTitleCell = try {$document.Tables.Item(1).Cell(9, 7).Range.Text} catch {""}
            if ($valueInDocVersionCell.Length -eq 0 -or $valueInNotificationNoCell.Length -eq 0 -or $valueInDocNameCell.Length -eq 0 -or $valueInDocTitleCell.Length -eq 0) {
            Write-Host "Невозможно получить значения из штампа. Штамп либо отсутствует, либо неверно заверстан."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<tr>
<td class=""filename"">$currentDocument
</td>
<td>
<font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
<div class=""hide"" id=""div_$script:JSvariable"">
Невозможно получить значения из штампа. Штамп либо отсутствует, либо неверно заверстан.
</div>
</td>
</tr>" -Encoding UTF8
$script:JSvariable += 1
#========Statistics========
            } else {
            Write-Host "Значения из штампа получены."
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<tr>
<td class=""filename"">$currentDocument
</td>
<td>
<font color=""green""><b>+</b></font>
</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
    $document.Close()
    $word.Quit()
    Write-Host "-------End of document-------"
    }
}

$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно соответствие требованиям."
Measure-Command {
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
.filename {
    text-align: left;  
}
.hide {
    display: none;
	position: absolute;
	background-color: white;
	text-align: left;
	border: solid 1px black;
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
<h3>Проверка на соответвствие требованиям оформления документа</h3>
<ul>
<li>Файл спецификации должен содержать только одну таблицу, строки которой при необходимости переносятся на следующую страницу;</li>
<li>При указании обозначения и наименования документа, обозначение размещается в четвертом столбце таблицы, а наименование в пятом;</li>
<li>Обозначение документов в спецификации должно состоять только из заглавных/строчных латинских букв, точки ('.'), знака минус ('-') и цифр;</li>
<li>Обозначение документов в спецификации должно соответствовать маске <b>ББББББ-ББ-ББ-ЦЦ.ЦЦ.ЦЦ.бБББ.ЦЦ.ЦЦ</b>, где
<ul>
    <li>Б — заглавная латинская буква;</li>
    <li>Ц — цифра;</li>
    <li>б — строчная латинская буква.</li>
</ul>
</li>
<li>При указании контрольной суммы файла, название файла указывается в четвертом столбце таблицы, а его контрольная сумма в правом крайнем;</li>
<li>При указании контрольной суммы файла, необходимо использовать маску <b>MD5: <контрольная сумма></b>, где
<ul>
    <li>MD5: — слово-ключ, с помощью которого скрипт понимает, что в данной ячейке указана контрольная сумма. Использование знака ':' обязательно, иначе скрипт не сможет забрать значение контрольной суммы;</li>
    <li><контрольная сумма> — контрольная сумма файла, рассчитанная по алгоритму MD5. Например, 988C393310E97032890DB2BD6BD74135;</li>
</ul>
</li>
<li>При указании названия файла, оно должно обязательно иметь расширение. Например, meteo-server-6.3.0.11.<b>war</b>;</li>
<li>Во всех документах должен использоваться штамп, который позволяет считывать указанные в нем значения.</li>
</ul>
<h3>Спецификации</h3>
<table style=""width:60%"">
<tr>
<th>Документ</th>
<th>Обозначения/MD5/<br>Таблица</th> 
<th>Штамп</th>
<th>Ссылки на<br>документы</th>
<th>Ссылки на<br>файлы</th>
</tr>" -Encoding UTF8

#========Statistics========
$spcCount = Get-ChildItem -Path "$pathToFolder\*" -Include "*.doc*", "*.xls*" | Where-Object  {$_.BaseName -match 'SPC'}
    if ($spcCount.Count -eq 0) {
#========Statistics========
Add-Content -Path "$PSScriptRoot\Test Report.html" "<tr>
<td colspan=""5"">
Спецификации не найдены
</td>
</tr>"
#========Statistics========
    } else {
    Get-ChildItem -Path "$pathToFolder\*" -Include "*.doc*", "*.xls*" | Where-Object  {$_.BaseName -match 'SPC'} | % {
    Check-Specification -selectedFolder $pathToFolder -currentSpecification $_.Name
    }
    }
Add-Content "$PSScriptRoot\Test Report.html" "</table>
<h3>Остальные документы</h3>
<table style=""width:60%"">
<tr>
<th>Документ</th>
<th>Штамп</th>
</tr>" -Encoding UTF8
Get-ChildItem -Path "$pathToFolder\*" -Include "*.doc*", "*.xls*" | Where-Object  {$_.BaseName -notmatch 'SPC'} | % {
Check-Rest -selectedFolder $pathToFolder -currentDocument $_.Name
}
#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</table>
</div>
</body>" -Encoding UTF8
#========Statistics========
}
