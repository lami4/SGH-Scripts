clear
#Global arrays and variables
$script:PathToFolder = ""
$script:PathToFile = ""
$script:UserInputNotification = ""
$script:JSvariable = 0
Function Add-HtmlData ($DocumentsCount)
{
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Check-Changes Report</title>
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
#changegr {
background-color: #FFC;
}
#notificationgr {
background-color: #FFC;
}
#changered {
background-color: #FFC;
}
#notificationred {
background-color: #FFC;
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
<h3>Анализ</h3>
<h3>Обработано документов: $DocumentsCount</h3>
<table style=""width:60%"">
<tr>
<th>Обозначение</th>
<th>Статус<br>изменений</th> 
<th>Номер<br>изменения</th>
<th>Номер<br>извещения</th>
</tr>" -Encoding UTF8
}

Function Select-File {
Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = "$PSScriptRoot"
$f.Filter = "MS Excel Files (*.xls*)|*.xls*|All Files (*.*)|*.*"
$show = $f.ShowDialog()
If ($show -eq "OK") {$script:PathToFile = $f.FileName}
}

Function Select-Folder ($Description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {$script:PathToFolder = $objForm.SelectedPath}
}

Function Custom-Form {
Add-Type -AssemblyName  System.Windows.Forms
$dialog = New-Object System.Windows.Forms.Form
$dialog.ShowIcon = $false
$dialog.AutoSize = $true
$dialog.Text = "Настройки"
$dialog.AutoSizeMode = "GrowAndShrink"
$dialog.WindowState = "Normal"
$dialog.SizeGripStyle = "Hide"
$dialog.ShowInTaskbar = $true
$dialog.StartPosition = "CenterScreen"
$dialog.MinimizeBox = $false
$dialog.MaximizeBox = $false
#Buttons
#Browse folder
$buttonBrowseFolder = New-Object System.Windows.Forms.Button
$buttonBrowseFolder.Height = 35
$buttonBrowseFolder.Width = 100
$buttonBrowseFolder.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$buttonBrowseFolder.Location = $SystemDrawingPoint
$buttonBrowseFolder.Add_Click({
                           Select-Folder -Description "Укажите путь к папке, содержащей файлы, данные которых необходимо сравнить с данными файла учета ПД."
                           if ($script:PathToFolder -ne "") {
                                $labelBrowseFolder.Text = "Указан путь: $script:PathToFolder"
                                if ($script:PathToFolder -ne "" -and $script:PathToFile -ne "" -and $MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                           }
})
#Browse file
$buttonBrowseFile = New-Object System.Windows.Forms.Button
$buttonBrowseFile.Height = 35
$buttonBrowseFile.Width = 100
$buttonBrowseFile.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 80
$buttonBrowseFile.Location = $SystemDrawingPoint
$buttonBrowseFile.Add_Click({
                        Select-File
                        if ($script:PathToFile -ne "") {
                            $labelBrowseFile.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:PathToFile))"
                            if ($MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d' -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                        }
})
#Run Script
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Height = 35
$buttonRunScript.Width = 100
$buttonRunScript.Text = "Запустить скрипт"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 170
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click({
                            $script:UserInputNotification = $MaskedTextBox.Text
                            $dialog.DialogResult = "OK";
                            $dialog.Close()
                           })
$buttonRunScript.Enabled = $false
#Exit
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Height = 35
$buttonExit.Width = 100
$buttonExit.Text = "Закрыть"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 170
$buttonExit.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$SystemWindowsFormsMargin.Right = 25
$buttonExit.Margin = $SystemWindowsFormsMargin
$buttonExit.Add_Click({
$dialog.Close();
$dialog.DialogResult = "Cancel"
})
#Browse folder label
$labelBrowseFolder = New-Object System.Windows.Forms.Label
$labelBrowseFolder.Text = "Укажите путь к папке, содержащей файлы, данные которых необходимо сравнить с данными файла учета ПД."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 30
$labelBrowseFolder.Location = $SystemDrawingPoint
$labelBrowseFolder.Width = 350
$labelBrowseFolder.Height = 30
#Browse file label
$labelBrowseFile = New-Object System.Windows.Forms.Label
$labelBrowseFile.Text = "Укажите путь к файлу учета ПД."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 90
$labelBrowseFile.Location = $SystemDrawingPoint
$labelBrowseFile.Width = 350
$labelBrowseFile.Height = 30
#Input field label
$labelBrowseInput = New-Object System.Windows.Forms.Label
$labelBrowseInput.Text = "Укажите номер текущего извещения:"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 135
$labelBrowseInput.Location = $SystemDrawingPoint
$labelBrowseInput.Width = 350
$labelBrowseInput.Height = 30
#TextBox
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.Location = New-Object System.Drawing.Size(222,133) 
$MaskedTextBox.Mask = "00-00-0000"
$MaskedTextBox.Width = 61
$MaskedTextBox.BorderStyle = 2
$MaskedTextBox.Add_TextChanged({
                            if ($MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d' -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                           })
#Add UI elements to the form
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($buttonBrowseFolder)
$dialog.Controls.Add($labelBrowseFolder)
$dialog.Controls.Add($buttonBrowseFile)
$dialog.Controls.Add($labelBrowseFile)
$dialog.Controls.Add($MaskedTextBox)
$dialog.Controls.Add($labelBrowseInput)
$dialog.ShowDialog()
}

Function Get-DataFromDocuments 
{
$xlfiles = @()
$basenames = @()
$notifications = @()
$changenumbers = @()
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Get-ChildItem -Path "$script:PathToFolder\*.*" -Include "*.doc*", "*.xls*" | % {
    #If file has *.xls or *.xlsx extensions, adds file to the special array to append it to the table in the report
    if ($_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx") {
        Write-Host "$($_.BaseName): файл с расширением *.xls/*.xlsx, требуется ручная проверка."
        $xlfiles += $_.BaseName
        return
    }
    Write-Host "$($_.BaseName): забираю значения номера изменения и номера извещения..."
    $document = $word.Documents.Open($_.FullName)
    #If file is a specification, gets data using coordinates required for a specification
    if ($_.BaseName -match "SPC") {
        $basenames += $_.BaseName
        $notifications += ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
        $changenumbers += ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
    } else {
        #if file is a any other file, gets data using coordinate requiret for the table title
        $basenames += $_.BaseName
        $notifications += ((($document.Tables.Item(1).Cell(7, 5).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
        $changenumbers += ((($document.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
    }
    $document.Close([ref]0)
}
$word.Quit()
$CollectedData = $basenames, $notifications, $changenumbers, $xlfiles
return $CollectedData
}

Function Get-DataFromDocumentRegister ($ExcelActiveSheet, $LookFor)
{
$CollectedData = @()
$Changes = @()
$NotificationCoordinatesRow = @()
$Range = $ExcelActiveSheet.Range("E:E")
$Target = $Range.Find("$LookFor", [Type]::Missing, [Type]::Missing, 1)
if ($Target -eq $null) {
    return $null
    } else {
    $FirstHit = $Target
    Do
    {
        $ChangeAddress = $Target.AddressLocal($false, $false) -replace "E", ""
        if ($ExcelActiveSheet.Cells.Item($ChangeAddress, "J").Interior.ColorIndex -eq -4142) {
        Write-Host "White background found"
        [int]$ValueInCellChange = $ExcelActiveSheet.Cells.Item($ChangeAddress, "J").Value()
        if ($ValueInCellChange.Length -eq 0) {$Changes += 0} else {$Changes += $ValueInCellChange}
        $NotificationCoordinatesRow += $ChangeAddress
        }
        $Target = $Range.FindNext($Target)
    }
    While ($Target -ne $NULL -and $Target.AddressLocal() -ne $FirstHit.AddressLocal())
    $GreatestValue = $Changes | Measure-Object -Maximum
    $Index = [array]::IndexOf($NotificationCoordinatesRow, $GreatestValue.Maximum)
    [string]$NotificationNumber = $ExcelActiveSheet.Cells.Item($NotificationCoordinatesRow[$Index], "K").Value()
    #Write-Host $LookFor "GreatestValue:"$GreatestValue.Maximum
    #Write-Host $LookFor "Notification:"$NotificationNumber
    #Write-Host "---------------------------"
    $CollectedData += [int]$GreatestValue.Maximum
    $CollectedData += [string]$NotificationNumber
    return $CollectedData
}
}

$result = Custom-Form
if ($result -ne "OK") {Exit}
$DataFromDocuments = Get-DataFromDocuments
<#for ($i = 0; $i -lt $DataFromDocuments[0].Length; $i++) {
Write-Host "Обозначение:"$DataFromDocuments[0][$i] "Количество символов:"$DataFromDocuments[0][$i].Length 
Write-Host "Номер извещения:"$DataFromDocuments[1][$i] "Количество символов:"$DataFromDocuments[1][$i].Length
Write-Host "Номер изменения:"$DataFromDocuments[2][$i] "Количество символов:"$DataFromDocuments[2][$i].Length
if ($DataFromDocuments[1][$i] -eq "") {Write-Host "Ячейка для номера извещения пуста" -ForegroundColor Red}
if ($DataFromDocuments[2][$i] -eq "") {Write-Host "Ячейка для номера изменения пуста" -ForegroundColor Red}
}#>
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Open("$script:PathToFile")
$worksheet = $workbook.Worksheets.Item(1)
if ($worksheet.AutoFilterMode -eq $true) {$worksheet.ShowAllData()}
#========Statistics========
Add-HtmlData -DocumentsCount ($DataFromDocuments[0].Length + $DataFromDocuments[3].Length)
#========Statistics========
for ($i = 0; $i -lt $DataFromDocuments[0].Length; $i++) {
    $DocumentData = @{BaseName = [string]$DataFromDocuments[0][$i]; Notification = [string]$DataFromDocuments[1][$i]; Version = [string]$DataFromDocuments[2][$i]}
    $DataFromRegister = Get-DataFromDocumentRegister -ExcelActiveSheet $worksheet -LookFor $DocumentData.BaseName
    $DocumentDataInRegister = @{Notification = [string]$DataFromRegister[1]; Version = [int]$DataFromRegister[0]}
#Write-Host $DataFromRegister[0]
#Write-Host "----------------------"

    if ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DataFromRegister -eq $null) {
        #new file
        Write-Host $DataFromDocuments[0][$i] "is a new file."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>
<td>
    $($DocumentData.BaseName)
</td>
<td>
    <font onclick=""my_f('div_$script:JSvariable')"">Новый документ</font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">в документе отсутствует номер изменения и номер извещения</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">в файле учета отсутсвуют записи о документе (обозначении)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
$script:JSvariable += 1
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changegr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>?</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">номер изменения не указан (пустая ячейка)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">в файле учета отсутсвуют записи о данном документе (обозначении)</td>
            </tr>
        </table>
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""notificationgr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>?</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">номер извещения не указан (пустая ячейка)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">в файле учета отсутсвуют записи о данном документе (обозначении)</td>
            </tr>
        </table>
    </div>
</td>
</tr>" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1
        } else {
        if ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DocumentDataInRegister.Version -eq 0 -and $DocumentDataInRegister.Notification -eq "" -and $DocumentData.Notification -ne $script:UserInputNotification) {
            #file with empty cells has not changed
            Write-Host $DocumentData.BaseName "has not changed, but has empty cells for Change No. and Notification No."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>
<td>
    $($DocumentData.BaseName)
</td>
<td>
    <font onclick=""my_f('div_$script:JSvariable')"">Без изменений*</font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Номер извещения, указанный в документе, и номер извещения, введеный пользователем, не совпадают.</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
$script:JSvariable += 1   
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changegr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>+</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">номер изменения не указан (пустая ячейка)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">номер изменения не указан (пустая ячейка)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
$script:JSvariable += 1
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""notificationgr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>+</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">номер извещения не указан (пустая ячейка)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">номер извещения не указан (пустая ячейка)</td>
            </tr>
        </table>
    </div>
</td>
</tr>
" -Encoding UTF8
$script:JSvariable += 1 
#========Statistics========     
        }
        if ($DocumentData.Notification -ne "" -and $DocumentData.Version -ne "" <#-and $DataFromDocuments[2][$i] -ne $DataFromRegister[0]#> -and $DocumentData.Notification -eq $script:UserInputNotification) {
            #file changed
            Write-Host $DocumentData.BaseName "has changed."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>
<td>
    $($DocumentData.BaseName)
</td>
<td>
    <font onclick=""my_f('div_$script:JSvariable')"">Изменен</font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Номер извещения, указанный в документе, и номер извещения, введеный пользователем, совпадают.</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1
            if ($DocumentData.Version -ne [string]($DocumentDataInRegister.Version + 1)) {
            #if document change number does not match (register change number +1)
#========Statistics========  
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changered"">
    <font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Version)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$([string]($DocumentDataInRegister.Version + 1)) ($($DocumentDataInRegister.Version)+1)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1  
            } else {
            #if document change number matches (register change number +1)
#========Statistics========  
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changegr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>+</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Version)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$([string]($DocumentDataInRegister.Version + 1)) ($($DocumentDataInRegister.Version)+1)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1 
            }
            #Add notification number info ?
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""notificationgr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>?</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Номер извещения указанный в документе:</td>
            <td id=""indication"">$($DocumentData.Notification)</td>
            </tr>
            <tr>
            <td id=""indication"">Номер извещения введенный пользователем:</td>
            <td id=""indication"">$script:UserInputNotification</td>
            </tr>
        </table>
    </div>
</td>
</tr>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1
        }
        if ($DocumentData.Notification -ne "" -and $DocumentData.Version -ne "" <#-and $DataFromDocuments[2][$i] -eq $DataFromRegister[0]#> -and $DocumentData.Notification -ne $script:UserInputNotification) {
            #file has not changed at all
            Write-Host $DocumentData.BaseName "has not changed at all."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>
<td>
    $($DocumentData.BaseName)
</td>
<td>
    <font onclick=""my_f('div_$script:JSvariable')"">Без изменений</font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Номер извещения, указанный в документе, и номер извещения, введеный пользователем, не совпадают.</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1   
            #if document change number does not match register change number
            if ($DocumentData.Version -ne  $DocumentDataInRegister.Version) {
#========Statistics========  
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changered"">
    <font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Version)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$($DocumentDataInRegister.Version)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1             
            } else {
            #if document change number matches register change number
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""changegr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>+</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Version)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$($DocumentDataInRegister.Version)</td>
            </tr>
        </table>
    </div>
</td>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1           
            }
            if ($DocumentData.Notification -ne $DocumentDataInRegister.Notification) {
            #if document notification number does not match register notification number
#========Statistics========  
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""notificationred"">
    <font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>-</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Notification)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$($DocumentDataInRegister.Notification)</td>
            </tr>
        </table>
    </div>
</td>
</tr>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1  
            } else {
            #if document notification number matches register notification number
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td id=""notificationgr"">
    <font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>+</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$($DocumentData.Notification)</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$($DocumentDataInRegister.Notification)</td>
            </tr>
        </table>
    </div>
</td>
</tr>
" -Encoding UTF8
#========Statistics========
$script:JSvariable += 1   
            }
        }
        }
        #APPEND INFORMATION ABOUT THE EXCELL FILES TO THE END OF THE FILE!
}
$workbook.Close($false)
$excel.Quit()
Add-Content "$PSScriptRoot\Check-Changes Report.html" "
</table>
</body>
</html>" -Encoding UTF8
