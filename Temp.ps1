clear
#Global arrays and variables
$script:PathToFolder = ""
$script:PathToFile = ""
$script:UserInputNotification = ""
$script:JSvariable = 0
$script:CheckPublishedDocuments = $false
$script:CheckDocumentsToBePublished = $false

Function Add-HtmlData ($DocumentsCount, $ExtraColumn)
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
$ExtraColumn
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
                           if ($script:PathToFolder -ne "") {$labelBrowseFolder.Text = "Указан путь: $script:PathToFolder"}
                           
                        if ($script:PathToFolder -eq "" -and $script:PathToFile -eq "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -eq "" -and $script:PathToFile -eq "" -and $radioPublished.Checked -eq $true) {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -ne "" -and $script:PathToFile -ne "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -ne "" -and $script:PathToFile -eq "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -ne "" -and $script:PathToFile -eq "" -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        }
                        else {$buttonRunScript.Enabled = $true}
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
                        if ($script:PathToFile -ne "") {$labelBrowseFile.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:PathToFile))"}
                        
                        if ($script:PathToFolder -eq "" -and $script:PathToFile -eq "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -eq "" -and $script:PathToFile -eq "" -and $radioPublished.Checked -eq $true) {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -ne "" -and $script:PathToFile -ne "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -ne "" -and $script:PathToFile -eq "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        } elseif ($script:PathToFolder -eq "" -and $script:PathToFile -ne "" -and $radioBeingPublished.Checked -eq $true -and $MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        }
                        elseif ($script:PathToFolder -eq "" -and $script:PathToFile -ne "" -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                        $buttonRunScript.Enabled = $false
                        }
                        else {$buttonRunScript.Enabled = $true}

})
#Run Script
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Height = 35
$buttonRunScript.Width = 100
$buttonRunScript.Text = "Запустить скрипт"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 275
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click({
                            if ($radioPublished.Checked -eq $true) {$script:CheckPublishedDocuments = $true}
                            if ($radioBeingPublished.Checked -eq $true) {$script:CheckDocumentsToBePublished = $true}
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
$SystemDrawingPoint.Y = 275
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
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 95
$labelBrowseInput.Location = $SystemDrawingPoint
$labelBrowseInput.Width = 350
$labelBrowseInput.Height = 30
$labelBrowseInput.Enabled = $false
#TextBox
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.Location = New-Object System.Drawing.Size(247,93) 
$MaskedTextBox.Mask = "00-00-0000"
$MaskedTextBox.Width = 61
$MaskedTextBox.BorderStyle = 2
$MaskedTextBox.Enabled = $false
$MaskedTextBox.Add_TextChanged({
                            if ($MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d' -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                           })

#radio buttons
#group
$groupboxSelectType = New-Object System.Windows.Forms.GroupBox
$groupboxSelectType.Height = 130
$groupboxSelectType.Width = 475
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 127
$groupboxSelectType.Text = "Выберите тип проверки"
$groupboxSelectType.Location = $SystemDrawingPoint
$groupboxSelectType.Enabled = $true
#radiobutton published
$radioPublished = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 10
$SystemDrawingPoint.Y = 25
$radioPublished.Location = $SystemDrawingPoint
$radioPublished.Text = "Проверка опубликованного комплекта (изменения уже внесены в файл учета ПД)"
$radioPublished.Width = 460
$radioPublished.Height = 30
$radioPublished.Checked = $true
$radioPublished.Add_Click({
                          if ($radioPublished.Checked -eq $true) {$MaskedTextBox.Enabled = $false; $labelBrowseInput.Enabled = $false}
                          if ($radioPublished.Checked -eq $true -and $script:PathToFolder -eq "" -or $script:PathToFile -eq "") {$buttonRunScript.Enabled = $false} else {$buttonRunScript.Enabled = $true}
                          })
#radiobutton BeingPublished
$radioBeingPublished = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 10
$SystemDrawingPoint.Y = 60
$radioBeingPublished.Location = $SystemDrawingPoint
$radioBeingPublished.Text = "Проверка публикуемого комплекта"
$radioBeingPublished.Width = 435
$radioBeingPublished.Add_Click({
                               if ($radioBeingPublished.Checked -eq $true) {$MaskedTextBox.Enabled = $true; $labelBrowseInput.Enabled = $true}
                               if ($radioBeingPublished.Checked -eq $true -and $script:PathToFolder -eq "" -or $script:PathToFile -eq "" -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                               $buttonRunScript.Enabled = $false
                               } elseif ($radioBeingPublished.Checked -eq $true -and $script:PathToFolder -eq "" -or $script:PathToFile -eq "" -and $MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {
                               $buttonRunScript.Enabled = $false
                               } elseif ($radioBeingPublished.Checked -eq $true -and $script:PathToFolder -ne "" -or $script:PathToFile -ne "" -and $MaskedTextBox.Text -notmatch '\d\d-\d\d-\d\d\d\d') {
                               $buttonRunScript.Enabled = $false
                               }
                               else {$buttonRunScript.Enabled = $true}
                               })

#Add UI elements to the form
$groupboxSelectType.Controls.Add($MaskedTextBox)
$groupboxSelectType.Controls.Add($labelBrowseInput)
$groupboxSelectType.Controls.Add($radioPublished)
$groupboxSelectType.Controls.Add($radioBeingPublished)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($buttonBrowseFolder)
$dialog.Controls.Add($labelBrowseFolder)
$dialog.Controls.Add($buttonBrowseFile)
$dialog.Controls.Add($labelBrowseFile)
$dialog.Controls.Add($groupboxSelectType)
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
        #try catch???
        
        $basenames += $_.BaseName
        $notifications += try {((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
        $changenumbers += try {((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
    } else {
        #if file is a any other file, gets data using coordinate requiret for the table title
        #try catch???
        
        $basenames += $_.BaseName
        $notifications += try {((($document.Tables.Item(1).Cell(7, 5).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
        $changenumbers += try {((($document.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
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
    if ($GreatestValue.Maximum -eq 0) {$CollectedData += [string]""} else {$CollectedData += [string]$GreatestValue.Maximum}
    $CollectedData += [string]$NotificationNumber
    return $CollectedData
}
}

Function Compare-Strings ($FontColor, $DataInDocument, $DataInRegister, $ComparisonResult) 
{
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td>
    <font color=""$FontColor"" onclick=""my_f('div_$script:JSvariable')""><b>$ComparisonResult</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$DataInDocument</td>
            </tr>
            <tr>
            <td id=""indication"">Файл учета:</td>
            <td id=""indication"">$DataInRegister</td>
            </tr>
        </table>
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
}

Function Add-Error ($MessageInDiv) 
{
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td>
    <font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>Ошибка</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
    $MessageInDiv
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td>
    <font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>Ошибка</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
    $MessageInDiv
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
}

Function Add-Status ($Status, $MessageInDiv) 
{
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td>
    <b>$Status</b>
    <div class=""hide"" id=""div_$script:JSvariable"">
    $MessageInDiv
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
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
if ($script:CheckPublishedDocuments -eq $true) {
#========Statistics========
Add-HtmlData -DocumentsCount ($DataFromDocuments[0].Length + $DataFromDocuments[3].Length) -ExtraColumn "" 
#========Statistics========
} else {
#========Statistics========
Add-HtmlData -DocumentsCount ($DataFromDocuments[0].Length + $DataFromDocuments[3].Length) -ExtraColumn "<th>Статус<br>публикуемого документа</th> " 
#========Statistics========
}
for ($i = 0; $i -lt $DataFromDocuments[0].Length; $i++) {
    $DocumentData = @{BaseName = [string]$DataFromDocuments[0][$i]; Notification = [string]$DataFromDocuments[1][$i]; Version = [string]$DataFromDocuments[2][$i]}
    Write-Host "Working on $($DocumentData.BaseName)..."
    $DataFromRegister = Get-DataFromDocumentRegister -ExcelActiveSheet $worksheet -LookFor $DocumentData.BaseName
    if ($DataFromRegister -ne $null) {
    $DocumentDataInRegister = @{Notification = [string]$DataFromRegister[1]; Version = [string]$DataFromRegister[0]}
    }
    if ($script:CheckPublishedDocuments -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<tr>
<td>
    $($DocumentData.BaseName)
</td>"
#========Statistics========
    #if file does not exist in the register
    if ($DataFromRegister -eq $null) {
    Add-Error -MessageInDiv "Ошибка: В файле учета ПД не существует записи о данном документе."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
    #if script cannot get any data from the document title
    } elseif ($DocumentData.Notification -eq "error" -or $DocumentData.Version -eq "error") {
    Add-Error -MessageInDiv "Ошибка: Невозможно получить данные титульного листа."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8   
#========Statistics========  
    } elseif ($DataFromRegister -eq $null -and $DocumentData.Notification -eq "error" -or $DocumentData.Version -eq "error") {
    Add-Error -MessageInDiv "Ошибка: Невозможно получить данные титульного листа.<br>Ошибка: В файле учета ПД не существует записи о данном документе."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8   
#========Statistics========
    } else {
    #if document versions match
    if ($DocumentData.Version -eq $DocumentDataInRegister.Version) {
    Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Совпадает"
    } else {
    #if if document versions do not match
    Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Не совпадает"
    }
    #if notification numbers match
    if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
    Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Совпадает"
    } else {
    Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Не совпадает"
    }
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
    }
    }

    if ($script:CheckDocumentsToBePublished -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<tr>
<td>
    $($DocumentData.BaseName)
</td>"
#========Statistics========
        #if script cannot get value from the document title
        if ($DocumentData.Notification -eq "error" -or $DocumentData.Version -eq "error") {
            Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: Невозможно получить данные титульного листа."
            Add-Error -MessageInDiv "Ошибка: Невозможно получить данные титульного листа."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
        #if cells in the document title are empty and there is no record about this document in the register
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DataFromRegister -eq $null) {
            Add-Status -Status "Новый документ" -MessageInDiv "В документе не указаны номер изменения и номер извещения (пустые ячейки), а в<br> в файле учета ПД отсутсвуют записи о данном документе."
            Compare-Strings -FontColor "Green" -DataInDocument "Номер изменения не указан (пустая ячейка)." -DataInRegister "В файле учета ПД отсутсвуют записи о данном документе." -ComparisonResult "?"
            Compare-Strings -FontColor "Green" -DataInDocument "Номер извещения не указан (пустая ячейка)." -DataInRegister "В файле учета ПД отсутсвуют записи о данном документе." -ComparisonResult "?"
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
        #if cells in the document title are empty and there is a record about this document in the register where the required cells are also empty
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DocumentDataInRegister.Version -eq "" -and $DocumentDataInRegister.Notification -eq "") {
            Add-Status -Status "Старая версия***" -MessageInDiv "В документе не указаны номер изменения и номер извещения (пустые ячейки), а в<br> в файле учета ПД присутсвует запись о данном документе, но в ней не указаны<br>номер изменения и номер извещения (пустые ячейки)."
            Compare-Strings -FontColor "Green" -DataInDocument "Номер изменения не указан (пустая ячейка)." -DataInRegister "Номер изменения не указан (пустая ячейка)." -ComparisonResult "Совпадает"
            Compare-Strings -FontColor "Green" -DataInDocument "Номер извещения не указан (пустая ячейка)." -DataInRegister "Номер извещения не указан (пустая ячейка)." -ComparisonResult "Совпадает"
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
        #if notification number from the document equals the notification number entered by the user (new version)
        } elseif ($DocumentData.Notification -eq $script:UserInputNotification) {
            Add-Status -Status "Новая версия" -MessageInDiv "Номер извещения, указанный в документе, и номер текущего извещения, введеный пользователем, совпадают."
            #if the document version from document equals the document version from register
            if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
            #Add statistics
            #if they do not match
            } else {
            #Add statistics
            }

            #if the notification number from document equals the notification number from register
            if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
            #Add statistics
            #if they do not match
            } else {
            #Add statistics
            }
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
        #if notification number from the document do not equal the notification number entered by the user (old version)
        } elseif ($DocumentData.Notification -ne $script:UserInputNotification) {
            Add-Status -Status "Текущая версия" -MessageInDiv "Номер извещения, указанный в документе, и номер текущего извещения, введеный пользователем, НЕ совпадают."
            #if the document version from document equals the document version from register
            if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
            #Add statistics
            #if they do not match
            } else {
            #Add statistics
            }

            #if the notification number from document equals the notification number from register
            if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
            #Add statistics
            #if they do not match
            } else {
            #Add statistics
            }
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>"
#========Statistics========
        }
    }









    if ($script:CheckDocumentsToBePublished -eq $true) {
    if ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DataFromRegister -eq $null) {
        #new file
        Write-Host $DataFromDocuments[0][$i] "is a new file."
        } else {
        if ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DocumentDataInRegister.Version -eq 0 -and $DocumentDataInRegister.Notification -eq "" -and $DocumentData.Notification -ne $script:UserInputNotification) {
            #file with empty cells has not changed
            Write-Host $DocumentData.BaseName "has not changed, but has empty cells for Change No. and Notification No."  
        }
        if ($DocumentData.Notification -ne "" -and $DocumentData.Version -ne "" <#-and $DataFromDocuments[2][$i] -ne $DataFromRegister[0]#> -and $DocumentData.Notification -eq $script:UserInputNotification) {
            #file changed
            Write-Host $DocumentData.BaseName "has changed."
            if ($DocumentData.Version -ne [string]($DocumentDataInRegister.Version + 1)) {
            #if document change number does not match (register change number +1)
            } else {
            #if document change number matches (register change number +1)
            }
            #Add notification number info ?
        }
        if ($DocumentData.Notification -ne "" -and $DocumentData.Version -ne "" <#-and $DataFromDocuments[2][$i] -eq $DataFromRegister[0]#> -and $DocumentData.Notification -ne $script:UserInputNotification) {
            #file has not changed at all
            Write-Host $DocumentData.BaseName "has not changed at all."  
            #if document change number does not match register change number
            if ($DocumentData.Version -ne  $DocumentDataInRegister.Version) {           
            } else {
            #if document change number matches register change number        
            }
            if ($DocumentData.Notification -ne $DocumentDataInRegister.Notification) {
            #if document notification number does not match register notification number
            } else {
            #if document notification number matches register notification number
            }
        }
        }
        #APPEND INFORMATION ABOUT THE EXCELL FILES TO THE END OF THE FILE!
    }
}
$workbook.Close($false)
$excel.Quit()
