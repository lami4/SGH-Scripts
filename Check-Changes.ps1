clear
#Global arrays and variables
$script:PathToFolder = ""
$script:PathToFile = ""
$script:UserInputNotification = ""
$script:JSvariable = 0
$script:CheckPublishedDocuments = $false
$script:CheckDocumentsToBePublished = $false
$script:xlfiles = @()
$script:pdffiles = @()

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

Function Show-MessageBox ()
{ 
    param($Message, $Title, [ValidateSet("OK", "OKCancel", "YesNo")]$Type)
    Add-Type –AssemblyName System.Windows.Forms 
    if ($Type -eq "OK") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title")}  
    if ($Type -eq "OKCancel") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::OKCancel)}
    if ($Type -eq "YesNo") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::YesNo)}
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
    if (Test-Path -Path "$PSScriptRoot\Check-Changes Report.html") {
        if ((Show-MessageBox -Message "Отчет Check-Changes Report.html уже существует в папке.`r`n`r`nНажмите Да, чтобы продолжить (отчет будет перезаписан).`r`nНажмите Нет, чтобы приостановить проверку." -Title "Отчет уже существует" -Type YesNo) -eq "Yes") {
        Remove-Item -Path "$PSScriptRoot\Check-Changes Report.html"
        if ($radioPublished.Checked -eq $true) {$script:CheckPublishedDocuments = $true}
        if ($radioBeingPublished.Checked -eq $true) {$script:CheckDocumentsToBePublished = $true}
        $script:UserInputNotification = $MaskedTextBox.Text
        $dialog.DialogResult = "OK";
        $dialog.Close()
        }
    } else {
        if ($radioPublished.Checked -eq $true) {$script:CheckPublishedDocuments = $true}
        if ($radioBeingPublished.Checked -eq $true) {$script:CheckDocumentsToBePublished = $true}
        $script:UserInputNotification = $MaskedTextBox.Text
        $dialog.DialogResult = "OK";
        $dialog.Close()
    }
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
$labelBrowseFolder.Text = "Укажите путь к папке с файлами, данные которых необходимо сверить с данными файла учета ПД."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 30
$labelBrowseFolder.Location = $SystemDrawingPoint
$labelBrowseFolder.Width = 395
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
$labelBrowseInput.Text = "Номер извещения публикуемого комплекта:"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 95
$labelBrowseInput.Location = $SystemDrawingPoint
$labelBrowseInput.Width = 350
$labelBrowseInput.Height = 30
$labelBrowseInput.Enabled = $false
#TextBox
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.Location = New-Object System.Drawing.Size(284,93) 
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
Write-Host "Считываю значения из титульных листов документов..."
$basenames = @()
$notifications = @()
$changenumbers = @()
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Get-ChildItem -Path "$script:PathToFolder\*.*" -Include "*.doc*", "*.xls*", "*.pdf" | % {
    #If file has *.xls or *.xlsx extensions, adds file to the special array to append it to the table in the report
    if ($_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx") {
        Write-Host "$($_.BaseName): Файл с расширением *.xls/*.xlsx, требуется ручная проверка."
        $script:xlfiles += $_.BaseName
        return
    }
    if ($_.Extension -eq ".pdf" -and $_.BaseName -notmatch "DSG") {
    return
    }
    if ($_.Extension -eq ".pdf" -and $_.BaseName -match "DSG") {
        Write-Host "$($_.BaseName): DSG-файл с расширением *.pdf, требуется ручная проверка."
        $script:pdffiles += $_.BaseName
        return
    }
    Write-Host "$($_.BaseName): Считываю значения номера изменения и номера извещения."
    $document = $word.Documents.Open($_.FullName)
    #If file is a specification, gets data using coordinates required for a specification
    if ($_.BaseName -match "SPC" -or $_.BaseName -match "LPD") {
        $basenames += $_.BaseName
        $notifications += try {((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
        $changenumbers += try {((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
    #if file is any other file than SPC, gets data using coordinates required for the table title
    } else {
        $basenames += $_.BaseName
        $notifications += try {((($document.Tables.Item(1).Cell(7, 5).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
        $changenumbers += try {((($document.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')} catch {"error"}
    }
    $document.Close([ref]0)
}
$word.Quit()
$CollectedData = $basenames, $notifications, $changenumbers
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
        #Write-Host "White background found"
        [int]$ValueInCellChange = $ExcelActiveSheet.Cells.Item($ChangeAddress, "J").Value()
        if ($ValueInCellChange.Length -eq 0) {$Changes += 0} else {$Changes += $ValueInCellChange}
        $NotificationCoordinatesRow += $ChangeAddress
        }
        $Target = $Range.FindNext($Target)
    }
    While ($Target -ne $NULL -and $Target.AddressLocal() -ne $FirstHit.AddressLocal())
    if ($Changes.Length -gt 0) {
    $GreatestValue = $Changes | Measure-Object -Maximum
    $Index = [array]::IndexOf($Changes, [int]$GreatestValue.Maximum)
    [string]$NotificationNumber = $ExcelActiveSheet.Cells.Item($NotificationCoordinatesRow[$Index], "K").Value()
    if ($GreatestValue.Maximum -eq 0) {$CollectedData += [string]""} else {$CollectedData += [string]$GreatestValue.Maximum}
    $CollectedData += [string]$NotificationNumber
    return $CollectedData
    } else {
    return "NoWhiteBackground"
    }
}
}

Function Compare-Strings ($FontColor, $DataInDocument, $DataInRegister, $ComparisonResult, $Title) 
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
            <td id=""indication"">$($Title):</td>
            <td id=""indication"">$DataInRegister</td>
            </tr>
        </table>
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
}

Function Compare-StringsPlustNotificationEnteredByUser ($FontColor, $DataInDocument, $DataInRegister, $ComparisonResult, $Title, $UserInput) 
{
Add-Content "$PSScriptRoot\Check-Changes Report.html" "<td>
    <font color=""$FontColor"" onclick=""my_f('div_$script:JSvariable')""><b>$ComparisonResult</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
        <table>
            <tr>
            <td id=""indication"">Введенный пользователем:</td>
            <td id=""indication"">$UserInput</td>
            </tr>
            <tr>
            <td id=""indication"">Документ:</td>
            <td id=""indication"">$DataInDocument</td>
            </tr>
            <tr>
            <td id=""indication"">$($Title):</td>
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
    <font color=""black"" onclick=""my_f('div_$script:JSvariable')""><b>$Status</b></font>
    <div class=""hide"" id=""div_$script:JSvariable"">
    $MessageInDiv
    </div>
</td>" -Encoding UTF8
$script:JSvariable += 1
}

Function Add-ExecutionTimeToReport ($Time, $ReportName, $StringToReplace) {
$StringForHTML = "<font color=""black"" size=""1"">Для создания данного отчета мне потребовалось всего лишь:`r`n<br>"
$StringForHTML += "$($Time.Days) дней "
$StringForHTML += "$($Time.Hours) часов "
$StringForHTML += "$($Time.Minutes) минут "
$StringForHTML += "$($Time.Seconds) секунд`r`n<br>`r`n</font>`r`n$StringToReplace"
(Get-Content -Path "$PSScriptRoot\$ReportName.html").Replace($StringToReplace, $StringForHTML) | Set-Content("$PSScriptRoot\$ReportName.html") -Encoding UTF8
}

$result = Custom-Form
if ($result -ne "OK") {Exit}
$ExecutionTime = Measure-Command {
$DataFromDocuments = Get-DataFromDocuments
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
Write-Host "Начал проверку считанных значений..."
for ($i = 0; $i -lt $DataFromDocuments[0].Length; $i++) {
    $DocumentData = @{BaseName = [string]$DataFromDocuments[0][$i]; Notification = [string]$DataFromDocuments[1][$i]; Version = [string]$DataFromDocuments[2][$i]}
    Write-Host "Проверяю $($DocumentData.BaseName)..."
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
        #if the Version field is filled out, but the Notification number field is not in the register
        if ($DocumentDataInRegister.Notification -eq "" -and $DocumentDataInRegister.Version -ne "") {
        Add-Error -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics======== 
        #if the Notification number field is filled out, but the Version field is not in the register
        } elseif ($DocumentDataInRegister.Notification -ne "" -and $DocumentDataInRegister.Version -eq "") {
        Add-Error -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========        
        #if the Version filed is filled out, but the Notification number field is not in the document
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -ne "") {
            Add-Error -MessageInDiv "Ошибка: В документе заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if the Notification number field is filled out, but the Version field is not in the document
        } elseif ($DocumentData.Notification -ne "" -and $DocumentData.Version -eq "") {
            Add-Error -MessageInDiv "Ошибка: В документе заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if file does not exist in the register
        } elseif ($DataFromRegister -eq $null) {
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
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Совпадает" -Title "Файл учета"
            } else {
                #if if document versions do not match
                Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Не совпадает" -Title "Файл учета"
            }
            #if notification numbers match
            if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Совпадает" -Title "Файл учета"
            } else {
                Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Не совпадает" -Title "Файл учета"
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
        #if the Version field is filled out, but the Notification number field is not in the register
        if ($DocumentDataInRegister.Notification -eq "" -and $DocumentDataInRegister.Version -ne "") {
        Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
        Add-Error -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics======== 
        #if the Notification number field is filled out, but the Version field is not in the register
        } elseif ($DocumentDataInRegister.Notification -ne "" -and $DocumentDataInRegister.Version -eq "") {
        Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
        Add-Error -MessageInDiv "Ошибка: В файле учета ПД заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========        
        #if the Version filed is filled out, but the Notification number field is not in the document
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -ne "") {
            Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: В документе заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
            Add-Error -MessageInDiv "Ошибка: В документе заполнено поле 'Номер изменения', но не заполнено поле 'Номер извещение'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #below
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DataFromRegister -eq "NoWhiteBackground") {
            Add-Status -Status "Новый документ***" -MessageInDiv "В документе не указаны номер изменения и номер извещения (пустые ячейки), но файл учета ПД содержит аннулированный документ с таким же обозначением.<br>Крайне вероятно, что данное обозначение используется повторно (т.е. обозначение перестало быть уникальным в рамках данного проекта)."
            Compare-Strings -FontColor "Green" -DataInDocument "Номер изменения не указан (пустая ячейка)." -DataInRegister "Файл учета ПД содержит аннулированный документ с таким же обозначением." -ComparisonResult "?" -Title "Файл учета"
            Compare-Strings -FontColor "Green" -DataInDocument "Номер извещения не указан (пустая ячейка)." -DataInRegister "Файл учета ПД содержит аннулированный документ с таким же обозначением." -ComparisonResult "?" -Title "Файл учета"
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if the Notification number field is filled out, but the Version field is not in the document
        } elseif ($DocumentData.Notification -ne "" -and $DocumentData.Version -eq "") {
            Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: В документе заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
            Add-Error -MessageInDiv "Ошибка: В документе заполнено поле 'Номер извещение', но не заполнено поле 'Номер изменения'."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if script cannot get value from the document title
        } elseif ($DocumentData.Notification -eq "error" -or $DocumentData.Version -eq "error") {
            Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: Невозможно получить данные титульного листа."
            Add-Error -MessageInDiv "Ошибка: Невозможно получить данные титульного листа."
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if cells in the document title are empty and there is no record about this document in the register
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DataFromRegister -eq $null) {
            Add-Status -Status "Новый документ" -MessageInDiv "В документе не указаны номер изменения и номер извещения (пустые ячейки), а в<br> в файле учета ПД отсутсвуют записи о данном документе."
            Compare-Strings -FontColor "Green" -DataInDocument "Номер изменения не указан (пустая ячейка)." -DataInRegister "В файле учета ПД отсутсвуют записи о данном документе." -ComparisonResult "?" -Title "Файл учета"
            Compare-Strings -FontColor "Green" -DataInDocument "Номер извещения не указан (пустая ячейка)." -DataInRegister "В файле учета ПД отсутсвуют записи о данном документе." -ComparisonResult "?" -Title "Файл учета"
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if cells in the document title are empty and there is a record about this document in the register where the required cells are also empty
        } elseif ($DocumentData.Notification -eq "" -and $DocumentData.Version -eq "" -and $DocumentDataInRegister.Version -eq "" -and $DocumentDataInRegister.Notification -eq "") {
            Add-Status -Status "Старая версия***" -MessageInDiv "В документе не указаны номер изменения и номер извещения (пустые ячейки), а в<br> в файле учета ПД присутсвует запись о данном документе, но в ней не указаны<br>номер изменения и номер извещения (пустые ячейки)."
            Compare-Strings -FontColor "Green" -DataInDocument "Номер изменения не указан (пустая ячейка)." -DataInRegister "Номер изменения не указан (пустая ячейка)." -ComparisonResult "Совпадает" -Title "Файл учета"
            Compare-Strings -FontColor "Green" -DataInDocument "Номер извещения не указан (пустая ячейка)." -DataInRegister "Номер извещения не указан (пустая ячейка)." -ComparisonResult "Совпадает" -Title "Файл учета"
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if notification number from the document equals the notification number entered by the user (new version)
        } elseif ($DocumentData.Notification -eq $script:UserInputNotification) {
            Add-Status -Status "Новая версия" -MessageInDiv "Номер извещения, указанный в документе, и номер текущего извещения, введеный пользователем, совпадают."
            #if the document version from document equals the document version from register
            if ($DocumentData.Version -eq ([int]$DocumentDataInRegister.Version + 1)) {
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Version -DataInRegister "$(([int]$DocumentDataInRegister.Version + 1)) ($([int]$DocumentDataInRegister.Version)+1)" -ComparisonResult "Совпадает" -Title "Файл учета + 1"
                #if they do not match
                } else {
                Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Version -DataInRegister "$(([int]$DocumentDataInRegister.Version + 1)) ($([int]$DocumentDataInRegister.Version)+1)" -ComparisonResult "Не совпадает" -Title "Файл учета + 1"
                }
            #if the notification number from document equals the notification number from register
            if ($DocumentData.Notification -eq $script:UserInputNotification) {
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Notification -DataInRegister $script:UserInputNotification -ComparisonResult "Совпадает" -Title "Значение, введенное пользователем"
                #if they do not match
                } else {
                Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Notification -DataInRegister $script:UserInputNotification -ComparisonResult "Не совпадает" -Title "Значение, введенное пользователем"
                }
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        #if notification number from the document do not equal the notification number entered by the user (old version)
        } elseif ($DocumentData.Notification -ne $script:UserInputNotification) {
            #if the register does not contain any records about the document
            if ($DataFromRegister -eq $null) {
                Add-Status -Status "Ошибка" -MessageInDiv "Ошибка: В файле учета ПД не существует записи о данном документе."
                Add-Error -MessageInDiv "Ошибка: В файле учета ПД не существует записи о данном документе."
            } else {
                Add-Status -Status "Текущая версия" -MessageInDiv "Номер извещения, указанный в документе, и номер текущего извещения, введеный пользователем, НЕ совпадают."
                #if the document version from document equals the document version from register
                if ($DocumentData.Version -eq $DocumentDataInRegister.Version) {
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Совпадает" -Title "Файл учета"
                #if they do not match
                } else {
                Compare-Strings -FontColor "Red" -DataInDocument $DocumentData.Version -DataInRegister $DocumentDataInRegister.Version -ComparisonResult "Не совпадает" -Title "Файл учета"
                }
                #if the notification number from document equals the notification number from register
                if ($DocumentData.Notification -eq $DocumentDataInRegister.Notification) {
                Compare-Strings -FontColor "Green" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Совпадает" -Title "Файл учета"
                #if they do not match
                } else {
                Compare-StringsPlustNotificationEnteredByUser -FontColor "Red" -DataInDocument $DocumentData.Notification -DataInRegister $DocumentDataInRegister.Notification -ComparisonResult "Не совпадает" -Title "Файл учета" -UserInput $script:UserInputNotification
                }
            }
#========Statistics========
Add-Content "$PSScriptRoot\Check-Changes Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        }
    }
}
$workbook.Close($false)
$excel.Quit()
}
Add-ExecutionTimeToReport -Time $ExecutionTime -ReportName "Check-Changes Report" -StringToReplace "<h3>Анализ</h3>"
#========Statistics========
$script:xlfiles | % {
Add-Content "$PSScriptRoot\Check-Changes Report.html" "
<tr>
<td>$_</td>
<td colspan=""2"">Файл в формате XLS/XLSX. Требуется ручная проверка.</td>
</tr>" -Encoding UTF8
}

$script:pdffiles | % {
Add-Content "$PSScriptRoot\Check-Changes Report.html" "
<tr>
<td>$_</td>
<td colspan=""2"">DSG-файл в формате PDF. Требуется ручная проверка.</td>
</tr>" -Encoding UTF8
}

Add-Content "$PSScriptRoot\Check-Changes Report.html" "</table>
</div>
</body>
<br>
<br>
<br>
<br>
<br>
<br>
<br>" -Encoding UTF8
#========Statistics========
Write-Host "Проверка закончена. Результаты см. в отчете"
