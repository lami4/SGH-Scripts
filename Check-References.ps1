clear
#Script arrays and variables
$script:JSvariable = 0
$script:CheckTitlesAndNames = $false
$script:CheckMD5 = $false
$script:CalculateMD5 = $false
$script:UseList = $false
$script:SelectedList = ""
$script:SelectedCurrentVersionFolder = ""
$script:yesNoUserInput = 0

#Functions
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
#Run Script
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Height = 35
$buttonRunScript.Width = 100
$buttonRunScript.Text = "Запустить скрипт"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 270
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click({
                            if ($checkboxCheckTitlesAndNames.Checked) {$script:CheckTitlesAndNames = $true};
                            if ($checkboxCheckMD5.Checked) {$script:CheckMD5 = $true};
                            if ($radioCalculateMD5.Checked) {$script:CalculateMD5 = $true; $script:UseList = $false};
                            if ($radioUsePrecalculatedMD5.Checked) {$script:CalculateMD5 = $false; $script:UseList = $true};
                            $dialog.DialogResult = "OK";
                            $dialog.Close()})
$buttonRunScript.Enabled = $false
#Exit
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Height = 35
$buttonExit.Width = 100
$buttonExit.Text = "Закрыть"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 270
$buttonExit.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonExit.Margin = $SystemWindowsFormsMargin
$buttonExit.Add_Click({
$dialog.Close();
$dialog.DialogResult = "Cancel"
})
#Browse
$buttonBrowseCV = New-Object System.Windows.Forms.Button
$buttonBrowseCV.Height = 35
$buttonBrowseCV.Width = 100
$buttonBrowseCV.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 78
$buttonBrowseCV.Location = $SystemDrawingPoint
$buttonBrowseCV.Enabled = $false
$buttonBrowseCV.Add_Click({
                           Select-CurrentVersionFolder -description "Укажите путь к папке с документами и файлами текущей версии"
                           if ($script:SelectedCurrentVersionFolder -ne "") {
                           $buttonRunScript.Enabled = $true
                           $labelBrowseCV.Text = "Указан путь: $script:SelectedCurrentVersionFolder"
                           }
})
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Height = 35
$buttonBrowse.Width = 100
$buttonBrowse.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 28
$SystemDrawingPoint.Y = 91
$buttonBrowse.Location = $SystemDrawingPoint
$buttonBrowse.Enabled = $false
$buttonBrowse.Add_Click({
                        Select-File
                        if ($script:SelectedList -ne "") {
                            if ($script:SelectedCurrentVersionFolder -ne "") {$buttonRunScript.Enabled = $true}
                            $labelBrowse.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:SelectedList))"
                        }
})
#Labels
#Browse label
$labelBrowse = New-Object System.Windows.Forms.Label
$labelBrowse.Text = "Укажите файл со списком MD5 неизмененных файлов"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 140
$SystemDrawingPoint.Y = 102
$labelBrowse.Location = $SystemDrawingPoint
$labelBrowse.Width = 305
$labelBrowse.Enabled = $false
#Browse current version label
$labelBrowseCV = New-Object System.Windows.Forms.Label
$labelBrowseCV.Text = "Укажите папку с файлами текущего релиза"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 140
$SystemDrawingPoint.Y = 88
$labelBrowseCV.Location = $SystemDrawingPoint
$labelBrowseCV.Width = 400
$labelBrowseCV.Height = 30
$labelBrowseCV.Enabled = $false
#Check Titles and Names
$checkboxCheckTitlesAndNames = New-Object System.Windows.Forms.CheckBox
$checkboxCheckTitlesAndNames.Width = 475
$checkboxCheckTitlesAndNames.Text = "Сравнить обозначения и имена документов"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$checkboxCheckTitlesAndNames.Location = $SystemDrawingPoint
$checkboxCheckTitlesAndNames.Add_CheckStateChanged({
                                                  if ($checkboxCheckTitlesAndNames.Checked -eq $true -and $checkboxCheckMD5.Checked -eq $false -and $script:SelectedCurrentVersionFolder -eq "") {$buttonRunScript.Enabled = $true};
                                                  if ($checkboxCheckTitlesAndNames.Checked -eq $false -and $checkboxCheckMD5.Checked -eq $false -and $script:SelectedCurrentVersionFolder -eq "") {$buttonRunScript.Enabled = $false};
                                                  if ($checkboxCheckTitlesAndNames.Checked -eq $false -and $checkboxCheckMD5.Checked -eq $true -and $script:SelectedCurrentVersionFolder -ne "") {$buttonRunScript.Enabled = $true};
                                                  if ($checkboxCheckTitlesAndNames.Checked -eq $false -and $checkboxCheckMD5.Checked -eq $false -and $script:SelectedCurrentVersionFolder -ne "") {$buttonRunScript.Enabled = $false};
                                                  if ($checkboxCheckTitlesAndNames.Checked -eq $true -and $checkboxCheckMD5.Checked -eq $false -and $script:SelectedCurrentVersionFolder -ne "") {$buttonRunScript.Enabled = $true};
                                                  })
#Check MD5
$checkboxCheckMD5 = New-Object System.Windows.Forms.CheckBox
$checkboxCheckMD5.Width = 475
$checkboxCheckMD5.Text = "Сравнить контрольные суммы"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 50
$checkboxCheckMD5.Location = $SystemDrawingPoint
$checkboxCheckMD5.Add_CheckStateChanged({
                                        if ($checkboxCheckTitlesAndNames.Checked -or $checkboxCheckMD5.Checked -and $script:SelectedCurrentVersionFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false};
                                        if ($checkboxCheckMD5.Checked) {$groupboxMD5.Enabled = $true} else {$groupboxMD5.Enabled = $false};
                                        if ($checkboxCheckMD5.Checked) {$radioCalculateMD5.Checked = $true} else {$radioCalculateMD5.Checked = $false; $radioUsePrecalculatedMD5.Checked = $false; $buttonBrowse.Enabled = $false; $labelBrowse.Enabled = $false};
                                        if ($checkboxCheckMD5.Checked) {$labelBrowseCV.Enabled = $true; $buttonBrowseCV.Enabled = $true} else {$labelBrowseCV.Enabled = $false; $buttonBrowseCV.Enabled = $false};
                                        if ($checkboxCheckMD5.Checked -eq $false -and $checkboxCheckTitlesAndNames.Checked) {$buttonRunScript.Enabled = $true}
                                        })
#radio button
$groupboxMD5 = New-Object System.Windows.Forms.GroupBox
$groupboxMD5.Height = 140
$groupboxMD5.Width = 450
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 120
$groupboxMD5.Text = "Выберите способ"
$groupboxMD5.Location = $SystemDrawingPoint
$groupboxMD5.Enabled = $false
$radioCalculateMD5 = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 10
$SystemDrawingPoint.Y = 25
$radioCalculateMD5.Location = $SystemDrawingPoint
$radioCalculateMD5.Text = "Подсчитывать MD5 для каждого неизмененного файла перед сравнением"
$radioCalculateMD5.Width = 435
$radioCalculateMD5.Add_CLick({
                            if ($radioCalculateMD5.Checked -and $script:SelectedCurrentVersionFolder -ne "") {$buttonBrowse.Enabled = $false; $labelBrowse.Enabled = $false; $buttonRunScript.Enabled = $true}
                            if ($radioCalculateMD5.Checked -and $script:SelectedCurrentVersionFolder -eq "") {$buttonBrowse.Enabled = $false; $labelBrowse.Enabled = $false; $buttonRunScript.Enabled = $false}
                            })
#$radioCalculateMD5.Enabled = $false
$radioUsePrecalculatedMD5 = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 10
$SystemDrawingPoint.Y = 60
$radioUsePrecalculatedMD5.Location = $SystemDrawingPoint
$radioUsePrecalculatedMD5.Text = "Использовать заранне подсчитанные MD5 неизмененных файлов"
$radioUsePrecalculatedMD5.Width = 435
$radioUsePrecalculatedMD5.Add_Click({
                                    if ($radioUsePrecalculatedMD5.Checked) {$buttonBrowse.Enabled = $true; $labelBrowse.Enabled = $true};
                                    if ($script:SelectedList -eq "" -or $script:SelectedCurrentVersionFolder -eq "") {$buttonRunScript.Enabled = $false} else {$buttonRunScript.Enabled = $true}
                                    })
#Add UI elements to the form
$dialog.Controls.Add($checkboxCheckTitlesAndNames)
$dialog.Controls.Add($checkboxCheckMD5)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($groupboxMD5)
$dialog.Controls.Add($buttonBrowseCV)
$dialog.Controls.Add($labelBrowseCV)
$groupboxMD5.Controls.Add($radioCalculateMD5)
$groupboxMD5.Controls.Add($radioUsePrecalculatedMD5)
$groupboxMD5.Controls.Add($buttonBrowse)
$groupboxMD5.Controls.Add($labelBrowse)
$dialog.ShowDialog()
}

Function Select-File {
Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = "$PSScriptRoot"
$f.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
$show = $f.ShowDialog()
If ($show -eq "OK") {$script:SelectedList = $f.FileName}
}

Function Select-CurrentVersionFolder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {$script:SelectedCurrentVersionFolder = $objForm.SelectedPath}
}

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

Function Input-YesOrNo ($Question, $BoxTitle) {
$a = New-Object -ComObject wscript.shell
$intAnswer = $a.popup($Question,0,$BoxTitle,4)
If ($intAnswer -eq 6) {
$script:yesNoUserInput = 1
} else {Exit}
}

Function Compare-Strings ($SPCvalue, $valueFromDocument, $message, $positive, $negative) 
{
    if ($valueFromDocument -eq $SPCvalue) {
    Write-Host "$message совпадает" -ForegroundColor Green
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green"" onclick=""my_f('div_$script:JSvariable')""><b>$positive</b></font>
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
    Write-Host "$message не совпадает" -ForegroundColor Red
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""red"" onclick=""my_f('div_$script:JSvariable')""><b>$negative</b></font>
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

Function Add-ExecutionTimeToReport ($Time, $ReportName, $StringToReplace) {
$StringForHTML = "<font color=""black"" size=""1"">Для создания данного отчета мне потребовалось всего лишь:`r`n<br>"
$StringForHTML += "$($Time.Days) дней "
$StringForHTML += "$($Time.Hours) часов "
$StringForHTML += "$($Time.Minutes) минут "
$StringForHTML += "$($Time.Seconds) секунд`r`n<br>`r`n</font>`r`n$StringToReplace"
(Get-Content -Path "$PSScriptRoot\$ReportName.html").Replace($StringToReplace, $StringForHTML) | Set-Content("$PSScriptRoot\$ReportName.html") -Encoding UTF8
}

Function Get-DataFromSpecification ($selectedFolder, $currentSPCName) {
    $documentNames = @()
    $documentTitles = @()
    $fileNames = @()
    $fileMd5s = @()
    $FileNameFromList = @()
    $MD5FromList = @()
    Write-Host "--------------------------------------------------------"
    Write-Host "Собираю ссылки на файлы и документы в $currentSPCName..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open("$selectedFolder\$currentSPCName")
    [int]$rowCount = $document.Tables.Item(1).Rows.Count + 1
    for ($i = 1; $i -lt $rowCount; $i++) {
        [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
        if ($valueInDocumentNameCell.length -ne 0) {
        if ($valueInDocumentNameCell -match '\b([A-Z]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d([^\s]*)') {
            if ($script:CheckTitlesAndNames -eq $true) {
                [string]$valueInDocumentTitleCell = (((($document.Tables.Item(1).Cell($i,5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
                $documentNames += $valueInDocumentNameCell
                $documentTitles += $valueInDocumentTitleCell
                }
            } else {
            if ($script:CheckMD5 -eq $true) {
                [string]$valueInFileMd5Cell = (((($document.Tables.Item(1).Cell($i,7).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')).ToLower()
                if ($valueInFileMd5Cell -match '([m,M]\s*[d,D]\s*5)\s*:') {
                    [string]$valueInFileNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                    $fileMd5s += $valueInFileMd5Cell
                    $fileNames += $valueInFileNameCell
                }
            }
            }
        }
        }
    $document.Close([ref]0)
    $documentData = $documentNames, $documentTitles
    $fileData = $fileNames, $fileMd5s
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<th>Название документа/Файла</th>
<th>Документ/Файл</th> 
<th>Обозначение</th>
<th>Наименование</th>
<th>MD5</th>
</tr>" -Encoding UTF8
if ($documentNames.Length -eq 0 -and $documentTitles.Length -eq 0 -and $script:CheckTitlesAndNames -eq $true -and $script:CheckMD5 -eq $false) {
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td colspan=""5"">
Файл не содержит ссылок на документы.
</td>
</tr>" -Encoding UTF8
}
if ($fileMd5s.Length -eq 0 -and $fileNames.Length -eq 0 -and $script:CheckMD5 -eq $true -and $script:CheckTitlesAndNames -eq $false) {
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td colspan=""5"">
Файл не содержит ссылок на файлы.
</td>
</tr>" -Encoding UTF8
}
if ($fileMd5s.Length -eq 0 -and $fileNames.Length -eq 0 -and $documentNames.Length -eq 0 -and $documentTitles.Length -eq 0 -and $script:CheckMD5 -eq $true -and $script:CheckTitlesAndNames -eq $true) {
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td colspan=""5"">
Файл не содержит ссылок на файлы или документы.
</td>
</tr>" -Encoding UTF8
}
#========Statistics========
if ($script:CheckTitlesAndNames -eq $true) {
Write-Host "Сравниваю наименования и обозначения указанные в спецификации..."
    for ($i = 0; $i -lt $documentData[0].Length; $i++) {
    $currentDocumentBaseName = $documentData[0][$i]
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td>$currentDocumentBaseName</td>" -Encoding UTF8
#========Statistics========
    $documentExistence = Test-Path -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
        if ($documentExistence -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green""><b>Найден</b></font></td>" -Encoding UTF8
#========Statistics========
            if ($currentDocumentBaseName -match 'SPC') {
            #FOR SPC
            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            if ($currentDocumentFullName.Extension -eq ".xls" -or $currentDocumentFullName.Extension -eq ".xlsx") {
            Write-Host "$currentDocumentFullName найден (спецификация). Файл имеет расширение *.xls/*.xlsx. Требуется ручная проверка."
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td colspan=""3"">Файл имеет расширение *.xls/*.xlsx. Требуется ручная проверка.</td>
</tr>" -Encoding UTF8
#========Statistics========
            } else {
            Write-Host "$currentDocumentFullName найден (спецификация). Результаты сравнения:"
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "Обозначение" -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "Наименование" -positive "Совпадает" -negative "Не совпадает"
            $document.Close([ref]0)
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
            } else {
            #FOR REST
            $currentDocumentFullName = Get-ChildItem -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
            if ($currentDocumentFullName.Extension -eq ".xls" -or $currentDocumentFullName.Extension -eq ".xlsx") {
            Write-Host "$currentDocumentFullName найден. Файл имеет расширение *.xls/*.xlsx. Требуется ручная проверка."
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td colspan=""3"">Файл имеет расширение *.xls/*.xlsx. Требуется ручная проверка.</td>
</tr>" -Encoding UTF8
#========Statistics========
            } else {
            Write-Host "$currentDocumentFullName найден. Результаты сравнения:"
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Tables.Item(1).Cell(9, 7).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Tables.Item(1).Cell(6, 8).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "Обозначение"  -positive "Совпадает" -negative "Не совпадает"
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "Наименование"  -positive "Совпадает" -negative "Не совпадает"
            $document.Close([ref]0)
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
            }
        } else {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "
<td><font color=""red""><b>Не найден</b></font></td>
<td>---</td>
<td>---</td>
<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
        }
    }
}

if ($script:CheckMD5 -eq $true) {
    for ($i = 0; $i -lt $fileData[0].Length; $i++) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td>$($fileData[0][$i])</td>" -Encoding UTF8
#========Statistics========
        #Check existence in the folder with new/updated files
        if ((Test-Path -Path "$selectedFolder\$($fileData[0][$i])") -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green""><b>Найден</b></font></td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========
            Write-Host "$($fileData[0][$i]) найден. Результаты сравнения:"
            #Get file hash and compare it by using function Compare-String
            Compare-Strings -SPCvalue (($fileData[1][$i] -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$selectedFolder\$($fileData[0][$i])" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает"
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</tr>" -Encoding UTF8
#========Statistics========
        } else {
            #Check existence in the folder with current release
            if ((Test-Path -Path "$script:SelectedCurrentVersionFolder\$($fileData[0][$i])") -eq $true) {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green""><b>Найден</b></font></td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========
                Write-Host "$($fileData[0][$i]) найден. Результаты сравнения:"
                #if user selects to calculate each MD5 before comparing
                if ($script:CalculateMD5 -eq $true) {
                    #Get file hash and compare it by using function Compare-String
                    Compare-Strings -SPCvalue (($fileData[1][$i] -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$script:SelectedCurrentVersionFolder\$($fileData[0][$i])" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает"
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</tr>" -Encoding UTF8
#========Statistics========                    
                }
                #if user selects to use the list
                if ($script:UseList -eq $true) {
                    #1) Put values from the list to an array
                    Get-Content -Path "$script:SelectedList" | % {
                        $FileNameFromList += ($_ -split (":"))[0]
                        $MD5FromList += (($_ -split (":"))[1].ToLower()).Trim(" ")
                    }
                    $FileProperties = $FileNameFromList, $MD5FromList
                    #2) Find the file name in an array
                    if ($FileProperties[0] -contains "$($fileData[0][$i])") {
                        #if the array contains the file name, 3) Get the MD5 of the found file and compare it
                        $index = [array]::IndexOf($FileProperties[0], $fileData[0][$i])
                        Compare-Strings -SPCvalue (($fileData[1][$i] -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument $FileProperties[1][$index].ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает"
                        } else {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""red""><b>Не найден в списке</b></font></td>" -Encoding UTF8
Write-Host "Файл найден в указанной директории, но не найден в списке с MD5 суммами. Сравнить контрольные суммы невозможно."
#========Statistics======== 
                    }
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</tr>" -Encoding UTF8
#========Statistics========               
                }
            } else {
            #file does not exist in any folders -> record in report
            Write-Host "$($fileData[0][$i]) не найден."
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "
<td><font color=""red""><b>Не найден</b></font></td>
<td>---</td>
<td>---</td>
<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
        }
    }
}
    $word.Quit()
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</table>
<br>
<hr>" -Encoding UTF8
#========Statistics========
}


#script code
$result = Custom-Form
if ($result -ne "OK") {Exit}
#<Write-Host $script:CheckTitlesAndNames
Write-Host $script:CheckMD5
Write-Host "Считать суммы" $script:CalculateMD5
Write-Host "Использовать список" $script:UseList
#>
$reportExistence = Test-Path -Path "$PSScriptRoot\Check-References-Report.html"
if ($reportExistence) {
$nl = [System.Environment]::NewLine
Input-YesOrNo -Question "Отчет Check-References-Report.html уже существует. Продолжить?$nl$nl`Да - перезаписать и продолжить исполнение скрипта.$nl`Нет - не перезаписывать и остановить исполнение скрипта.$nl$nl`Если вы не хотите перезаписывать существующий отчет, но хотите продолжить исполнение скрипта - переместите отчет из папки, где расположен файл скрипта, в любое удобное для вас место и нажмите 'Да'." -BoxTitle "Отчет Check-References-Report.html уже существует"
if ($script:yesNoUserInput -eq 1) {Remove-Item -Path "$PSScriptRoot\Check-References-Report.html"}
$script:yesNoUserInput = 0
}
$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно проверить входимость."
$ExecutionTime = Measure-Command {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Check-References Report Report</title>
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
$curSpc = $_.Name
if ($_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx") {
Add-Content "$PSScriptRoot\Check-References-Report.html" "
<table style=""width:80%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
</tr>
<tr>
<td colspan=""5"">Спецификация с расширением *.xls/*.xlsx. Требуется ручная проверка.</td>
</tr>
</table>
<br>
<hr>" -Encoding UTF8
Write-Host "--------------------------------------------------------
Спецификация с расширеникм *.xls/*.xlsx. Требуется ручная проверка."
} else {
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "
<table style=""width:80%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
</tr>" -Encoding UTF8
#========Statistics========
Get-DataFromSpecification -selectedFolder $pathToFolder -currentSPCName $_.Name
}
}
}
Write-Host $executionTime.
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "
</div>
</body>
</html>" -Encoding UTF8
#========Statistics========
}
Add-ExecutionTimeToReport -Time $ExecutionTime -ReportName "Check-References-Report" -StringToReplace "<h3>Результаты сравнения</h3>"
Invoke-Item "$PSScriptRoot\Check-References-Report.html"
