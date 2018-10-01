clear
#Global variables
$script:IntegrityCheckPathToFolder = $null
$script:CollectedReferences = @(), @()
$script:RowCounter = 0
$script:ReportString = ""

Function Show-MessageBox 
{ 
    param($Message, $Title, [ValidateSet("OK", "OKCancel", "YesNo")]$Type)
    Add-Type –AssemblyName System.Windows.Forms 
    if ($Type -eq "OK") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title")}  
    if ($Type -eq "OKCancel") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::OKCancel)}
    if ($Type -eq "YesNo") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::YesNo)}
}

Function Select-Folder ($Description)
{
    Add-Type -AssemblyName System.Windows.Forms
    $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $SelectFolderDialog.Rootfolder = "Desktop"
    $SelectFolderDialog.Description = $Description
    $DialogResult = $SelectFolderDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SelectFolderDialog.SelectedPath} else {return $null}
}
   
Function Check-Integrity () 
{
    Kill -Name WINWORD -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    #СОБИРАЕМ ДАННЫЕ ИЗ СПЕЦИФИКАЦИЙ В УКАЗАННОЙ ПАПКЕ
    #Создать экземпляр приложения MS Word
    $WordApp = New-Object -ComObject Word.Application
    #Сделать вызванное приложение невидемым
    $WordApp.Visible = $false
    #DOCX
    Get-ChildItem -Path "$script:IntegrityCheckPathToFolder\*.docx" | % {
        if ($_.BaseName -match 'SPC' -or $_.BaseName -match 'LPD') {
            Write-Host "Собираю ссылки в файле $($_.Name)..."
            Collect-DataFromSpecification -Word $WordApp -PathToSpecification $_
        }
    }
    #DOC
    Get-ChildItem -Path "$script:IntegrityCheckPathToFolder\*.doc" | % {
        if ($_.BaseName -match 'SPC' -or $_.BaseName -match 'LPD') {
            Write-Host "Собираю ссылки в файле $($_.Name)..."
            Collect-DataFromSpecification -Word $WordApp -PathToSpecification $_
        }
    }
    $WordApp.Quit()
    Kill -Name WINWORD -ErrorAction SilentlyContinue
    #Write-Host $script:CollectedReferences[0]
    #Write-Host $script:CollectedReferences[1]
    #Write-Host "1"
    #СОБИРАЕМ ДАННЫЕ ИЗ СПЕЦИФИКАЦИЙ В УКАЗАННОЙ ПАПКЕ
    #В УКАЗАННОЙ ПАПКЕ ПЕРЕДАЕМ ВСЕ ФАЙЛЫ НА КОНВЕЕР, КОТОРЫЙ ПРОВЕРЯЕТ ССЫЛАЮТСЯ ЛИ НА НИХ СПЕЦИФИКАЦИИ
    Create-HtmlReport
    Get-ChildItem -Path "$script:IntegrityCheckPathToFolder\*.*" | % {
        $script:ReportString = ""
        $script:RowCounter += 1
        $NameWithExtension = $_.Name
        $NameWoExtension = $_.BaseName
        $ArrayOfIndices = @()
        if ($script:CollectedReferences[0] -contains $NameWithExtension) {
            #SOFTWARE
            Write-Host "Software found: $($NameWithExtension)"
            $ArrayOfIndices = Get-IndicesOf -Array $script:CollectedReferences[0] -Value $NameWithExtension
            $ArrayOfIndices | % {$script:ReportString += "<br> <font color=""green""><b>$($script:CollectedReferences[1][$_])</b></font>"}
            $script:ReportString = $script:ReportString.Trim("<br>")
            Add-RowToHtmlReport -File $NameWithExtension -ReferenceInfo $script:ReportString
        } elseif ($script:CollectedReferences[0] -contains $NameWoExtension) {
            #DOCUMENT
            Write-Host "Document found: $($NameWithExtension)"
            $ArrayOfIndices = Get-IndicesOf -Array $script:CollectedReferences[0] -Value $NameWoExtension
            $ArrayOfIndices | % {$script:ReportString += "<br> <font color=""green""><b>$($script:CollectedReferences[1][$_])</b></font>"}
            $script:ReportString = $script:ReportString.Trim("<br>")
            Add-RowToHtmlReport -File $NameWithExtension -ReferenceInfo $script:ReportString
        } else {
            #FILE EXIST BUT NOT REFERENCED BY SPECIFICATIONS
            Write-Host "NOTHING FOUND: $($NameWithExtension)"
            Add-RowToHtmlReport -File $NameWithExtension -ReferenceInfo "<font color=""red""><b>Ссылок не найдено</b></font>"
        }
    }
    Close-HtmlReport
}

Function Get-IndicesOf ($Array, $Value) 
{
  $FoundIndexes = @()
  $i = 0
  foreach ($el in $Array) { 
    if ($el -eq $Value) {$FoundIndexes += $i} 
    ++$i
  }
  return $FoundIndexes
}

Function Collect-DataFromSpecification ($WordApp, $PathToSpecification)
{
    $Specification = $WordApp.Documents.Open("$PathToSpecification")
    $Specification.Tables.Item(1).Rows | % {
        if ($_.Cells.Count -eq 7) {
            if (((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ') -replace [char]13, '') -ne '') {
                Write-Host ((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '')
                $script:CollectedReferences[0] += [string]((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '')
                $script:CollectedReferences[1] += [string][System.IO.Path]::GetFileName("$PathToSpecification")            
            }
        }
    }
    $Specification.Close([ref]0)
}

Function Create-HtmlReport ()
{
    Add-Content "$PSScriptRoot\Отчет.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Отчет о выполненных действиях</title>
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
    text-align:left;
    background-color: #bfbfbf;
}
td {
	padding: 3px;
	border: 1px solid black;
    text-align:left;
    background-color: #FFC;
}
.Item_number {
    width: 3%;
    text-align: left;
}
.File {
    width:47%;
    text-align: left;
}
.Referenced_by {
    width:50%;
    text-align: left;
}
</style>
</head>
<body>
<div>
<h4>Результаты проверки входимости для $($script:IntegrityCheckPathToFolder):</h4>
<div>
</div>
<table style=""width:100%"">
    <tr>
        <th class=""Item_number"">№</th>
        <th class=""File"">Файл</th>
        <th class=""Referenced_by"">Упоминается в</th>
    </tr>" -Encoding UTF8
}

Function Add-RowToHtmlReport ($File, $ReferenceInfo) 
{
    Add-Content "$PSScriptRoot\Отчет.html" "    <tr>
        <td class=""Item_Number"">$($script:RowCounter)</td>
        <td class=""File"">$($File)</td>
        <td class=""Referenced_by"">$($ReferenceInfo)</td>
    </tr>" -Encoding UTF8
}

Function Close-HtmlReport () 
{
    Add-Content "$PSScriptRoot\Отчет.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
}

Function Custom-FormIntegrityCheck 
{
    Add-Type -AssemblyName  System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $CustomForm = New-Object System.Windows.Forms.Form
    $CustomForm.ShowIcon = $false
    $CustomForm.AutoSize = $true
    $CustomForm.Text = "Проверка входимости"
    $CustomForm.AutoSizeMode = "GrowAndShrink"
    $CustomForm.WindowState = "Normal"
    $CustomForm.SizeGripStyle = "Hide"
    $CustomForm.ShowInTaskbar = $true
    $CustomForm.StartPosition = "CenterScreen"
    $CustomForm.MinimizeBox = $false
    $CustomForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #BUTTONS
    #Browse
    $ButtonBrowse = New-Object System.Windows.Forms.Button
    $ButtonBrowse.Location = New-Object System.Drawing.Point(25,25)
    $ButtonBrowse.Size = New-Object System.Drawing.Point(100,25)
    $ButtonBrowse.Text = "Обзор..."
    $SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
    $SystemWindowsFormsMargin.Bottom = 25
    $ButtonBrowse.Margin = $SystemWindowsFormsMargin
    $ButtonBrowse.Add_Click({
        $script:IntegrityCheckPathToFolder = Select-Folder -Description "Укажите путь к папке с текущей версией проекта"
                if ($script:IntegrityCheckPathToFolder -ne $null) {
                    Write-Host $script:IntegrityCheckPathToFolder
                    if ($script:IntegrityCheckPathToFolder.Length -gt 75) {
                        $LabelButtonBrowseForFolder.Text = "Указанный путь не вмещается в поле. Наведите курсором, чтобы отобразить его."
                        $ToolTip.SetToolTip($LabelButtonBrowseForFolder, "$script:IntegrityCheckPathToFolder")
                    } else {
                        $LabelButtonBrowseForFolder.Text = "Указанный путь: " + $script:IntegrityCheckPathToFolder
                    }
                } else {
                     $LabelButtonBrowseForFolder.Text = "Укажите путь к папке с текущей версией проекта"
                }
    })
    $ButtonBrowse.Enabled = $true
    #Run Script
    $ButtonRun = New-Object System.Windows.Forms.Button
    $ButtonRun.Location = New-Object System.Drawing.Point(25,100)
    $ButtonRun.Size = New-Object System.Drawing.Point(100,25)
    $ButtonRun.Text = "Запустить"
    $SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
    $SystemWindowsFormsMargin.Bottom = 25
    $ButtonRun.Margin = $SystemWindowsFormsMargin
    $ButtonRun.Add_Click({
    if ($script:IntegrityCheckPathToFolder -eq $null) {
            Show-MessageBox -Message "Не указан путь к папке с текущей версией проекта!" -Title "Ошибка" -Type OK
        } else {
            if (Test-Path -Path "$PSScriptRoot\Отчет.html") {
                if ((Show-MessageBox -Message "Отчет Проверка входимости.html уже существует в папке.`r`n`r`nНажмите Да, чтобы продолжить (отчет будет перезаписан).`r`nНажмите Нет, чтобы приостановить проверку входимости." -Title "Отчет уже существует" -Type YesNo) -eq "Yes") {
                    Remove-Item -Path "$PSScriptRoot\Отчет.html"
                    Start-Sleep -Seconds 2
                    Check-Integrity
                }
            }
        }
    })
    $ButtonRun.Enabled = $true
    #Exit
    $ButtonExit = New-Object System.Windows.Forms.Button
    $ButtonExit.Location = New-Object System.Drawing.Point(130,100)
    $ButtonExit.Size = New-Object System.Drawing.Point(100,25)
    $ButtonExit.Text = "Закрыть"
    $SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
    $SystemWindowsFormsMargin.Bottom = 25
    $ButtonExit.Margin = $SystemWindowsFormsMargin
    $ButtonExit.Add_Click({$CustomForm.Close();})
    $ButtonExit.Enabled = $true
    #LABELS
    #Browse
    $LabelButtonBrowseForFolder = New-Object System.Windows.Forms.Label
    $LabelButtonBrowseForFolder.Text = "Укажите путь к папке с текущей версией проекта"
    $LabelButtonBrowseForFolder.Location = New-Object System.Drawing.Point(130,31)
    $LabelButtonBrowseForFolder.Width = 510
    $LabelButtonBrowseForFolder.Enabled = $true
    $CustomForm.Controls.Add($ButtonRun)
    $CustomForm.Controls.Add($ButtonExit)
    $CustomForm.Controls.Add($ButtonBrowse)
    $CustomForm.Controls.Add($LabelButtonBrowseForFolder)
    $CustomForm.ShowDialog()
}
Custom-FormIntegrityCheck
