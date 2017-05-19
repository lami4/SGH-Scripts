<#############################################
#UI FOR THE SCRIPT
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
                            $dialog.DialogResult = "OK";
                            $dialog.Close()
                           })
$buttonRunScript.Enabled = $false
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
                             #some action
                             })
#Labels
#Browse folder label
$labelBrowseFolder = New-Object System.Windows.Forms.Label
$labelBrowseFolder.Text = "Укажите путь к папке, в которой необходимо снять MD5 файлов"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 35
$labelBrowseFolder.Location = $SystemDrawingPoint
$labelBrowseFolder.Width = 360
$labelBrowseFolder.Height = 30
#radio buttons
$radioNewList = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 70
$radioNewList.Location = $SystemDrawingPoint
$radioNewList.Text = "Создать новый *.txt со списком файлов и их MD5 сумм"
$radioNewList.Width = 460
$radioNewList.Height = 30
$radioNewList.Checked = $true
$radioNewList.Add_Click({
                          #some action
                          })
$radioExistingList = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 160
$radioExistingList.Location = $SystemDrawingPoint
$radioExistingList.Text = "Использовать существующий *.txt со списком файлов и их MD5 сумм"
$radioExistingList.Width = 460
$radioExistingList.Height = 30
$radioExistingList.Checked = $false
$radioExistingList.Add_Click({
                          #some action
                          })
#inputbox
$TextBox = New-Object System.Windows.Forms.TextBox 
$TextBox.Location = New-Object System.Drawing.Size(120,110) 
$TextBox.Size = New-Object System.Drawing.Size(260,30)
$TextBox.Text = "MD5 файлов.txt"
#labels
$labelTextBox = New-Object System.Windows.Forms.Label
$labelTextBox.Text = "Имя файла:"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 110
$labelTextBox.Location = $SystemDrawingPoint
$labelTextBox.Width = 350
$labelTextBox.Height = 30

$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonBrowseFolder)
$dialog.Controls.Add($labelBrowseFolder)
$dialog.Controls.Add($radioNewList)
$dialog.Controls.Add($radioExistingList)
$dialog.Controls.Add($TextBox)
$dialog.Controls.Add($labelTextBox)
$dialog.ShowDialog()
}
Custom-Form
#############################################>

$script:yesNoUserInput = 0

Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {Return $objForm.SelectedPath} Else {Exit}
}

Function Input-YesOrNo ($Question, $BoxTitle) 
{
    $a = New-Object -ComObject wscript.shell
    $intAnswer = $a.popup($Question,0,$BoxTitle,4)
    If ($intAnswer -eq 6) {$script:yesNoUserInput = 1} else {Exit}
}

$SelectedFolder = Select-Folder -description "Укажите папку, в которой необходимо снять контрольный суммы файлов."
if ((Test-Path -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt") -eq $true) {
$nl = [System.Environment]::NewLine
Input-YesOrNo  -Question "Список 'MD5 файлов текущего релиза.txt' уже сущетвует. Продолжить?$nl$nl`Да - перезаписать и продолжить исполнение скрипта.$nl`Нет - не перезаписывать и остановить исполнение скрипта.$nl$nl`Если вы не хотите перезаписывать существующий список, но хотите продолжить исполнение скрипта - переместите список из папки, где расположен файл скрипта, в любое удобное для вас место и нажмите 'Да'." -BoxTitle "Список уже существует"
if ($script:yesNoUserInput -eq 1) {Remove-Item -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt"}
$script:yesNoUserInput = 0
}
Get-ChildItem -Path "$SelectedFolder\*.*" -Exclude "*.pdf", "*.doc*", "*.xls*" | % {
Add-Content -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt" -Value "$($_.Name):$((Get-FileHash -Path $_.FullName -Algorithm MD5).Hash)"
}
