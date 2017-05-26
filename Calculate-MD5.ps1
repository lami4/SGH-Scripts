<#############################################
#UI FOR THE SCRIPT
clear
#Global variables
$script:PathToFolder = ""
$script:PathToFile = ""
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
#Exit
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Height = 35
$buttonExit.Width = 100
$buttonExit.Text = "Закрыть"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 325
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
$SystemDrawingPoint.Y = 325
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
                             $FolderSelectedByUser = Select-Folder -description "Укажите путь к папке, в которой необходимо снять хеш-суммы файлов."
                             if ($FolderSelectedByUser -ne $null) {
                                $script:PathToFolder = $FolderSelectedByUser
                             }
                             if ($script:PathToFolder -ne "") {
                                $labelBrowseFolder.Text = "Указан путь: $script:PathToFolder"
                                Write-Host "Выбран путь: $script:PathToFolder"
                             }
                             if ($radioNewList.Checked -eq $true -and $script:PathToFolder -ne "" -and $TextBox.Text -ne "") {
                                $buttonRunScript.Enabled = $true
                             }
                             if ($radioExistingList.Checked -eq $true -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {
                                $buttonRunScript.Enabled = $true
                             }
                             })
#Browse file
$buttonBrowseFile = New-Object System.Windows.Forms.Button
$buttonBrowseFile.Height = 35
$buttonBrowseFile.Width = 100
$buttonBrowseFile.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 75
$SystemDrawingPoint.Y = 230
$buttonBrowseFile.Location = $SystemDrawingPoint
$buttonBrowseFile.Enabled = $false
$buttonBrowseFile.Add_Click({
                             $FileSelectedByUser = Select-File
                             if ($FileSelectedByUser -ne $null) {
                                $script:PathToFile = $FileSelectedByUser
                             }
                             if ($script:PathToFile -ne "") {
                                $labelBrowseFile.Text = "Указан файл: $([System.IO.Path]::GetFileName($script:PathToFile))"
                                Write-Host "Выбран файл: $script:PathToFile"
                             }
                                if ($radioExistingList.Checked -eq $true -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {$buttonRunScript.Enabled = $true}
                             })
#Labels
#Browse folder label
$labelBrowseFolder = New-Object System.Windows.Forms.Label
$labelBrowseFolder.Text = "Укажите путь к папке, в которой необходимо снять хеш-суммы файлов"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 130
$SystemDrawingPoint.Y = 36
$labelBrowseFolder.Location = $SystemDrawingPoint
$labelBrowseFolder.Width = 400
$labelBrowseFolder.Height = 30
#Browse file label
$labelBrowseFile = New-Object System.Windows.Forms.Label
$labelBrowseFile.Text = "Укажите путь к *.txt файлу"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 180
$SystemDrawingPoint.Y = 240
$labelBrowseFile.Location = $SystemDrawingPoint
$labelBrowseFile.Width = 400
$labelBrowseFile.Height = 30
$labelBrowseFile.Enabled = $false
#textbox label
$labelTextBox = New-Object System.Windows.Forms.Label
$labelTextBox.Text = "Имя файла:"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 75
$SystemDrawingPoint.Y = 165
$labelTextBox.Location = $SystemDrawingPoint
$labelTextBox.Width = 350
$labelTextBox.Height = 30
#combobox label
$labelAlgorithm = New-Object System.Windows.Forms.Label
$labelAlgorithm.Text = "Алгоритм:"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 73
$labelAlgorithm.Location = $SystemDrawingPoint
$labelAlgorithm.Width = 350
$labelAlgorithm.Height = 30
#radio buttons
$radioNewList = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 125
$radioNewList.Location = $SystemDrawingPoint
$radioNewList.Text = "Создать новый *.txt со списком файлов и их хеш-сумм"
$radioNewList.Width = 460
$radioNewList.Height = 30
$radioNewList.Checked = $true
$radioNewList.Add_Click({
                          if ($radioNewList.Checked -eq $true) {$buttonBrowseFile.Enabled = $false; $labelBrowseFile.Enabled = $false; $CheckboxSkipCalculation.Enabled = $false; $labelTextBox.Enabled = $true; $TextBox.Enabled = $true}
                          if ($radioNewList.Checked -eq $true -and ($TextBox.Text -eq "" -or $PathToFolder -eq "")) {$buttonRunScript.Enabled = $false} else {$buttonRunScript.Enabled = $true}
                          })
$radioExistingList = New-Object System.Windows.Forms.RadioButton
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 195
$radioExistingList.Location = $SystemDrawingPoint
$radioExistingList.Text = "Использовать существующий *.txt со списком файлов и их хеш-сумм"
$radioExistingList.Width = 460
$radioExistingList.Height = 30
$radioExistingList.Checked = $false
$radioExistingList.Add_Click({
                          if ($radioExistingList.Checked -eq $true) {$buttonBrowseFile.Enabled = $true; $labelBrowseFile.Enabled = $true; $CheckboxSkipCalculation.Enabled = $true; $labelTextBox.Enabled = $false; $TextBox.Enabled = $false}
                          if ($radioExistingList.Checked -eq $true -and ($PathToFolder -eq "" -or $PathToFile -eq "")) {$buttonRunScript.Enabled = $false} else {$buttonRunScript.Enabled = $true}
                          })
#inputbox
$TextBox = New-Object System.Windows.Forms.TextBox 
$TextBox.Location = New-Object System.Drawing.Size(140,163) 
$TextBox.Size = New-Object System.Drawing.Size(200,30)
$TextBox.Text = "MD5 суммы файлов.txt"
$TextBox.Add_TextChanged({
if ($radioNewList.Checked -eq $true -and $script:PathToFolder -ne "" -and $TextBox.Text -ne "") {
$buttonRunScript.Enabled = $true
} else {
$buttonRunScript.Enabled = $false
}
})
#combobox
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(80,70) 
$DropDownBox.Size = New-Object System.Drawing.Size(180,20) 
$DropDownBox.DropDownHeight = 200
$AlgorithmList = @("MACTripleDES","MD5","RIPEMD160","SHA1","SHA256","SHA384", "SHA512")
$AlgorithmList | % {$DropDownBox.Items.Add($_)} | Out-Null
$DropDownBox.SelectedIndex = 1
$DropDownBox.DropDownStyle = "DropDownList"
$DropDownBox.Add_SelectedIndexChanged({
                                $TextBox.Text = "$($AlgorithmList[$DropDownBox.SelectedIndex]) суммы файлов.txt"
                                if ($radioExistingList.Checked -eq $true -and $script:PathToFolder -ne "" -and $script:PathToFile -ne "") {$buttonRunScript.Enabled = $true}   
                                })
#checkboxes
#IgnoreExtensions
$CheckboxIgnoreExtensions = New-Object System.Windows.Forms.CheckBox
$CheckboxIgnoreExtensions.Width = 475
$CheckboxIgnoreExtensions.Text = "Игнорировать *.pdf и файлы MS Office"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 100
$CheckboxIgnoreExtensions.Location = $SystemDrawingPoint
$CheckboxIgnoreExtensions.Checked = $true
#Do not calculate if hash exist
$CheckboxSkipCalculation = New-Object System.Windows.Forms.CheckBox
$CheckboxSkipCalculation.Width = 510
$CheckboxSkipCalculation.Text = "Не считать хеш-сумму, если *.txt уже содержит хеш-сумму для обрабатываемого файла"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 75
$SystemDrawingPoint.Y = 275
$CheckboxSkipCalculation.Location = $SystemDrawingPoint
$CheckboxSkipCalculation.Enabled = $false
$dialog.Controls.Add($DropDownBox) 
$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonBrowseFolder)
$dialog.Controls.Add($labelBrowseFolder)
$dialog.Controls.Add($radioNewList)
$dialog.Controls.Add($radioExistingList)
$dialog.Controls.Add($TextBox)
$dialog.Controls.Add($labelTextBox)
$dialog.Controls.Add($buttonBrowseFile)
$dialog.Controls.Add($labelBrowseFile)
$dialog.Controls.Add($CheckboxIgnoreExtensions)
$dialog.Controls.Add($CheckboxSkipCalculation)
$dialog.Controls.Add($labelAlgorithm)
$dialog.ShowDialog()
}
Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $show = $objForm.ShowDialog()
    If ($show -eq "OK") {Return $objForm.SelectedPath}
}
Function Select-File 
{
    Add-Type -AssemblyName System.Windows.Forms
    $f = new-object Windows.Forms.OpenFileDialog
    $f.InitialDirectory = "$PSScriptRoot"
    $f.Filter = "MS Excel Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $show = $f.ShowDialog()
    If ($show -eq "OK") {Return $f.FileName}
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
