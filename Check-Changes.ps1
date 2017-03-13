$changes = @()
$excel = new-object -com excel.application
$excel.Visible = $true
$book = $excel.Workbooks.Open("Z:\OTD.Translate\Учет программ и ПД.xls")
$sheet = $book.WorkSheets.Item("Лист1")
#снимать все фильтры
$range = $sheet.Range("E:E")
$target = $range.Find("PABKRF-GL-RU-00.00.00.dRNT.01.00")
if ($target -eq $null) {
    Write-Host "No match found"
    } else {
    $firstHit = $target
    Do
    {
        $changeAddress = $target.AddressLocal($false, $false) -replace "E", ""
        [int]$valueInCell = $sheet.Cells.Item($changeAddress, "J").Value()
        if ($valueInCell.Count -eq 0) {$changes += 0} else {$changes += $valueInCell}
        $target = $range.FindNext($target)
    }
    While ($target -ne $NULL -and $target.AddressLocal() -ne $firstHit.AddressLocal())
    $greatestValue = $changes | Measure-Object -Maximum
    Write-Host $changes
    Write-Host $greatestValue.Maximum
}

#Global arrays and variables
$script:PathToFolder = ""
$script:PathToFile = ""
$script:UserInputNotification = ""

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
                        if ($script:PathToFile-ne "") {
                            $labelBrowseFile.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:PathToFile))"
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
$MaskedTextBox.Location = New-Object System.Drawing.Size(222,135) 
$MaskedTextBox.Mask = "00-00-0000"
$MaskedTextBox.Width = 55
$MaskedTextBox.BorderStyle = 0
$MaskedTextBox.Add_TextChanged({
                            if ($MaskedTextBox.Text -match '\d\d-\d\d-\d\d\d\d') {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
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
$result = Custom-Form
if ($result -ne "OK") {Exit}
Write-Host $script:PathToFolder
Write-Host $script:PathToFile
Write-Host $script:UserInputNotification
