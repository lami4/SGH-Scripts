clear
#Script arrays and variables
$script:JSvariable = 0
$script:CheckTitlesAndNames = $false
$script:CheckMD5 = $false
$script:CheckConsistency = $false
$script:PathToFilesBeingPublished = ""
$script:PathToPublishedFiles = ""
$script:SelectedMd5ListForFilesBeingPublished = ""
$script:SelectedMd5ListForPublishedFiles = ""
$script:UseMd5ListForFilesBeingPublished = $false
$script:UseMd5ListForPublishedFiles = $false
$script:yesNoUserInput = 0
$script:ReferencesToDocuments = 0
$script:ReferencesToFiles = 0
$script:ItemsReferencedTo = @()

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
$SystemDrawingPoint.Y = 335
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click({
                            if ($checkboxCheckConsistency.Checked -eq $true) {$script:CheckConsistency = $true}
                            if ($checkboxCheckTitlesAndNames.Checked -eq $true) {$script:CheckTitlesAndNames = $true}
                            if ($checkboxCheckMD5.Checked -eq $true) {$script:CheckMD5 = $true}
                            if ($checkboxUseListBeingPublished.Checked -eq $true) {$script:UseMd5ListForFilesBeingPublished = $true}
                            if ($checkboxUseListPublished.Checked -eq $true) {$script:UseMd5ListForPublishedFiles = $true}
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
$SystemDrawingPoint.Y = 335
$buttonExit.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonExit.Margin = $SystemWindowsFormsMargin
$buttonExit.Add_Click({
$dialog.Close();
$dialog.DialogResult = "Cancel"
})
#Browse for the folder with the new release
$ButtonBrowseNewRelease = New-Object System.Windows.Forms.Button
$ButtonBrowseNewRelease.Height = 35
$ButtonBrowseNewRelease.Width = 100
$ButtonBrowseNewRelease.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$ButtonBrowseNewRelease.Location = $SystemDrawingPoint
$ButtonBrowseNewRelease.Enabled = $true
$ButtonBrowseNewRelease.Add_Click({
    $FolderSelectedByUser = Select-Folder -description "Выберите папку, в которой находится комплект скомплектованных публикуемых документов."
    if ($FolderSelectedByUser -ne $null) {
        $script:PathToFilesBeingPublished = $FolderSelectedByUser
    }
        if ($script:PathToFilesBeingPublished -ne "") {
        $labelBrowseForReadyRelease.Text = "Указан путь: $script:PathToFilesBeingPublished"
    }
    Write-Host "($script:PathToFilesBeingPublished)"
    if ($checkboxCheckMD5.Checked -eq $false -and $checkboxCheckTitlesAndNames.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "") {
    $buttonRunScript.Enabled = $true    
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $false -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})
#Browse for files of the current release
$buttonBrowseCV = New-Object System.Windows.Forms.Button
$buttonBrowseCV.Height = 35
$buttonBrowseCV.Width = 100
$buttonBrowseCV.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 145
$buttonBrowseCV.Location = $SystemDrawingPoint
$buttonBrowseCV.Enabled = $false
$buttonBrowseCV.Add_Click({
    $FolderSelectedByUser = Select-Folder -description "Выберите папку, в которой находятся программы (файлы) текущего релиза."
    if ($FolderSelectedByUser -ne $null) {
        $script:PathToPublishedFiles = $FolderSelectedByUser
    }
        if ($script:PathToPublishedFiles -ne "") {
        $labelBrowseCV.Text = "Указан путь: $script:PathToPublishedFiles"
    }
    Write-Host "($script:PathToPublishedFiles)"
    if ($checkboxCheckMD5.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $false -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})
#Browse for the MD5 list of the published files
$ButtonBrowseForMd5ListPublished = New-Object System.Windows.Forms.Button
$ButtonBrowseForMd5ListPublished.Height = 35
$ButtonBrowseForMd5ListPublished.Width = 100
$ButtonBrowseForMd5ListPublished.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 75
$SystemDrawingPoint.Y = 285
$ButtonBrowseForMd5ListPublished.Location = $SystemDrawingPoint
$ButtonBrowseForMd5ListPublished.Enabled = $false
$ButtonBrowseForMd5ListPublished.Add_Click({
    $FileSelectedByUser = Select-File
    if ($FileSelectedByUser -ne $null) {
        $script:SelectedMd5ListForPublishedFiles = $FileSelectedByUser
    }
    if ($script:SelectedMd5ListForPublishedFiles -ne "") 
    {
       $labelBrowseForMd5ListPublished.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:SelectedMd5ListForPublishedFiles))"
    }
    Write-Host "($script:SelectedMd5ListForPublishedFiles)"
    if ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $script:SelectedMd5ListForPublishedFiles  -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "" -and $script:SelectedMd5ListForFilesBeingPublished -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }            
})
#Browse for the MD5 list for the files being published
$ButtonBrowseForMd5ListBeingPublished = New-Object System.Windows.Forms.Button
$ButtonBrowseForMd5ListBeingPublished.Height = 35
$ButtonBrowseForMd5ListBeingPublished.Width = 100
$ButtonBrowseForMd5ListBeingPublished.Text = "Обзор..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 75
$SystemDrawingPoint.Y = 215
$ButtonBrowseForMd5ListBeingPublished.Location = $SystemDrawingPoint
$ButtonBrowseForMd5ListBeingPublished.Enabled = $false
$ButtonBrowseForMd5ListBeingPublished.Add_Click({
    $FileSelectedByUser = Select-File
    if ($FileSelectedByUser -ne $null) {
        $script:SelectedMd5ListForFilesBeingPublished = $FileSelectedByUser
    }
    if ($script:SelectedMd5ListForFilesBeingPublished -ne "") 
    {
       $labelBrowseForMd5ListBeingPublished.Text = "Выбран файл: $([System.IO.Path]::GetFileName($script:SelectedMd5ListForFilesBeingPublished))"
    }
    Write-Host "($script:SelectedMd5ListForFilesBeingPublished)"
    if ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }          
})
#LABELS
#Browse for the folder with a new release prepared to get checked
$labelBrowseForReadyRelease = New-Object System.Windows.Forms.Label
$labelBrowseForReadyRelease.Text = "Укажите путь к скомплектованному комплекту публикуемых документов и программ"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 135
$SystemDrawingPoint.Y = 35
$labelBrowseForReadyRelease.Location = $SystemDrawingPoint
$labelBrowseForReadyRelease.Width = 500
$labelBrowseForReadyRelease.Enabled = $true
#Browse for the md5 list for the published files
$labelBrowseForMd5ListPublished = New-Object System.Windows.Forms.Label
$labelBrowseForMd5ListPublished.Text = "Укажите файл со списком MD5 программ (файлов) текущего релиза"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 185
$SystemDrawingPoint.Y = 295
$labelBrowseForMd5ListPublished.Location = $SystemDrawingPoint
$labelBrowseForMd5ListPublished.Width = 400
$labelBrowseForMd5ListPublished.Enabled = $false
#Browse for md5 list for the files being published
$labelBrowseForMd5ListBeingPublished = New-Object System.Windows.Forms.Label
$labelBrowseForMd5ListBeingPublished.Text = "Укажите файл со списком MD5 публикуемых программ (файлов)"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 185
$SystemDrawingPoint.Y = 225
$labelBrowseForMd5ListBeingPublished.Location = $SystemDrawingPoint
$labelBrowseForMd5ListBeingPublished.Width = 400
$labelBrowseForMd5ListBeingPublished.Enabled = $false
#Browse current version label
$labelBrowseCV = New-Object System.Windows.Forms.Label
$labelBrowseCV.Text = "Укажите папку с программами (файлами) текущего релиза"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 160
$SystemDrawingPoint.Y = 155
$labelBrowseCV.Location = $SystemDrawingPoint
$labelBrowseCV.Width = 400
$labelBrowseCV.Height = 30
$labelBrowseCV.Enabled = $false
#CHECKBOXES
#Check consistency
$checkboxCheckConsistency = New-Object System.Windows.Forms.CheckBox
$checkboxCheckConsistency.Width = 475
$checkboxCheckConsistency.Text = "Проверить комплектность"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 65
$checkboxCheckConsistency.Location = $SystemDrawingPoint
$checkboxCheckConsistency.Enabled = $false
$checkboxCheckConsistency.Add_CheckStateChanged({})
#Check Titles and Names
$checkboxCheckTitlesAndNames = New-Object System.Windows.Forms.CheckBox
$checkboxCheckTitlesAndNames.Width = 475
$checkboxCheckTitlesAndNames.Text = "Сравнить обозначения и имена документов"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 90
$checkboxCheckTitlesAndNames.Location = $SystemDrawingPoint
$checkboxCheckTitlesAndNames.Add_CheckStateChanged({
    if ($checkboxCheckMD5.Checked -eq $true -and $checkboxCheckTitlesAndNames.Checked -eq $true) {$checkboxCheckConsistency.Enabled = $true} else {$checkboxCheckConsistency.Enabled = $false}
    if ($checkboxCheckTitlesAndNames.Checked -eq $true -and $checkboxCheckMD5.Checked -eq $false -and $script:PathToFilesBeingPublished -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false  -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -eq "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $false
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -eq "") {
    $buttonRunScript.Enabled = $false
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})
#Check MD5
$checkboxCheckMD5 = New-Object System.Windows.Forms.CheckBox
$checkboxCheckMD5.Width = 475
$checkboxCheckMD5.Text = "Сравнить контрольные суммы"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 115
$checkboxCheckMD5.Location = $SystemDrawingPoint
$checkboxCheckMD5.Add_CheckStateChanged({
    if ($checkboxCheckTitlesAndNames.Checked -eq $true -and $checkboxCheckMD5.Checked -eq $true) {$checkboxCheckConsistency.Enabled = $true} else {$checkboxCheckConsistency.Enabled = $false}
    if ($checkboxCheckMD5.Checked -eq $true) {
    $buttonBrowseCV.Enabled = $true
    $labelBrowseCV.Enabled = $true
    $checkboxUseListBeingPublished.Enabled = $true
    $checkboxUseListPublished.Enabled = $true
    } else {
    $buttonBrowseCV.Enabled = $false
    $labelBrowseCV.Enabled = $false
    $checkboxUseListBeingPublished.Enabled = $false
    $checkboxUseListPublished.Enabled = $false
    $checkboxUseListBeingPublished.Checked = $false
    $checkboxUseListPublished.Checked = $false
    }
    if ($checkboxCheckMD5.Checked -eq $false -and $checkboxCheckTitlesAndNames.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $script:PathToFilesBeingPublished -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $true -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($checkboxCheckTitlesAndNames.Checked -eq $false -and $script:PathToFilesBeingPublished -ne "" -and $checkboxCheckMD5.Checked -eq $true -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})
#Use MD5 list for files being published
$checkboxUseListBeingPublished = New-Object System.Windows.Forms.CheckBox
$checkboxUseListBeingPublished.Width = 475
$checkboxUseListBeingPublished.Text = "Использовать файл с MD5 публикуемых программ"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 185
$checkboxUseListBeingPublished.Location = $SystemDrawingPoint
$checkboxUseListBeingPublished.Enabled = $false
$checkboxUseListBeingPublished.Add_CheckStateChanged({
    if ($checkboxUseListBeingPublished.Checked -eq $true) {
    $labelBrowseForMd5ListBeingPublished.Enabled = $true; 
    $ButtonBrowseForMd5ListBeingPublished.Enabled = $true
    } else {
    $labelBrowseForMd5ListBeingPublished.Enabled = $false; 
    $ButtonBrowseForMd5ListBeingPublished.Enabled = $false
    }
    
    if ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $true -and $script:SelectedMd5ListForPublishedFiles -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})
#Use MD5 list for published files
$checkboxUseListPublished = New-Object System.Windows.Forms.CheckBox
$checkboxUseListPublished.Width = 475
$checkboxUseListPublished.Text = "Использовать файл с MD5 программ текущего релиза"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 50
$SystemDrawingPoint.Y = 255
$checkboxUseListPublished.Location = $SystemDrawingPoint
$checkboxUseListPublished.Enabled = $false
$checkboxUseListPublished.Add_CheckStateChanged({
    if ($checkboxUseListPublished.Checked -eq $true) {
    $labelBrowseForMd5ListPublished.Enabled = $true; 
    $ButtonBrowseForMd5ListPublished.Enabled = $true
    } else {
    $labelBrowseForMd5ListPublished.Enabled = $false; 
    $ButtonBrowseForMd5ListPublished.Enabled = $false
    }
    if ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $script:SelectedMd5ListForPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "") {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $true -and $script:SelectedMd5ListForFilesBeingPublished -ne "" -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } elseif ($script:PathToFilesBeingPublished -ne "" -and $script:PathToPublishedFiles -ne "" -and $checkboxUseListBeingPublished.Checked -eq $false -and $checkboxUseListPublished.Checked -eq $false) {
    $buttonRunScript.Enabled = $true
    } else {
    $buttonRunScript.Enabled = $false
    }
})

#Add UI elements to the form
$dialog.Controls.Add($checkboxCheckTitlesAndNames)
$dialog.Controls.Add($checkboxCheckMD5)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonExit)
$dialog.Controls.Add($groupboxMD5)
$dialog.Controls.Add($buttonBrowseCV)
$dialog.Controls.Add($labelBrowseCV)
$dialog.Controls.Add($ButtonBrowseNewRelease)
$dialog.Controls.Add($checkboxUseListBeingPublished)
$dialog.Controls.Add($ButtonBrowseForMd5ListBeingPublished)
$dialog.Controls.Add($checkboxUseListPublished)
$dialog.Controls.Add($ButtonBrowseForMd5ListPublished)
$dialog.Controls.Add($labelBrowseForMd5ListPublished)
$dialog.Controls.Add($labelBrowseForMd5ListBeingPublished)
$dialog.Controls.Add($labelBrowseForReadyRelease)
$dialog.Controls.Add($checkboxCheckConsistency)
$dialog.ShowDialog()
}

Function Select-File {
Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = "$PSScriptRoot"
$f.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
$show = $f.ShowDialog()
If ($show -eq "OK") {Return $f.FileName}
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
        }
}

Function Input-YesOrNo ($Question, $BoxTitle) {
$a = New-Object -ComObject wscript.shell
$intAnswer = $a.popup($Question,0,$BoxTitle,4)
If ($intAnswer -eq 6) {
$script:yesNoUserInput = 1
} else {Exit}
}

Function Compare-Strings ($SPCvalue, $valueFromDocument, $message, $positive, $negative, $FileOrDocument) 
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
<td id=""indication"">$($FileOrDocument):</td>
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
<td id=""indication"">$($FileOrDocument):</td>
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
        if ($document.Tables.Item(1).Rows.Item($i).Cells.Count -ne 7) {continue}
        [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
        if ($valueInDocumentNameCell.length -ne 0) {
        if ($valueInDocumentNameCell -match '\b([A-Z0-9]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d([^\s]*)') {
            if ($script:CheckTitlesAndNames -eq $true) {
                [string]$valueInDocumentTitleCell = (((($document.Tables.Item(1).Cell($i,5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
                if ($script:CheckConsistency -eq $true) {$script:ItemsReferencedTo += $valueInDocumentNameCell}
                $script:ReferencesToDocuments += 1
                $documentNames += $valueInDocumentNameCell
                $documentTitles += $valueInDocumentTitleCell
                }
            } else {
            if ($script:CheckMD5 -eq $true) {
                [string]$valueInFileMd5Cell = (((($document.Tables.Item(1).Cell($i,7).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')).ToLower()
                if ($valueInFileMd5Cell -match '([m,M]\s*[d,D]\s*5)\s*:') {
                    [string]$valueInFileNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                    if ($script:CheckConsistency -eq $true) {$script:ItemsReferencedTo += $valueInFileNameCell}
                    $script:ReferencesToFiles += 1
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
Start-Sleep -Seconds 0.5
#========Statistics========
if ($script:CheckTitlesAndNames -eq $true) {
Write-Host "Проверяю наименования и обозначения указанные в спецификации..."
    for ($i = 0; $i -lt $documentData[0].Length; $i++) {
Start-Sleep -Seconds 0.3
    $currentDocumentBaseName = $documentData[0][$i]
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td>$currentDocumentBaseName</td>" -Encoding UTF8
#========Statistics========
    $documentExistence = Test-Path -Path "$selectedFolder\$currentDocumentBaseName.*" -Exclude "*.pdf"
        if ($documentExistence -eq $true) {
Start-Sleep -Seconds 0.2
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
            Write-Host "$($currentDocumentFullName.BaseName) найден (спецификация). Результаты сравнения:"
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
            Start-Sleep -Seconds 0.3
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "Обозначение" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "Документ"
            Start-Sleep -Seconds 0.3
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "Наименование" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "Документ"
            $document.Close([ref]0)
            Start-Sleep -Seconds 0.3
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
            Write-Host "$($currentDocumentFullName.BaseName) найден. Результаты сравнения:"
            $document = $word.Documents.Open("$currentDocumentFullName")
            [string]$valueForDocTitle = (((($document.Tables.Item(1).Cell(9, 7).Range.Text).Trim([char]0x0007)) -replace '\.', ' ' -replace ',', ' ' -replace 'ё', 'е' -replace [char]0x2010, '-' -replace '-', ' ' -replace '\s+', ' ').Trim(' ')).ToLower()
            [string]$valueForDocName = ((($document.Tables.Item(1).Cell(6, 8).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ' -replace [char]0x2010, '-').Trim(' ')
            Start-Sleep -Seconds 0.3
            Compare-Strings -SPCvalue $documentData[0][$i] -valueFromDocument $valueForDocName -message "Обозначение"  -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "Документ"
            Start-Sleep -Seconds 0.3
            Compare-Strings -SPCvalue $documentData[1][$i] -valueFromDocument $valueForDocTitle -message "Наименование"  -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "Документ"
            $document.Close([ref]0)
            Start-Sleep -Seconds 0.3
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td>---</td>
</tr>" -Encoding UTF8
#========Statistics========
            }
            }
        } else {
Start-Sleep -Seconds 0.2
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
    #Get data from the MD5 list for files being published
    if ($script:UseMd5ListForFilesBeingPublished -eq $true) {
        $FilesBeingPublishedData = @(), @()
        Get-Content -Path "$script:SelectedMd5ListForFilesBeingPublished" | % {
            if ($_ -match ":") {
                $FilesBeingPublishedData[0] += (($_ -split (":"))[0]).ToLower(); $FilesBeingPublishedData[1] += (($_ -split (":"))[1].ToLower()).Trim(" ")
            }
        }
    }
    #Get data from the MD5 list for published files
    if ($script:UseMd5ListForPublishedFiles -eq $true) {
        $PublishedFilesData = @(), @()
        Get-Content -Path "$script:SelectedMd5ListForPublishedFiles" | % {
            if ($_ -match ":") {
                $PublishedFilesData[0] += (($_ -split (":"))[0]).ToLower(); $PublishedFilesData[1] += (($_ -split (":"))[1].ToLower()).Trim(" ")
            }
        }
    }
    for ($i = 0; $i -lt $fileData[0].Length; $i++) {
    Start-Sleep -Seconds 0.3
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<tr>
<td>$($fileData[0][$i])</td>" -Encoding UTF8
#========Statistics========
        $FileDataFromSpecification = @{Name = ([string]$fileData[0][$i]).ToLower(); Checksum = [string]$fileData[1][$i]}
        #if the file found in the folder with files being published
        if ((Test-Path -Path "$selectedFolder\$($FileDataFromSpecification.Name)") -eq $true) {
        Write-Host "$($FileDataFromSpecification.Name) найден. Результаты сравнения:"
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green""><b>Найден</b></font></td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========
            #if user selects to use precalculated md5 for files being published
            if ($script:UseMd5ListForFilesBeingPublished -eq $true) {
                if ($FilesBeingPublishedData[0] -contains $FileDataFromSpecification.Name) {
                    $index = [array]::IndexOf($FilesBeingPublishedData[0], $FileDataFromSpecification.Name)
                    Start-Sleep -Seconds 1
                    Compare-Strings -SPCvalue ([string](($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower()) -valueFromDocument ([string]$FilesBeingPublishedData[1][$index]) -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
                    #Write-Host "File found in the list."
                } else {
                    Start-Sleep -Seconds 1
                    Compare-Strings -SPCvalue (($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$selectedFolder\$($FileDataFromSpecification.Name)" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
                    #Write-Host "File not found in the list. Calculating checksum."
                }
            } else {
            Start-Sleep -Seconds 1
            Compare-Strings -SPCvalue (($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$selectedFolder\$($FileDataFromSpecification.Name)" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
            #Write-Host "Calcuating all MD5s"
            }
         #if the file is not found in the folder with files being published, the script goes to the folder that contains files of the current release
        } elseif ((Test-Path -Path "$script:PathToPublishedFiles\$($FileDataFromSpecification.Name)") -eq $true) {
        Write-Host "$($FileDataFromSpecification.Name) найден. Результаты сравнения:"
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""green""><b>Найден</b></font></td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========
            #if user selects to use precalculated md5 for published files
            if ($script:UseMd5ListForPublishedFiles -eq $true) {
                if ($PublishedFilesData[0] -contains $FileDataFromSpecification.Name) {
                    $index = [array]::IndexOf($PublishedFilesData[0], $FileDataFromSpecification.Name)
                    Start-Sleep -Seconds 1
                    Compare-Strings -SPCvalue ([string](($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower()) -valueFromDocument ([string]$PublishedFilesData[1][$index]) -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
                    #Write-Host "File found in the list. (published)"
                } else {
                    Start-Sleep -Seconds 1
                    Compare-Strings -SPCvalue (($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$script:PathToPublishedFiles\$($FileDataFromSpecification.Name)" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
                    #Write-Host "File not found in the list. Calculating checksum. (published)"
                }
            } else {
                Start-Sleep -Seconds 1
                Compare-Strings -SPCvalue (($FileDataFromSpecification.Checksum -split (":"))[1].Trim(' ')).ToLower() -valueFromDocument (Get-FileHash -Path "$script:PathToPublishedFiles\$($FileDataFromSpecification.Name)" -Algorithm MD5).Hash.ToLower() -message "Контрольная сумма MD5" -positive "Совпадает" -negative "Не совпадает" -FileOrDocument "MD5 файла"
                #Write-Host "Calcuating all MD5s (published)"
            }
        #if file not found anywhere
        } else {
        Start-Sleep -Seconds 1
        Write-Host "$($FileDataFromSpecification.Name) не найден."
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "<td><font color=""red""><b>Не найден</b></font></td>
<td>---</td>
<td>---</td>
<td>---</td>" -Encoding UTF8
#========Statistics========        
        }
        Start-Sleep -Seconds 1
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</tr>" -Encoding UTF8
#========Statistics========
    }
}
    $word.Quit()
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "</table>
<br>
<hr>
</div>" -Encoding UTF8
#========Statistics========
}


#script code
$result = Custom-Form
if ($result -ne "OK") {Exit}
<#Write-Host $script:CheckTitlesAndNames
Write-Host $script:CheckMD5
Write-Host "Использовать список" $script:UseList
#>
$reportExistence = Test-Path -Path "$PSScriptRoot\Check-References-Report.html"
if ($reportExistence) {
$nl = [System.Environment]::NewLine
Input-YesOrNo -Question "Отчет Check-References-Report.html уже существует. Продолжить?$nl$nl`Да - перезаписать и продолжить исполнение скрипта.$nl`Нет - не перезаписывать и остановить исполнение скрипта.$nl$nl`Если вы не хотите перезаписывать существующий отчет, но хотите продолжить исполнение скрипта - переместите существующий отчет из папки, где расположен файл скрипта, в любое удобное место и нажмите 'Да'." -BoxTitle "Отчет Check-References-Report.html уже существует"
if ($script:yesNoUserInput -eq 1) {Remove-Item -Path "$PSScriptRoot\Check-References-Report.html"}
$script:yesNoUserInput = 0
}
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

function filterErrors() {
    var divs = document.getElementsByClassName('specification');
    if (document.getElementById('filterbutton').innerHTML == 'Показать только спецификации с ошибками') {   
        for (var i=0; i<divs.length; i++) {
            if (divs[i].innerHTML.toLowerCase().indexOf('red') == -1) {
            divs[i].style.display = ""none"";
            document.getElementById('filterbutton').innerHTML = 'Показать все спецификации';
            }
        }
    } else {
        for (var i=0; i<divs.length; i++) {
            divs[i].style.display = """";
            document.getElementById('filterbutton').innerHTML = 'Показать только спецификации с ошибками';
            }
    }
}
</script>
</head>
<body>
<div>
<h3>Анализ</h3>
" -Encoding UTF8
#========Statistics========
Measure-Command {
Get-ChildItem "$script:PathToFilesBeingPublished\*.*" -File -Exclude "*.pdf" | Where-Object {$_.Name -match "SPC"} | % {
$curSpc = $_.Name
if ($_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx") {
Add-Content "$PSScriptRoot\Check-References-Report.html" "
<div class=""specification"">
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
<div class=""specification"">
<table style=""width:80%"">
<tr>
<td colspan=""5"" id=""tableHeader""><h2>$curSpc</h2></td>
</tr>" -Encoding UTF8
#========Statistics========
Get-DataFromSpecification -selectedFolder $script:PathToFilesBeingPublished -currentSPCName $_.Name
}
}
}
#========Statistics========
Add-Content "$PSScriptRoot\Check-References-Report.html" "
</div>
</body>
<script>
function ErrorCounter() {
var counter = 0;
var divs = document.getElementsByClassName('specification');
    for (var i=0; i<divs.length; i++) {
        if (divs[i].innerHTML.toLowerCase().indexOf('red') != -1) {
        counter = counter + 1;
        }
    }
return counter;
}
var errors = ErrorCounter();
document.getElementById('errorsfound').innerHTML = 'Ошибок найдено: ' + errors;
var specifications = document.getElementsByClassName('specification').length;
document.getElementById('spcchecked').innerHTML = 'Спецификаций проверено: ' + specifications;
</script>
</html>" -Encoding UTF8
#========Statistics========
}

$StringForHTML = "<h3>Анализ</h3>`r`n<span id=""spcchecked"">Спецификаций проверено:</span>`r`n<br>`r`n<br>`r`n<span>Всего ссылок на документы: $script:ReferencesToDocuments</span>`r`n<br>`r`n<br>`r`n<span>Всего ссылок на программы: $script:ReferencesToFiles</span>`r`n<br>`r`n<br>`r`n<span id=""errorsfound"">Ошибок найдено:</span>`r`n<h3>Результаты проверки</h3>`r`n<span><button type=""button"" onclick=""filterErrors()"" id=""filterbutton"">Показать только спецификации с ошибками</button></span>"
(Get-Content -Path "$PSScriptRoot\Check-References-Report.html").Replace("<h3>Анализ</h3>", $StringForHTML) | Set-Content("$PSScriptRoot\Check-References-Report.html") -Encoding UTF8
Start-Sleep -Seconds 1
Add-ExecutionTimeToReport -Time $ExecutionTime -ReportName "Check-References-Report" -StringToReplace "<h3>Анализ</h3>"
Invoke-Item "$PSScriptRoot\Check-References-Report.html"
Write-Host $script:ItemsReferencedTo

Function Check-Consistency() {
#Processes documents in $script:PathToFilesBeingPublished
Get-ChildItem -Path "$script:PathToFilesBeingPublished\*.*" -Include "*.pdf", "*.xls*", "*.doc*"
#Processes documents in $script:PathToPublishedFiles
Get-ChildItem -Path "$script:PathToPublishedFiles\*.*" -Include "*.pdf", "*.xls*", "*.doc*"
#Processes software in $script:PathToFilesBeingPublished
Get-ChildItem -Path "$script:PathToFilesBeingPublished\*.*" -Exclude "*.pdf", "*.xls*", "*.doc*"
#Processes software in $script:PathToPublishedFiles
Get-ChildItem -Path "$script:PathToPublishedFiles\*.*" -Exclude "*.pdf", "*.xls*", "*.doc*"
}
