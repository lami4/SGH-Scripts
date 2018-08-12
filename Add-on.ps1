Function Create-HtmlReportForErrors ([array]$Errors) 
{
Add-Content "$PSScriptRoot\Ошибки.html" "<!DOCTYPE html>
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
.Item_Number {
    width: 5%;
    text-align: left;
}
.Details {
    width:95%;
    text-align: left;
}
</style>
</head>
<body>
<div>
<h4>Следующие ошибки должны быть исправлены перед применением изменений по ИИ:</h4>
<div>
</div>
<table style=""width:100%"">
    <tr>
        <th class=""Item_Number"">№</th>
        <th class=""Details"">Описание</th>
    </tr>" -Encoding UTF8
$ErrorCounter = 1
$Errors | % {
    Add-Content "$PSScriptRoot\Ошибки.html" "    <tr>
        <td class=""Item_Number"">$($ErrorCounter)</th>
        <td class=""Details"">$($_)</th>
    </tr>" -Encoding UTF8
    $ErrorCounter += 1
}

Add-Content "$PSScriptRoot\Ошибки.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
}

Function Apply-Changes ($BackupFlag)
{
#Замена файлов
Write-Host "Выполняю замену файлов..."
    $ListViewReplace.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            $MyEnumeration = Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*"
            $MyEnumeration | % {Write-Host $_}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            $MyEnumeration = Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)"
            $MyEnumeration | % {Write-Host $_}
        }
        <##Если файл является документом, то выполняется алгоритм ниже
        if ($_.SubItems[2].Text -eq "Документ") {
            #Проверяет, существует ли документ(ы) с указанным обозначением в папке с текущей версией проекта
            if (Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*") {
                #Если файл существует, то выполняется код ниже:
                #Write-Host "Файл существует!"
                #Если пользователь выбрал выполнение резервного копирования
                if ($BackupFlag -eq $true) {
                    #Выбирает все файлы с указанным обозначением в папке с текущей версией проекта и копирует их в указанную резервную папку
                    Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*" | % {
                        #Скрипт выполняет попытку копирования. Если копируемый файл уже содержится в папке назначения (резервной), то скрипт предлагает перезаписать его
                        #и в зависимости от выбора пользователя добавляет соответсвующую запись в отчет
                        $ExceptionCheck = try {
                            [System.IO.File]::Copy("$($_.FullName)", "$script:MakeChangesPathToBackupFolder\$($_.Name)", $false)
                            #!!!Добавить соответсвующую запись в отчет
                        } catch [Exception] {
                            $_.Exception.GetType().FullName
                        }
                        #Если копируемый файл уже существует в резервной папке, то предлагается его перезаписать
                        if ($ExceptionCheck  -eq 'System.IO.IOException') {
                            if ((Show-MessageBox -Message "Копируемый файл $($_.Name) уже существует в папке для резервного копирования!`r`n`r`nПерезаписать?" -Title "Файл уже существует" -Type YesNo) -eq 'Yes') {
                                #Пользователь перезаписывает файл
                                [System.IO.File]::Copy("$($_.FullName)", "$script:MakeChangesPathToBackupFolder\$($_.Name)", $true)
                                #!!!Добавить соответсвующую запись в отчет
                            } else {
                                #Пользователь отказался перезаписывать файл
                                #!!!Добавить соответсвующую запись в отчет
                            }
                        }
                    }
                }
                #Перемещает папку в архив
                $ExceptionCheck = try {
                [System.IO.File]::Move("$($_.FullName)", "$script:MakeChangesPathToBackupFolder\$($_.Name)")
                #!!!Добавить соответсвующую запись в отчет
                } catch [Exception] {
                    $_.Exception.GetType().FullName
                }
                #Если перещаемый файл уже существует в архивной папке, то файл не перезаписывается и в отчет добавляется соответствующая записьдобавляется соответ
                if ($ExceptionCheck  -eq 'System.IO.IOException') {
                    Show-MessageBox -Message "Перемещаемый файл $($_.Name) уже существует в архивной папке!`r`n`r`nФайл не будет перемещен. В отчет добавлена соответсвующая запись." -Title "Файл уже существует" -Type OK
                    #!!!Добавить соответсвующую запись в отчет
                    #Активируется флаг и публикуемый файл НЕ копируется в текущую папку (с соответсвующей записью).
                }
                #Копирует публикуемый файл в папку с текущей версией проекта
            } else {
                #Write-Host "Файл НЕ существует!"
                #В отчет добавляется ошибка о том, что файл не существует
            }
        #Если файл является программой, то выполняется алгоритм ниже
        } else {
            #Проверяет, существует ли программа в папке с текущей версией
            if (Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)") {
                #Write-Host "Файл существует!"
            } else {
                #Write-Host "Файл НЕ существует!"
            }
        }
        Write-Host $_.Text
    }
    #Test-Path -Path#>
    }
}

Function Check-Conditions ($BackupFlag)
{
$ErrorDetectedFlag = $false
$ErrorsToBePublished = @()
#ПРОВЕРКИ
#Папка для резервной копии пуста
if ($BackupFlag -eq $true) {
    if ((Get-ChildItem -Path $script:MakeChangesPathToBackupFolder).Count -gt 0) {$ErrorsToBePublished += "Папка для резервного копирования содержит файлы. Данная папка должна быть пустой перед началом операции."; $ErrorDetectedFlag = $true}
}
#Прхивная папка пуста
if ((Get-ChildItem -Path $script:MakeChangesPathToArchiveFolder).Count -gt 0) {$ErrorsToBePublished += "Архивная папка содержит файлы. Данная папка должна быть пустой перед началом операции."; $ErrorDetectedFlag = $true}
#Повторяющиеся значения
    $ItemsInTheLists = @()
    $CheckedItems = @()
    $ListViewAdd.Items | % {$ItemsInTheLists += $_.Text}
    $ListViewReplace.Items | % {$ItemsInTheLists += $_.Text}
    $ListViewRemove.Items | % {$ItemsInTheLists += $_.Text}
    $ItemsInTheLists | % {if ($CheckedItems -contains $_) {$ErrorsToBePublished += "Обозначение <i>$($_)</i> используется сразу несколькими записями в списках. Каждая запись в списках должна иметь уникальное обозначение."; $ErrorDetectedFlag = $true} else {$CheckedItems += $_}}
#Проверки для списка Выпустить.
#Все файлы из списка присутствуют в папке с публикуемыми файлами?
$ListViewAdd.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text).*").Count -eq 0) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text)").Count -eq 0) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }  
}
#Присутствуют ли файлы из списка в папке с текущей версией проекта?
$ListViewAdd.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*").Count -gt 0) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но папка с текущей версией проекта уже содержит данный документ."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)").Count -gt 0) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но папка с текущей версией проекта уже содержит данную программу."; $ErrorDetectedFlag = $true}
        }  
}
#Проверки для списка Заменить.
#Все файлы из списка присутствуют в папке с публикуемыми файлами?
$ListViewReplace.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text).*").Count -eq 0) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text)").Count -eq 0) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }  
}
#Присутствуют ли файлы из списка в папке с текущей версией проекта?
$ListViewReplace.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*").Count -eq 0) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)").Count -eq 0) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }  
}
#Проверки для списка Аннулировать.
#Присутствуют ли файлы из списка в папке с текущей версией проекта?
$ListViewRemove.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*").Count -eq 0) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Аннулировать</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)").Count -eq 0) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Аннулировать</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }  
}
#ПРОВЕРКИ
if ($ErrorDetectedFlag -eq $true) {Create-HtmlReportForErrors -Errors $ErrorsToBePublished}
}

Function ApplyChangesForm ()
{
    $ApplyChangesForm = New-Object System.Windows.Forms.Form
    $ApplyChangesForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ApplyChangesForm.ShowIcon = $false
    $ApplyChangesForm.AutoSize = $true
    $ApplyChangesForm.Text = "Внести изменения"
    $ApplyChangesForm.AutoSizeMode = "GrowAndShrink"
    $ApplyChangesForm.WindowState = "Normal"
    $ApplyChangesForm.SizeGripStyle = "Hide"
    $ApplyChangesForm.ShowInTaskbar = $true
    $ApplyChangesForm.StartPosition = "CenterScreen"
    $ApplyChangesForm.MinimizeBox = $false
    $ApplyChangesForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор
    $ApplyChangesFormFilesBeingPublished = New-Object System.Windows.Forms.Button
    $ApplyChangesFormFilesBeingPublished.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ApplyChangesFormFilesBeingPublished.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormFilesBeingPublished.Text = "Обзор..."
    $ApplyChangesFormFilesBeingPublished.TabStop = $false
    $ApplyChangesFormFilesBeingPublished.Add_Click({
        $script:MakeChangesPathToFilesBeingPublished = Select-Folder -Description "Укажите папку с публикуемыми файлами."
        if ($script:MakeChangesPathToFilesBeingPublished -ne $null) {
            $ApplyChangesFormFilesBeingPublishedLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToFilesBeingPublished -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesFormFilesBeingPublishedLabel, $script:MakeChangesPathToFilesBeingPublished)
            Write-Host $script:MakeChangesPathToFilesBeingPublished
        }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormFilesBeingPublished)
    #Поле к кнопке Обзор
    $ApplyChangesFormFilesBeingPublishedLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesFormFilesBeingPublishedLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $ApplyChangesFormFilesBeingPublishedLabel.Width = 500
    $ApplyChangesFormFilesBeingPublishedLabel.Text = "Укажите папку с публикуемыми файлами"
    $ApplyChangesFormFilesBeingPublishedLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesFormFilesBeingPublishedLabel)
    #Кнопка обзор
    $ApplyChangesFormCurrentVersion = New-Object System.Windows.Forms.Button
    $ApplyChangesFormCurrentVersion.Location = New-Object System.Drawing.Point(10,42) #x,y
    $ApplyChangesFormCurrentVersion.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormCurrentVersion.Text = "Обзор..."
    $ApplyChangesFormCurrentVersion.TabStop = $false
    $ApplyChangesFormCurrentVersion.Add_Click({
        $script:MakeChangesPathToCurrentVersion = Select-Folder -Description "Укажите папку с текущей версией проекта."
        if ($script:MakeChangesPathToCurrentVersion -ne $null) {
            $ApplyChangesFormCurrentVersionLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToCurrentVersion -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesFormCurrentVersionLabel, $script:MakeChangesPathToCurrentVersion)
            Write-Host $script:MakeChangesPathToCurrentVersion
        } 
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCurrentVersion)
    #Поле к кнопке Обзор
    $ApplyChangesFormCurrentVersionLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesFormCurrentVersionLabel.Location =  New-Object System.Drawing.Point(95,46) #x,y
    $ApplyChangesFormCurrentVersionLabel.Width = 500
    $ApplyChangesFormCurrentVersionLabel.Text = "Укажите папку с текущей версией проекта"
    $ApplyChangesFormCurrentVersionLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCurrentVersionLabel)
    #Кнопка обзор
    $ApplyChangesArchiveFolder = New-Object System.Windows.Forms.Button
    $ApplyChangesArchiveFolder.Location = New-Object System.Drawing.Point(10,74) #x,y
    $ApplyChangesArchiveFolder.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesArchiveFolder.Text = "Обзор..."
    $ApplyChangesArchiveFolder.TabStop = $false
    $ApplyChangesArchiveFolder.Add_Click({
        $script:MakeChangesPathToArchiveFolder = Select-Folder -Description "Укажите архивную папку для аннулируемых и заменяемых файлов."
        if ($script:MakeChangesPathToArchiveFolder -ne $null) {
            $ApplyChangesArchiveFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToArchiveFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesArchiveFolderLabel, $script:MakeChangesPathToArchiveFolder)
            Write-Host $script:MakeChangesPathToArchiveFolder
        }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesArchiveFolder)
    #Поле к кнопке Обзор
    $ApplyChangesArchiveFolderLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesArchiveFolderLabel.Location =  New-Object System.Drawing.Point(95,78) #x,y
    $ApplyChangesArchiveFolderLabel.Width = 500
    $ApplyChangesArchiveFolderLabel.Text = "Укажите архивную папку для аннулируемых и заменяемых файлов"
    $ApplyChangesArchiveFolderLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesArchiveFolderLabel)
    #Чекбокс 'Выполнить резервное копирование аннулируемых и заменяемых файлов'
    $ApplyChangesFormMakeBackupCopy = New-Object System.Windows.Forms.CheckBox
    $ApplyChangesFormMakeBackupCopy.Width = 518
    $ApplyChangesFormMakeBackupCopy.Text = "Выполнить резервное копирование аннулируемых и заменяемых файлов перед архивацией"
    $ApplyChangesFormMakeBackupCopy.Location = New-Object System.Drawing.Point(10,106) #x,y
    $ApplyChangesFormMakeBackupCopy.Enabled = $true
    $ApplyChangesFormMakeBackupCopy.Checked = $true
    $ApplyChangesFormMakeBackupCopy.Add_CheckStateChanged({
    if ($ApplyChangesFormMakeBackupCopy.Checked -eq $false) {
    $ApplyChangesFormBackup.Enabled = $false
    $ApplyChangesFormBackupLabel.Enabled = $false
    } else {
    $ApplyChangesFormBackup.Enabled = $true
    $ApplyChangesFormBackupLabel.Enabled = $true 
    }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormMakeBackupCopy)
    #Кнопка обзор
    $ApplyChangesFormBackup = New-Object System.Windows.Forms.Button
    $ApplyChangesFormBackup.Location = New-Object System.Drawing.Point(10,138) #x,y
    $ApplyChangesFormBackup.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormBackup.Text = "Обзор..."
    $ApplyChangesFormBackup.TabStop = $false
    $ApplyChangesFormBackup.Add_Click({
        $script:MakeChangesPathToBackupFolder = Select-Folder -Description "Укажите архивную папку для аннулируемых и заменяемых файлов."
        if ($script:MakeChangesPathToBackupFolder -ne $null) {
            $ApplyChangesFormBackupLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToBackupFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesFormBackupLabel, $script:MakeChangesPathToBackupFolder)
            Write-Host $script:MakeChangesPathToBackupFolder
        }  
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormBackup)
    #Поле к кнопке Обзор
    $ApplyChangesFormBackupLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesFormBackupLabel.Location =  New-Object System.Drawing.Point(95,142) #x,y
    $ApplyChangesFormBackupLabel.Width = 500
    $ApplyChangesFormBackupLabel.Text = "Укажите папку для резервной копии аннулируемых и заменяемых файлов"
    $ApplyChangesFormBackupLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesFormBackupLabel)
    #Обновление полей
    if ($script:MakeChangesPathToFilesBeingPublished -ne $null) {
        $ApplyChangesFormFilesBeingPublishedLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToFilesBeingPublished -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ApplyChangesFormFilesBeingPublishedLabel, $script:MakeChangesPathToFilesBeingPublished)
    }
    if ($script:MakeChangesPathToCurrentVersion -ne $null) {
        $ApplyChangesFormCurrentVersionLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToCurrentVersion -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ApplyChangesFormCurrentVersionLabel, $script:MakeChangesPathToCurrentVersion)
    }
    if ($script:MakeChangesPathToArchiveFolder -ne $null) {
        $ApplyChangesArchiveFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToArchiveFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ApplyChangesArchiveFolderLabel, $script:MakeChangesPathToArchiveFolder)
    }
    if ($script:MakeChangesPathToBackupFolder -ne $null) {
        $ApplyChangesFormBackupLabel.Text = "Указанная папка: $(Split-Path -Path $script:MakeChangesPathToBackupFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ApplyChangesFormBackupLabel, $script:MakeChangesPathToBackupFolder)
    }
    #Проверка условий
    $ApplyChangesFormCheckConditions = New-Object System.Windows.Forms.Button
    $ApplyChangesFormCheckConditions.Location = New-Object System.Drawing.Point(10,190) #x,y
    $ApplyChangesFormCheckConditions.Size = New-Object System.Drawing.Point(130,22) #width,height
    $ApplyChangesFormCheckConditions.Text = "Проверить условия"
    $ApplyChangesFormCheckConditions.Enabled = $true
    $ApplyChangesFormCheckConditions.Add_Click({
        $MakeChangesBackupFlag = $true
        if ($ApplyChangesFormMakeBackupCopy.Checked -eq $true) {$MakeChangesBackupFlag = $true} else {$MakeChangesBackupFlag = $false}
        if ($script:MakeChangesPathToFilesBeingPublished -eq $null -or $script:MakeChangesPathToCurrentVersion -eq $null -or $script:MakeChangesPathToArchiveFolder -eq $null -or ($script:MakeChangesPathToBackupFolder -eq $null -and $ApplyChangesFormMakeBackupCopy.Checked -eq $true)) {
                Show-MessageBox -Message "Не указан путь к одной или нескольким папкам." -Title "Невозможно выполнить операцию" -Type OK
            } else {
                if ($script:MakeChangesPathToFilesBeingPublished -eq $script:MakeChangesPathToCurrentVersion -or`
                $script:MakeChangesPathToFilesBeingPublished -eq $script:MakeChangesPathToArchiveFolder -or`
                $script:MakeChangesPathToFilesBeingPublished -eq $script:MakeChangesPathToBackupFolder -or`
                $script:MakeChangesPathToCurrentVersion -eq $script:MakeChangesPathToArchiveFolder -or`
                $script:MakeChangesPathToCurrentVersion -eq $script:MakeChangesPathToBackupFolder -or`
                $script:MakeChangesPathToArchiveFolder -eq $script:MakeChangesPathToBackupFolder) {
                    Show-MessageBox -Message "Один и тот же путь указан для двух разных папок." -Title "Невозможно выполнить операцию" -Type OK
                } else {
                    Check-Conditions -BackupFlag $MakeChangesBackupFlag
                }
            }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCheckConditions)
    #Кнопка Начать
    $ApplyChangesFormApplyButton = New-Object System.Windows.Forms.Button
    $ApplyChangesFormApplyButton.Location = New-Object System.Drawing.Point(150,190) #x,y
    $ApplyChangesFormApplyButton.Size = New-Object System.Drawing.Point(130,22) #width,height
    $ApplyChangesFormApplyButton.Text = "Внести изменения"
    $ApplyChangesFormApplyButton.Enabled = $false
    $ApplyChangesFormApplyButton.Add_Click({
        $MakeChangesBackupFlag = $true
        if ($ApplyChangesFormMakeBackupCopy.Checked -ne $true) {$MakeChangesBackupFlag = $false}
        Apply-Changes -BackupFlag $MakeChangesBackupFlag
        $ApplyChangesForm.Close()
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormApplyButton)
    #Кнопка закрыть
    $ApplyChangesFormCancelButton = New-Object System.Windows.Forms.Button
    $ApplyChangesFormCancelButton.Location = New-Object System.Drawing.Point(290,190) #x,y
    $ApplyChangesFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormCancelButton.Text = "Закрыть"
    $ApplyChangesFormCancelButton.Add_Click({
        $ApplyChangesForm.Close()
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCancelButton)
    $ApplyChangesForm.ShowDialog()
}
