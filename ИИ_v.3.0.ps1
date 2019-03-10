clear
#Глобальные переменные
#Переменная для стандартного списка рассылки
$script:SendTo = 'Стандартный список рассылки'
#Формат файлов
$script:FileFormats = @('sh;SH-файл', 'zip;ZIP-архив', 'gz;GZ-файл')
#Значение для поля 'Причина'
$script:GlobalReasonField = "Обновление документации"
#Значение для поля 'Указание о заделе'
$script:GlobalInStoreField = "Задела нет"
#Значение для поля 'Указание о внедрении'
$script:GlobalStartUsageField = "С момента выпуска"
#Значение для поля 'Применяемость'
$script:GlobalApplicableToField = "---"
#Значение для поля 'Разослать'
$script:GlobalSendToField = "По списку рассылки"
#Значение для поля 'Приложение'
$script:GlobalAppendixField = "Нет"

#Служебные переменные
$script:VerNumber = ""
$script:SuspiciousAction = $false
$script:IncorrectVersionDiscrepancy = $false
$script:CurrentVersionDocumentExists = $false
$script:HighlightChecboxStatus = $true
$script:PathToCurrentVrsion = $null
$script:BannedCharacters = '\/|\\|\?|%|\*|:|\||<|>|"'
$script:MakeChangesPathToFilesBeingPublished = $null
$script:MakeChangesPathToCurrentVersion = $null
$script:MakeChangesPathToArchiveFolder = $null
$script:MakeChangesPathToBackupFolder = $null
$script:PathToRegister = $null
$script:PathToRegisterColoring = $null
$script:SelectedRegister = $null
$script:SelectedFolderWithFilesBeingPublished = $null
$script:ManuallyEnteredValueForRegister = ""
$script:CollectedReferences = @(), @(), @()
$script:SelectedWordFile = $null
$script:SelectedClientFolder = $null
$script:SelectedAccessPath = $null
$script:AggregatingString = ""

Function Collect-DataFromSpecification ($WordApp, $PathToSpecification)
{
    $Specification = $WordApp.Documents.Open("$PathToSpecification")
    $Specification.Tables.Item(1).Rows | % {
        if ($_.Cells.Count -eq 7) {
            if (((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ') -replace [char]13, '') -ne '') {
                if ($script:CollectedReferences[0] -cnotcontains [string]((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '').Trim([char]0x0009)) {
                    $script:CollectedReferences[0] += [string]((($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '').Trim([char]0x0009).Trim(' ')
                    $script:CollectedReferences[1] += [string]((($_.Cells.Item(5).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '').Trim([char]0x0009).Trim(' ')
                    $script:CollectedReferences[2] += [string]((($_.Cells.Item(1).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, '').Trim([char]0x0009).Trim(' ')
                }         
            }
        }
    }
    $script:CollectedReferences[0] += [string]((($Specification.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text)).Trim(' ')  -replace [char]13, '').Trim([char]0x0007).Trim([char]0x0009).Trim(' ')
    $script:CollectedReferences[1] += [string]((($Specification.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text)).Trim(' ')  -replace [char]13, '').Trim([char]0x0007).Trim([char]0x0009).Trim(' ')
    $script:CollectedReferences[2] += "А4"
    $Specification.Close([ref]0)
}

Function Enter-DataToRegisterManually ($Title, $Label)
{
    $EnterDataToRegisterManuallyForm = New-Object System.Windows.Forms.Form
    $EnterDataToRegisterManuallyForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $EnterDataToRegisterManuallyForm.ShowIcon = $false
    $EnterDataToRegisterManuallyForm.AutoSize = $true
    $EnterDataToRegisterManuallyForm.Text = "$Title"
    $EnterDataToRegisterManuallyForm.AutoSizeMode = "GrowAndShrink"
    $EnterDataToRegisterManuallyForm.WindowState = "Normal"
    $EnterDataToRegisterManuallyForm.SizeGripStyle = "Hide"
    $EnterDataToRegisterManuallyForm.ShowInTaskbar = $true
    $EnterDataToRegisterManuallyForm.StartPosition = "CenterScreen"
    $EnterDataToRegisterManuallyForm.ControlBox = $false
    $EnterDataToRegisterManuallyForm.MinimizeBox = $false
    $EnterDataToRegisterManuallyForm.MaximizeBox = $false
    #Надпись к полю для ввода
    $EnterDataToRegisterManuallyFormInputFieldLabel = New-Object System.Windows.Forms.Label
    $EnterDataToRegisterManuallyFormInputFieldLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $EnterDataToRegisterManuallyFormInputFieldLabel.Width = 600
    $EnterDataToRegisterManuallyFormInputFieldLabel.Height = 16
    $EnterDataToRegisterManuallyFormInputFieldLabel.Text = "$Label"
    $EnterDataToRegisterManuallyFormInputFieldLabel.TextAlign = "TopLeft"
    $EnterDataToRegisterManuallyForm.Controls.Add($EnterDataToRegisterManuallyFormInputFieldLabel)
    #Поле для ввода
    $EnterDataToRegisterManuallyFormInputField = New-Object System.Windows.Forms.TextBox 
    $EnterDataToRegisterManuallyFormInputField.Location = New-Object System.Drawing.Point(10,35) #x,y
    $EnterDataToRegisterManuallyFormInputField.Width = 800
    $EnterDataToRegisterManuallyFormInputField.ForeColor = "Black"
    $EnterDataToRegisterManuallyForm.Controls.Add($EnterDataToRegisterManuallyFormInputField)
    #Кнопка применить
    $EnterDataToRegisterManuallyFormApplyButton = New-Object System.Windows.Forms.Button
    $EnterDataToRegisterManuallyFormApplyButton.Location = New-Object System.Drawing.Point(10,75) #x,y
    $EnterDataToRegisterManuallyFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EnterDataToRegisterManuallyFormApplyButton.Text = "Применить"
    $EnterDataToRegisterManuallyFormApplyButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $EnterDataToRegisterManuallyFormApplyButton.Add_Click({
    $script:ManuallyEnteredValueForRegister = $EnterDataToRegisterManuallyFormInputField.Text
    $EnterDataToRegisterManuallyForm.Close()
    })
    $EnterDataToRegisterManuallyForm.Controls.Add($EnterDataToRegisterManuallyFormApplyButton)
    $EnterDataToRegisterManuallyForm.ShowDialog()
}

Function Find-StringToBePopulated ($Sheet, $LookFor, $ColumnLetter)
{
    $InstanceCounter = @()
    $Range = $Sheet.Range("$($ColumnLetter):$($ColumnLetter)")
    $Target = $Range.Find("$LookFor", [Type]::Missing, [Type]::Missing, 1)
    if ($Target -eq $null) {
        #Обозначение заявлено к замене, но не существует в файле учета нет ни одного вхождения.
        Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($LookFor): ОШИБКА. Требуется ручное внесение данных. Обозначение заявлено к замене/аннулированию, но файле учета ПД и программ не существует ни одного вхождения с таким обозначением."
        return "not exist"
    } else {
        #Ищет
        $FirstHit = $Target
        Do
        {
            $FoundNameRowNumber = $Target.AddressLocal($false, $false) -replace "$($ColumnLetter)", ""
            if ($Sheet.Cells.Item($FoundNameRowNumber, "$($ColumnLetter)").Interior.ColorIndex -eq -4142 -and $Sheet.Cells.Item($FoundNameRowNumber, "L").Value() -eq $null -and $Sheet.Cells.Item($FoundNameRowNumber, "M").Value() -eq $null -and $Sheet.Cells.Item($FoundNameRowNumber, "N").Value() -eq $null) {
            $InstanceCounter += $FoundNameRowNumber
            }
            $Target = $Range.FindNext($Target)
        }
        While ($Target -ne $null -and $Target.AddressLocal() -ne $FirstHit.AddressLocal())
    }
    if ($InstanceCounter.Count -eq 0) {
        Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($LookFor): ОШИБКА. Требуется ручное внесение данных. В файле учета ПД и программ не удалось обнаружить ни одной строки для заменяемого/аннулируемого файла, которая подходит для автозаполнения. К нужной строке либо по ошибке применили заливку, либо в ней по ошибке заполнены столбцы L, M или N."
        return "not found"
        #Write-Host "Не удалось найти строку подходящую для заполнения. К строке либо по ошибке прмиенили заливку, либо в ней по ошибке заполнены столбцы L, M или N."
    } elseif ($InstanceCounter.Count -eq 1) {
        Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($LookFor): УСПЕШНО. В файл учета успешно внесены все необходимые данные."
        return [int]$InstanceCounter[0]
        #Write-Host "Найдена строка без заливки и без начений в столбцах L, M и N."
    } else {
        $RegisterStringForReport = ""
        $InstanceCounter | % {$RegisterStringForReport += "$([string]$_), "}
        Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($LookFor): ОШИБКА. Требуется ручное внесение данных. В файле учета ПД и программ найдено несколько строк подходящих для автозаполнения. Номера подходящих строк: $($RegisterStringForReport.Trim(', '))"
        return "multiple instances"
        #Write-Host "Найдено несколько строк подходящих для заполнения. Их номера:"
    }
}

Function Populate-Register ()
{
    if (Test-Path -Path "$PSScriptRoot\Журнал автозаполнения.txt") {Remove-Item -Path "$PSScriptRoot\Журнал автозаполнения.txt"}
    Kill -Name WINWORD, EXCEL -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    #СОБИРАЕМ ДАННЫЕ ИЗ СПЕЦИФИКАЦИЙ В УКАЗАННОЙ ПАПКЕ
    #Создать экземпляр приложения MS Word
    $RegisterWordApp = New-Object -ComObject Word.Application
    #Сделать вызванное приложение невидемым
    $RegisterWordApp.Visible = $false
    Write-Host "Сбор информации о документах и программах, указанной в публикуемых спецификациях..."
    #DOCX
    Get-ChildItem -Path "$script:SelectedFolderWithFilesBeingPublished\*.docx" | % {
        if ($_.BaseName -match 'SPC' -or $_.BaseName -match 'LPD') {
            Write-Host $_.Name
            Collect-DataFromSpecification -Word $RegisterWordApp -PathToSpecification $_
        }
    }
    #DOC
    Get-ChildItem -Path "$script:SelectedFolderWithFilesBeingPublished\*.doc" | % {
        if ($_.BaseName -match 'SPC' -or $_.BaseName -match 'LPD') {
            Write-Host $_.Name
            Collect-DataFromSpecification -Word $RegisterWordApp -PathToSpecification $_
        }
    }
    $RegisterWordApp.Quit()
    Kill -Name WINWORD -ErrorAction SilentlyContinue
    Write-Host $script:CollectedReferences
    Write-Host "Сбор информации закончен."
    Write-Host "Заполняю файл учета программ и ПД..."
    Start-Sleep -Seconds 2
    $WordApplication = New-Object -ComObject word.application
    $WordApplication.Visible = $false
    $Register = New-Object -ComObject Excel.Application
    $Register.Visible = $true
    $RegisterWorkbook = $Register.WorkBooks.Open($script:SelectedRegister)
    $RegisterWorksheet = $RegisterWorkbook.Worksheets.Item(1)
    if ($RegisterWorksheet.AutoFilterMode -eq $true) {$RegisterWorksheet.ShowAllData()}
    $ArrayOfRowNumbers = @()
    $ArrayOfRowNumbers += $RegisterWorksheet.Cells.Item($RegisterWorksheet.Rows.Count, "E").End(-4162).Row
    $ArrayOfRowNumbers += $RegisterWorksheet.Cells.Item($RegisterWorksheet.Rows.Count, "C").End(-4162).Row
    $ArrayOfRowNumbers += $RegisterWorksheet.Cells.Item($RegisterWorksheet.Rows.Count, "F").End(-4162).Row
    $ArrayOfRowNumbers += $RegisterWorksheet.Cells.Item($RegisterWorksheet.Rows.Count, "A").End(-4162).Row
    $RegisterLastRow = [int]($ArrayOfRowNumbers | Measure -Maximum).Maximum
    #Список ВЫПУСТИТЬ
    Write-Host "======================================="
    Write-Host "Работую со списком Выпустить..."
    $ListViewAdd.Items | % {
        $ArrayContainsExtensionFlag = $false
        $ArrayContainsExtensionIndex = $null
        $RegisterLastRow += 1
        #Программа
        if ($_.SubItems[2].Text -eq "Программа") {
            Write-Host "$($_.Text)"
            #КОД ПРОЕКТА
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 1) = $UpdateRegisterFormComboboxProjectName.SelectedItem
            #РАЗРАБОТЧИК
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 2) = $UpdateRegisterFormComboboxDeveloperName.SelectedItem
            #ФАЙЛ ПРОГРАММЫ 
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 3) = $_.Text
            #КОНТРОЛЬНАЯ СУММА
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 4) = [string]($_.SubItems[1].Text).ToUpper()
            #НАИМЕНОВАНИЕ
            if ($script:CollectedReferences[0].Contains("$($_.Text)")) {
                if ($script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                    $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
                } else {
                    Enter-DataToRegisterManually -Title 'Программа упоминается в спекицикации(ях), но для нее не указано наименование' -Label "Укажите наименование для программы $($_.Text):"
                    $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                } 
            } else {
                Enter-DataToRegisterManually -Title 'Программа не упоминается ни в одной из спецификаций' -Label "Укажите наименование для программы $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            }
            #ФОРМАТ
            for ($i = 0; $i -lt $script:FileFormats.Count; $i++) {
                if ((($script:FileFormats[$i] -split ';')[0]).ToLower() -eq ([System.IO.Path]::GetExtension($_.Text)).Trim('.').ToLower()) {$ArrayContainsExtensionFlag = $true; $ArrayContainsExtensionIndex = $i}
            }
            if ($ArrayContainsExtensionFlag -eq $true) {
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 7) = ($script:FileFormats[$ArrayContainsExtensionIndex] -split ';')[1]
            } else {
                Enter-DataToRegisterManually -Title 'Формат программы не может быть заполнен автоматически, так как отсутсвует в списке настроенных форматов' -Label "Укажите формат для программы $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 7) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            }
            #РАЗМЕР ФАЙЛА
            Get-ChildItem -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text)" | % {
                if ($_.Length -lt 1048576) {$RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = "$([math]::Round($_.Length/1KB, 2))" + " КБ " + "($($_.Length)" + " байт)"}
                if ($_.Length -gt 1048576 -and $_.Length -lt 1073741824) {$RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = "$([math]::Round($_.Length/1MB, 2))" + " МБ " + "($($_.Length)" + " байт)"}
                if ($_.Length -gt 1073741824) {$RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = "$([math]::Round($_.Length/1GB, 2))" + " ГБ " + "($($_.Length)" + " байт)"}
            }
            #ДАТА ПОСТУПЛЕНИЯ
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 9) = $CalendarIssueDateInput.Text
            Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($_.Text): УСПЕШНО. В файл учета успешно внесены все необходимые данные."
            $RegisterWorksheet.Rows.Item($RegisterLastRow).Cells.Item(1).Interior.Color = 139
        }
        #Документ
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "$($_.Text)"
            #КОД ПРОЕКТА
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 1) = $UpdateRegisterFormComboboxProjectName.SelectedItem
            #РАЗРАБОТЧИК
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 2) = $UpdateRegisterFormComboboxDeveloperName.SelectedItem
            #ОБОЗНАЧЕНИЕ
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 5) = $_.Text
            #НАИМЕНОВАНИЕ
            if ($script:CollectedReferences[0].Contains("$($_.Text)")) {
                if ($script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                    $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
                } else {
                    Enter-DataToRegisterManually -Title 'Документ упоминается в спекицикации(ях), но для него не указано наименование' -Label "Укажите наименование для документа $($_.Text):"
                    $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                }
            } else {
                Enter-DataToRegisterManually -Title 'Документ не упоминается ни в одной из спецификаций' -Label "Укажите наименование для документа $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 6) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            }
            #ФОРМАТ
            if ($script:CollectedReferences[2][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 7) = $script:CollectedReferences[2][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
            } else {
                Enter-DataToRegisterManually -Title 'Формат для документа не указан ни в одной из спецификаций' -Label "Укажите формат для документа $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 7) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            }
            #КОЛИЧЕСТВО ЛИСТОВ
            #Гид по стилю
            if ($($_.Text) -match "DSG") {
                Enter-DataToRegisterManually -Title 'Документ с кодом DSG' -Label "Укажите общее количество страниц для документа $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            } elseif ((Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).xlsx") -or (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).xls")) {
            #Excel-файл
                Enter-DataToRegisterManually -Title 'Документ созданный в приложении MS Excel' -Label "Укажите общее количество страниц для документа $($_.Text):"
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = $script:ManuallyEnteredValueForRegister
                $script:ManuallyEnteredValueForRegister = ""
            #Остальные файлы, т.е. WORD-файлы
            } else {
                if (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).docx") {$DocumentRegister = $WordApplication.Documents.Open("$script:SelectedFolderWithFilesBeingPublished\$($_.Text).docx")}
                if (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).doc") {$DocumentRegister = $WordApplication.Documents.Open("$script:SelectedFolderWithFilesBeingPublished\$($_.Text).doc")}
                $Wholestory = $DocumentRegister.Range()
                $TotalPages = $Wholestory.Information(4)
                $RegisterWorksheet.Cells.Item($RegisterLastRow, 8) = $TotalPages
                $DocumentRegister.Close()
            }
            #ДАТА ПОСТУПЛЕНИЯ
            $RegisterWorksheet.Cells.Item($RegisterLastRow, 9) = $CalendarIssueDateInput.Text
            Add-Content "$PSScriptRoot\Журнал автозаполнения.txt" "$($_.Text): УСПЕШНО. В файл учета успешно внесены все необходимые данные."
            $RegisterWorksheet.Rows.Item($RegisterLastRow).Cells.Item(1).Interior.Color = 139
        }
    }
    #Список ЗАМЕНИТЬ
    Write-Host "======================================="
    Write-Host "Работую со списком Заменить..."
    $ListViewReplace.Items | % {
        $ArrayContainsExtensionFlag = $false
        $ArrayContainsExtensionIndex = $null
        $RegisterLastRow += 1
        #Программа
        if ($_.SubItems[2].Text -eq "Программа") {
            Write-Host "$($_.Text)"
            $ActionFlag = $null
            #Определяем есть ли строка заменяемого файла, подходящая для заполнения
            $ActionFlag = Find-StringToBePopulated -Sheet $RegisterWorksheet -LookFor $_.Text -ColumnLetter "C"
            #Write-Host "Значение ячейки $ActionFlag"
            #Если подходящая строка найдена, то заполняем ее и создаем новую строку для заменяющего файла под строкой для заменяемого файла
            if ($ActionFlag -ne 'not exist' -and $ActionFlag -ne 'not found' -and $ActionFlag -ne 'multiple instances') {
                #Заполняем строку заменяемого файла
                $RegisterWorksheet.Cells.Item($ActionFlag, 12) = $CalendarIssueDateInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 13) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).HorizontalAlignment = -4131
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 14) = "заменен"
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(1).Interior.Color = 139
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(2).Interior.Color = 16436871
                #Создаем и заполняем строку заменяющего файла
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Insert(-4121)
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Interior.Color = -4142
                #КОД ПРОЕКТА
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 1) = $UpdateRegisterFormComboboxProjectName.SelectedItem
                #РАЗРАБОТЧИК
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 2) = $UpdateRegisterFormComboboxDeveloperName.SelectedItem
                #ФАЙЛ ПРОГРАММЫ 
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 3) = $_.Text
                #КОНТРОЛЬНАЯ СУММА
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 4) = [string]($_.SubItems[1].Text).ToUpper()
                #НАИМЕНОВАНИЕ
                if ($script:CollectedReferences[0].Contains("$($_.Text)")) {
                    if ($script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                        $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
                    } else {
                        Enter-DataToRegisterManually -Title 'Программа упоминается в спекицикации(ях), но для нее не указано наименование' -Label "Укажите наименование для программы $($_.Text):"
                        $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:ManuallyEnteredValueForRegister
                        $script:ManuallyEnteredValueForRegister = ""
                    } 
                } else {
                    Enter-DataToRegisterManually -Title 'Программа не упоминается ни в одной из спецификаций' -Label "Укажите наименование для программы $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                }
                #ФОРМАТ
                for ($i = 0; $i -lt $script:FileFormats.Count; $i++) {
                    if ((($script:FileFormats[$i] -split ';')[0]).ToLower() -eq ([System.IO.Path]::GetExtension($_.Text)).Trim('.').ToLower()) {$ArrayContainsExtensionFlag = $true; $ArrayContainsExtensionIndex = $i}
                }
                if ($ArrayContainsExtensionFlag -eq $true) {
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 7) = ($script:FileFormats[$ArrayContainsExtensionIndex] -split ';')[1]
                } else {
                    Enter-DataToRegisterManually -Title 'Формат программы не может быть заполнен автоматически, так как отсутсвует в списке настроенных форматов' -Label "Укажите формат для программы $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 7) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                }
                #РАЗМЕР ФАЙЛА
                Get-ChildItem -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text)" | % {
                    if ($_.Length -lt 1048576) {$RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = "$([math]::Round($_.Length/1KB, 2))" + " КБ " + "($($_.Length)" + " байт)"}
                    if ($_.Length -gt 1048576 -and $_.Length -lt 1073741824) {$RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = "$([math]::Round($_.Length/1MB, 2))" + " МБ " + "($($_.Length)" + " байт)"}
                    if ($_.Length -gt 1073741824) {$RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = "$([math]::Round($_.Length/1GB, 2))" + " ГБ " + "($($_.Length)" + " байт)"}
                }
                #ДАТА ПОСТУПЛЕНИЯ
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 9) = $CalendarIssueDateInput.Text
                #№ ДОКУМ. ВВЕДЕН
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(11).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 11) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(11).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(1).Interior.Color = 139
            }       
        }
        #Документ
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "$($_.Text)"
            $ActionFlag = $null
            #Определяем есть ли строка заменяемого файла, подходящая для заполнения
            $ActionFlag = Find-StringToBePopulated -Sheet $RegisterWorksheet -LookFor $_.Text -ColumnLetter "E"
            #Write-Host "Значение ячейки $ActionFlag"
            #Если подходящая строка найдена, то заполняем ее и создаем новую строку для заменяющего файла под строкой для заменяемого файла
            if ($ActionFlag -ne 'not exist' -and $ActionFlag -ne 'not found' -and $ActionFlag -ne 'multiple instances') {
                #Заполняем строку заменяемого файла
                $RegisterWorksheet.Cells.Item($ActionFlag, 12) = $CalendarIssueDateInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 13) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).HorizontalAlignment = -4131
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 14) = "заменен"
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(1).Interior.Color = 139
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(2).Interior.Color = 16436871
                #Создаем и заполняем строку заменяющего файла
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Insert(-4121)
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Interior.Color = -4142
                #КОД ПРОЕКТА
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 1) = $UpdateRegisterFormComboboxProjectName.SelectedItem
                #РАЗРАБОТЧИК
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 2) = $UpdateRegisterFormComboboxDeveloperName.SelectedItem
                #ОБОЗНАЧЕНИЕ
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 5) = $_.Text
                #НАИМЕНОВАНИЕ
                if ($script:CollectedReferences[0].Contains("$($_.Text)")) {
                    if ($script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                        $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:CollectedReferences[1][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
                    } else {
                        Enter-DataToRegisterManually -Title 'Документ упоминается в спекицикации(ях), но для него не указано наименование' -Label "Укажите наименование для документа $($_.Text):"
                        $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:ManuallyEnteredValueForRegister
                        $script:ManuallyEnteredValueForRegister = ""
                    }
                } else {
                    Enter-DataToRegisterManually -Title 'Документ не упоминается ни в одной из спецификаций' -Label "Укажите наименование для документа $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 6) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                }
                #ФОРМАТ
                if ($script:CollectedReferences[2][$script:CollectedReferences[0].IndexOf("$($_.Text)")] -ne "") {
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 7) = $script:CollectedReferences[2][$script:CollectedReferences[0].IndexOf("$($_.Text)")]
                } else {
                    Enter-DataToRegisterManually -Title 'Формат для документа не указан ни в одной из спецификаций' -Label "Укажите формат для документа $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 7) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                }
                #КОЛИЧЕСТВО ЛИСТОВ
                #Гид по стилю
                if ($($_.Text) -match "DSG") {
                    Enter-DataToRegisterManually -Title 'Документ с кодом DSG' -Label "Укажите общее количество страниц для документа $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                #Excel-файл
                } elseif ((Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).xlsx") -or (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).xls")) {
                    Enter-DataToRegisterManually -Title 'Документ созданный в приложении MS Excel' -Label "Укажите общее количество страниц для документа $($_.Text):"
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = $script:ManuallyEnteredValueForRegister
                    $script:ManuallyEnteredValueForRegister = ""
                #Остальные файлы, т.е. WORD-файлы
                } else {
                    if (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).docx") {$DocumentRegister = $WordApplication.Documents.Open("$script:SelectedFolderWithFilesBeingPublished\$($_.Text).docx")}
                    if (Test-Path -Path "$script:SelectedFolderWithFilesBeingPublished\$($_.Text).doc") {$DocumentRegister = $WordApplication.Documents.Open("$script:SelectedFolderWithFilesBeingPublished\$($_.Text).doc")}
                    $Wholestory = $DocumentRegister.Range()
                    $TotalPages = $Wholestory.Information(4)
                    $RegisterWorksheet.Cells.Item($ActionFlag + 1, 8) = $TotalPages
                    $DocumentRegister.Close()
                }
                #ДАТА ПОСТУПЛЕНИЯ
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 9) = $CalendarIssueDateInput.Text
                #ИЗМ.
                if ($_.SubItems[1].Text -eq '-') {$RegisterWorksheet.Cells.Item($ActionFlag + 1, 10) = (0 + 1)} else {$RegisterWorksheet.Cells.Item($ActionFlag + 1, 10) = ([int]($_.SubItems[1].Text) + 1)}
                #№ ДОКУМ. ВВЕДЕН
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(11).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag + 1, 11) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(11).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag + 1).Cells.Item(1).Interior.Color = 139
            }
        }   
    }
    #Список АННУЛИРОВАТЬ
    Write-Host "======================================="
    Write-Host "Работую со списком Аннулировать..."
    $ListViewRemove.Items | % {
        $ArrayContainsExtensionFlag = $false
        $ArrayContainsExtensionIndex = $null
        $RegisterLastRow += 1
        if ($_.SubItems[2].Text -eq "Программа") {
            Write-Host "$($_.Text)"
            $ActionFlag = $null
            #Определяем есть ли строка заменяемого файла, подходящая для заполнения
            $ActionFlag = Find-StringToBePopulated -Sheet $RegisterWorksheet -LookFor $_.Text -ColumnLetter "C"
            #Write-Host "Значение ячейки $ActionFlag"
            #Если подходящая строка найдена, то заполняем ее и создаем новую строку для заменяющего файла под строкой для заменяемого файла
            if ($ActionFlag -ne 'not exist' -and $ActionFlag -ne 'not found' -and $ActionFlag -ne 'multiple instances') {
                #Заполняем строку заменяемого файла
                $RegisterWorksheet.Cells.Item($ActionFlag, 12) = $CalendarIssueDateInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 13) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).HorizontalAlignment = -4131
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 14) = "аннулирован"
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(1).Interior.Color = 139
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(2).Interior.Color = 16436871
            }
        }
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "$($_.Text)"
            $ActionFlag = $null
            #Определяем есть ли строка заменяемого файла, подходящая для заполнения
            $ActionFlag = Find-StringToBePopulated -Sheet $RegisterWorksheet -LookFor $_.Text -ColumnLetter "E"
            #Write-Host "Значение ячейки $ActionFlag"
            #Если подходящая строка найдена, то заполняем ее и создаем новую строку для заменяющего файла под строкой для заменяемого файла
            if ($ActionFlag -ne 'not exist' -and $ActionFlag -ne 'not found' -and $ActionFlag -ne 'multiple instances') {
                #Заполняем строку заменяемого файла
                $RegisterWorksheet.Cells.Item($ActionFlag, 12) = $CalendarIssueDateInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 13) = $UpdateNotificationNumberInput.Text
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(13).HorizontalAlignment = -4131
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).NumberFormat = "@"
                $RegisterWorksheet.Cells.Item($ActionFlag, 14) = "аннулирован"
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(14).HorizontalAlignment = -4131
                #Залить первую ячейку в строке
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(1).Interior.Color = 139
                $RegisterWorksheet.Rows.Item($ActionFlag).Cells.Item(2).Interior.Color = 16436871
            }
        }
    }
    Write-Host "АВТОЗАПОЛНЕНИЕ ЗАКОНЧЕНО. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
    Write-Host "АВТОЗАПОЛНЕНИЕ ЗАКОНЧЕНО. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
    Write-Host "АВТОЗАПОЛНЕНИЕ ЗАКОНЧЕНО. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
    $WordApplication.Quit()
}

Function Save-File
{ 
    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "XML file (*.xml)| *.xml"
    $DialogResult = $SaveFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SaveFileDialog.FileName} else {return $null}
}

Function Import-EntryToList ($ItemName, $ItemAttribute, $ItemType, $ItemBackColor, $ListToAdd)
{
    $ItemsOnTheList = @()
    $ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}
    if ($ItemsOnTheList -contains "$ItemName") {
        if ((Show-MessageBox -Message "Файл с обозначением $($ItemName) уже содержится в списках. Перезаписать?`r`n`r`nНажмите Да, чтобы удалить существующую запись и добавить новую.`r`nНажмите Нет, чтобы продолжить без внесения изменений." -Title "Выберите действие" -Type YesNo) -eq "Yes") {
            #Ответ да. Скрипт меняет удаляет существующую запись и добавляет новую в указанный список
            $ListViewAdd.Items | % {if ($_.Text -eq "$ItemName") {$_.Remove()}}
            $ListViewReplace.Items | % {if ($_.Text -eq "$ItemName") {$_.Remove()}}
            $ListViewRemove.Items | % {if ($_.Text -eq "$ItemName") {$_.Remove()}}
            $ItemToImport = New-Object System.Windows.Forms.ListViewItem("$ItemName")
            $ItemToImport.SubItems.Add("$ItemAttribute")
            $ItemToImport.SubItems.Add("$ItemType")
            $ItemToImport.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
            $ArgbSettings = $ItemBackColor -split ","
            $ItemToImport.BackColor = [System.Drawing.Color]::FromArgb([int]$ArgbSettings[0],[int]$ArgbSettings[1],[int]$ArgbSettings[2],[int]$ArgbSettings[3])
            $ListToAdd.Items.Insert($ListToAdd.Items.Count, $ItemToImport)
        }
    } else {
        $ItemToImport = New-Object System.Windows.Forms.ListViewItem("$ItemName")
        $ItemToImport.SubItems.Add("$ItemAttribute")
        $ItemToImport.SubItems.Add("$ItemType")
        $ItemToImport.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
        $ArgbSettings = $ItemBackColor -split ","
        $ItemToImport.BackColor = [System.Drawing.Color]::FromArgb([int]$ArgbSettings[0],[int]$ArgbSettings[1],[int]$ArgbSettings[2],[int]$ArgbSettings[3])
        $ListToAdd.Items.Insert($ListToAdd.Items.Count, $ItemToImport)
    }
}

Function Import-FromXml ($SpecifiedFile)
{
    $InputXmlFile = New-Object System.Xml.XmlDocument
    $InputXmlFile.Load("$SpecifiedFile")
    $InputXmlFile.SelectSingleNode("/script-data/lists/publish-list/item").Attributes.GetNamedItem("").Value
    $ImportPublishList = $InputXmlFile.SelectNodes("/script-data/lists/publish-list/item")
    ForEach ($Item in $ImportPublishList) {Import-EntryToList -ListToAdd $ListViewAdd -ItemName $Item.InnerText -ItemAttribute $Item.Attributes.GetNamedItem("version-checksum").Value -ItemType $Item.Attributes.GetNamedItem("type").Value -ItemBackColor $Item.Attributes.GetNamedItem("color").Value}
    $ImportReplaceList = $InputXmlFile.SelectNodes("/script-data/lists/replace-list/item")
    ForEach ($Item in $ImportReplaceList) {Import-EntryToList -ListToAdd $ListViewReplace -ItemName $Item.InnerText -ItemAttribute $Item.Attributes.GetNamedItem("version-checksum").Value -ItemType $Item.Attributes.GetNamedItem("type").Value -ItemBackColor $Item.Attributes.GetNamedItem("color").Value}
    $ImportRemoveList = $InputXmlFile.SelectNodes("/script-data/lists/remove-list/item")
    ForEach ($Item in $ImportRemoveList) {Import-EntryToList -ListToAdd $ListViewRemove -ItemName $Item.InnerText -ItemAttribute $Item.Attributes.GetNamedItem("version-checksum").Value -ItemType $Item.Attributes.GetNamedItem("type").Value -ItemBackColor $Item.Attributes.GetNamedItem("color").Value}
    $UpdateNotificationNumberInput.Text = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/notification-number").InnerText
    $CalendarIssueDateInput.Text = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/issue-date").InnerText
    $CalendarApplyUpdatesUntilInput.Text = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/apply-until").InnerText
    $script:GlobalReasonField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/reason").InnerText
    $script:GlobalInStoreField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/in-store").InnerText
    $script:GlobalStartUsageField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/start-usage").InnerText
    $script:GlobalApplicableToField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/applicable-to").InnerText
    $script:GlobalSendToField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/send-to").InnerText
    $script:GlobalAppendixField = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/appendix").InnerText
    if ($ComboboxDepartmentName.Items.Contains("$($InputXmlFile.SelectSingleNode("/script-data/notification-settings/department-name").InnerText)") -eq $true) {
        $ComboboxDepartmentName.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/department-name").InnerText
    } else {
        $ComboboxDepartmentName.Items.Add($($InputXmlFile.SelectSingleNode("/script-data/notification-settings/department-name").InnerText))
        $ComboboxDepartmentName.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/department-name").InnerText
    }
    if ($ComboboxCreatedBy.Items.Contains("$($InputXmlFile.SelectSingleNode("/script-data/notification-settings/created-by").InnerText)") -eq $true) {
        $ComboboxCreatedBy.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/created-by").InnerText
    } else {
        $ComboboxCreatedBy.Items.Add($($InputXmlFile.SelectSingleNode("/script-data/notification-settings/created-by").InnerText))
        $ComboboxCreatedBy.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/created-by").InnerText
    }
    if ($ComboboxCheckedBy.Items.Contains("$($InputXmlFile.SelectSingleNode("/script-data/notification-settings/checked-by").InnerText)") -eq $true) {
        $ComboboxCheckedBy.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/checked-by").InnerText
    } else {
        $ComboboxCheckedBy.Items.Add($($InputXmlFile.SelectSingleNode("/script-data/notification-settings/checked-by").InnerText))
        $ComboboxCheckedBy.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/checked-by").InnerText
    }
    $ComboboxCodes.SelectedItem = $InputXmlFile.SelectSingleNode("/script-data/notification-settings/reason-code").InnerText
    Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
}

Function Export-ToXmlFile ($SpecifiedFile)
{

    $OutputXmlFile = New-Object System.Xml.XmlDocument
    $OutputXmlFile.CreateXmlDeclaration("1.0","UTF-8",$null)
    $OutputXmlFile.AppendChild($OutputXmlFile.CreateXmlDeclaration("1.0","UTF-8",$null))
$InfoForXml = @"
Файл сгенерирован: $(Get-Date)
"@
    $OutputXmlFile.AppendChild($OutputXmlFile.CreateComment($InfoForXml))
    $RootElement = $OutputXmlFile.CreateNode("element","script-data",$null)
    $OutputXmlFile.AppendChild($RootElement)
    $XmlExportLists = $OutputXmlFile.CreateNode("element","lists",$null)
    $RootElement.AppendChild($XmlExportLists)
    $NewList = $OutputXmlFile.CreateNode("element","publish-list",$null)
    $XmlExportLists.AppendChild($NewList)
    $NewList = $OutputXmlFile.CreateNode("element","replace-list",$null)
    $XmlExportLists.AppendChild($NewList)
    $NewList = $OutputXmlFile.CreateNode("element","remove-list",$null)
    $XmlExportLists.AppendChild($NewList)
    Foreach ($ListItem in $ListViewAdd.Items)  {
        $ItemBackColor = ""
        $Element = $OutputXmlFile.CreateNode("element","item",$null)
        $Element.InnerText = $ListItem.Text
        $ElementAttribute = $OutputXmlFile.CreateAttribute("version-checksum")
        $ElementAttribute.Value = $ListItem.SubItems[1].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("type")
        $ElementAttribute.Value = $ListItem.SubItems[2].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("color")
        $ItemBackColor = "$($ListItem.BackColor.A)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.R)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.G)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.B)"
        $ElementAttribute.Value = $ItemBackColor
        $Element.Attributes.Append($ElementAttribute)
        $OutputXmlFile.SelectSingleNode("/script-data/lists/publish-list").AppendChild($Element)
    }
    Foreach ($ListItem in $ListViewReplace.Items)  {
        $ItemBackColor = ""
        $Element = $OutputXmlFile.CreateNode("element","item",$null)
        $Element.InnerText = $ListItem.Text
        $ElementAttribute = $OutputXmlFile.CreateAttribute("version-checksum")
        $ElementAttribute.Value = $ListItem.SubItems[1].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("type")
        $ElementAttribute.Value = $ListItem.SubItems[2].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("color")
        $ItemBackColor = "$($ListItem.BackColor.A)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.R)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.G)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.B)"
        $ElementAttribute.Value = $ItemBackColor
        $Element.Attributes.Append($ElementAttribute)
        $OutputXmlFile.SelectSingleNode("/script-data/lists/replace-list").AppendChild($Element)
    }
    Foreach ($ListItem in $ListViewRemove.Items)  {
        $ItemBackColor = ""
        $Element = $OutputXmlFile.CreateNode("element","item",$null)
        $Element.InnerText = $ListItem.Text
        $ElementAttribute = $OutputXmlFile.CreateAttribute("version-checksum")
        $ElementAttribute.Value = $ListItem.SubItems[1].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("type")
        $ElementAttribute.Value = $ListItem.SubItems[2].Text
        $Element.Attributes.Append($ElementAttribute)
        $ElementAttribute = $OutputXmlFile.CreateAttribute("color")
        $ItemBackColor = "$($ListItem.BackColor.A)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.R)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.G)"
        $ItemBackColor = $ItemBackColor + ",$($ListItem.BackColor.B)"
        $ElementAttribute.Value = $ItemBackColor
        $Element.Attributes.Append($ElementAttribute)
        $OutputXmlFile.SelectSingleNode("/script-data/lists/remove-list").AppendChild($Element)
    }
    $NotificationSettings = $OutputXmlFile.CreateNode("element","notification-settings",$null)
    $RootElement.AppendChild($NotificationSettings)
    $NewSetting = $OutputXmlFile.CreateNode("element","notification-number",$null)
    $NewSetting.InnerText = $UpdateNotificationNumberInput.Text
    $NotificationSettings.AppendChild($NewSetting) 
    $NewSetting = $OutputXmlFile.CreateNode("element","issue-date",$null)
    $NewSetting.InnerText = $CalendarIssueDateInput.Text
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","apply-until",$null)
    $NewSetting.InnerText = $CalendarApplyUpdatesUntilInput.Text
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","reason",$null)
    $NewSetting.InnerText = $script:GlobalReasonField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","in-store",$null)
    $NewSetting.InnerText = $script:GlobalInStoreField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","start-usage",$null)
    $NewSetting.InnerText = $script:GlobalStartUsageField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","applicable-to",$null)
    $NewSetting.InnerText = $script:GlobalApplicableToField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","send-to",$null)
    $NewSetting.InnerText = $script:GlobalSendToField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","appendix",$null)
    $NewSetting.InnerText = $script:GlobalAppendixField
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","department-name",$null)
    $NewSetting.InnerText = $ComboboxDepartmentName.SelectedItem
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","created-by",$null)
    $NewSetting.InnerText = $ComboboxCreatedBy.SelectedItem
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","checked-by",$null)
    $NewSetting.InnerText = $ComboboxCheckedBy.SelectedItem
    $NotificationSettings.AppendChild($NewSetting)
    $NewSetting = $OutputXmlFile.CreateNode("element","reason-code",$null)
    $NewSetting.InnerText = $ComboboxCodes.SelectedItem
    $NotificationSettings.AppendChild($NewSetting)
    $OutputXmlFile.Save("$SpecifiedFile")
}

Function BulkImportAdd-ItemToList ($FileName, $VersionNumber, $FileType, $HighlightFlag, $TestPathFullName)
{
    $ItemsOnTheList = @()
    $ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}
    if ($ItemsOnTheList -contains "$($FileName)") {
        if ((Show-MessageBox -Message "Файл с обозначением $($FileName) уже содержится в списках. Перезаписать?`r`n`r`nНажмите Да, чтобы удалить существующую запись и добавить новую.`r`nНажмите Нет, чтобы продолжить без внесения изменений." -Title "Выберите действие" -Type YesNo) -eq "Yes") {
            #Ответ да. Скрипт меняет удаляет существующую запись и добавляет новую в указанный список
            $ListViewAdd.Items | % {if ($_.Text -eq "$($FileName)") {$_.Remove()}}
            $ListViewReplace.Items | % {if ($_.Text -eq "$($FileName)") {$_.Remove()}}
            $ListViewRemove.Items | % {if ($_.Text -eq "$($FileName)") {$_.Remove()}}
            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($FileName)")
            $ItemToAdd.SubItems.Add("$($VersionNumber)")
            $ItemToAdd.SubItems.Add("$($FileType)")
            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
            if ($HighlightFlag -eq 1) {
                if ($script:HighlightChecboxStatus -eq $true) {$ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 15, 177, 255)}
            }
            if ($script:SuspiciousAction -eq $true) {
                Show-MessageBox -Message "Документ: $FileName`r`n`r`nДействие: Вы указали знак минус (-) в качестве значения для Изм. (т.е. документ новый и никогда не публиковался), но в текущей версии проекта уже существует документ с таким обозначением.`r`n`r`nДанная запись будет выделена красным цветом и должна быть проверена после окончания пакетного импорта." -Title "Подозрительное действие" -Type OK
                $ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 1)
                Write-Host "Ошибка в Изм." -ForegroundColor Red
            }
            if ($script:IncorrectVersionDiscrepancy -eq $true -and $script:SuspiciousAction -eq $false) {                
                Show-MessageBox -Message "Документ: $FileName`r`n`r`nОшибка: Номер Изм., присвоенный импортируемому документу, не является валидным.`r`n`r`nВозможные причины:`r`n*Значение Изм. в заменяемом документе больше или равно значению Изм. в импортируемом документе.`r`n*Разница между значениями Изм., указанными в импортируемом и заменяемом документах, не является 1.`r`n`r`nДанная запись будет выделена красным цветом и должна быть проверена после окончания пакетного импорта." -Title "Ошибка" -Type OK
                $ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 1)
                Write-Host "Ошибка в Изм." -ForegroundColor Red
            }
            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                if (Test-Path -Path $TestPathFullName) {
                #Если файл существует в папке с текущей версией проекта, то заменяем
                    $ListViewReplace.Items.Insert(0, $ItemToAdd)
                } else {
                #Если файл не существует в папке с текущей версией проекта, то выпускаем
                    $ListViewAdd.Items.Insert(0, $ItemToAdd)
                }
            } else {
                if ($BulkImportFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                if ($BulkImportFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                if ($BulkImportFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
            }
            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
            Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove
        }
    } else {
            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($FileName)")
            $ItemToAdd.SubItems.Add("$($VersionNumber)")
            $ItemToAdd.SubItems.Add("$($FileType)")
            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
            if ($HighlightFlag -eq 1) {
                if ($script:HighlightChecboxStatus -eq $true) {$ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 15, 177, 255)}
            }
            if ($script:SuspiciousAction -eq $true) {
                Show-MessageBox -Message "Документ: $FileName`r`n`r`nДействие: Вы указали знак минус (-) в качестве значения для Изм. (т.е. документ новый и никогда не публиковался), но в текущей версии проекта уже существует документ с таким обозначением.`r`n`r`nДанная запись будет выделена красным цветом и должна быть проверена после окончания пакетного импорта." -Title "Подозрительное действие" -Type OK
                $ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 1)
                Write-Host "Ошибка в Изм." -ForegroundColor Red
            }
            if ($script:IncorrectVersionDiscrepancy -eq $true -and $script:SuspiciousAction -eq $false) {                
                Show-MessageBox -Message "Документ: $FileName`r`n`r`nОшибка: Номер Изм., присвоенный импортируемому документу, не является валидным.`r`n`r`nВозможные причины:`r`n*Значение Изм. в заменяемом документе больше или равно значению Изм. в импортируемом документе.`r`n*Разница между значениями Изм., указанными в импортируемом и заменяемом документах, не является 1.`r`n`r`nДанная запись будет выделена красным цветом и должна быть проверена после окончания пакетного импорта." -Title "Ошибка" -Type OK
                $ItemToAdd.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 1)
                Write-Host "Ошибка в Изм." -ForegroundColor Red
            }
            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                if (Test-Path -Path $TestPathFullName) {
                #Если файл существует в папке с текущей версией проекта, то заменяем
                    $ListViewReplace.Items.Insert(0, $ItemToAdd)
                } else {
                #Если файл не существует в папке с текущей версией проекта, то выпускаем
                    $ListViewAdd.Items.Insert(0, $ItemToAdd)
                }
            } else {
                if ($BulkImportFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                if ($BulkImportFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                if ($BulkImportFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
            }
            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
            Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove
    }
}

Function BulkImport-InputFileDataForm ($FileName, $FileType, $FormTitle)
{
    Write-Host "Пожалуйста, укажите необходимые данные для документа $FileName"
    $InputFileDataForm = New-Object System.Windows.Forms.Form
    $InputFileDataForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $InputFileDataForm.ShowIcon = $false
    $InputFileDataForm.AutoSize = $true
    $InputFileDataForm.Text = $FormTitle
    $InputFileDataForm.AutoSizeMode = "GrowAndShrink"
    $InputFileDataForm.WindowState = "Normal"
    $InputFileDataForm.SizeGripStyle = "Hide"
    $InputFileDataForm.ShowInTaskbar = $true
    $InputFileDataForm.StartPosition = "CenterScreen"
    $InputFileDataForm.MinimizeBox = $false
    $InputFileDataForm.MaximizeBox = $false
    $InputFileDataForm.ControlBox = $false
    #Надпись к поля для ввода обозначение
    $InputFileDataFormFileNameLabel = New-Object System.Windows.Forms.Label
    $InputFileDataFormFileNameLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $InputFileDataFormFileNameLabel.Width = 81
    $InputFileDataFormFileNameLabel.Text = "Обозначение:"
    $InputFileDataFormFileNameLabel.TextAlign = "TopLeft"
    $InputFileDataForm.Controls.Add($InputFileDataFormFileNameLabel)
    
    #Поле для ввода обозначение
    $InputFileDataFormFileNameInput = New-Object System.Windows.Forms.TextBox 
    $InputFileDataFormFileNameInput.Location = New-Object System.Drawing.Point(95,13) #x,y
    $InputFileDataFormFileNameInput.Width = 270
    $InputFileDataFormFileNameInput.Text = "$($FileName)"
    $InputFileDataFormFileNameInput.ForeColor = "Gray"
    $InputFileDataFormFileNameInput.Enabled = $false
    $InputFileDataForm.Controls.Add($InputFileDataFormFileNameInput)
    
    #Надпись к списку для указания типа файла
    $InputFileDataFormFileTypeLabel = New-Object System.Windows.Forms.Label
    $InputFileDataFormFileTypeLabel.Location = New-Object System.Drawing.Point(10,45) #x,y
    $InputFileDataFormFileTypeLabel.Width = 81
    $InputFileDataFormFileTypeLabel.Text = "Тип файла:"
    $InputFileDataFormFileTypeLabel.TextAlign = "TopLeft"
    $InputFileDataForm.Controls.Add($InputFileDataFormFileTypeLabel)
    
    #Список содержащий доступные типы файлов
    $DataTypes = @("Документ","Программа")
    $InputFileDataFormFileTypeCombobox = New-Object System.Windows.Forms.ComboBox
    $InputFileDataFormFileTypeCombobox.Location = New-Object System.Drawing.Point(95,43) #x,y
    $InputFileDataFormFileTypeCombobox.DropDownStyle = "DropDownList"
    $InputFileDataFormFileTypeCombobox.Enabled = $false
    $DataTypes | % {$InputFileDataFormFileTypeCombobox.Items.add($_)}
    if ($FileType -eq "Документ") {$InputFileDataFormFileTypeCombobox.SelectedIndex = 0} else {$InputFileDataFormFileTypeCombobox.SelectedIndex = 1}
    $InputFileDataForm.Controls.Add($InputFileDataFormFileTypeCombobox)
    
    #Надпись к полю для ввода MD5 и Изм.
    $InputFileDataFormAttributeValueLabel = New-Object System.Windows.Forms.Label
    $InputFileDataFormAttributeValueLabel.Location =  New-Object System.Drawing.Point(10,75) #x,y
    $InputFileDataFormAttributeValueLabel.Width = 81
    $InputFileDataFormAttributeValueLabel.Text = "Изм./MD5:"
    $InputFileDataFormAttributeValueLabel.TextAlign = "TopLeft"
    $InputFileDataForm.Controls.Add($InputFileDataFormAttributeValueLabel)
    
    #Поле для ввода MD5 и Изм.
    $InputFileDataFormAttributeValueInput = New-Object System.Windows.Forms.TextBox 
    $InputFileDataFormAttributeValueInput.Location = New-Object System.Drawing.Point(95,73) #x,y
    $InputFileDataFormAttributeValueInput.Width = 270
    $InputFileDataFormAttributeValueInput.Text = "-"
    $InputFileDataFormAttributeValueInput.Add_GotFocus({
        if ($InputFileDataFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            $InputFileDataFormAttributeValueInput.Text = ""
            $InputFileDataFormAttributeValueInput.ForeColor = "Black"
        }
        })
    $InputFileDataFormAttributeValueInput.Add_LostFocus({
        if ($InputFileDataFormAttributeValueInput.Text -eq "") {
            $InputFileDataFormAttributeValueInput.Text = "Укажите Изм. или MD5..."
            $InputFileDataFormAttributeValueInput.ForeColor = "Gray"
        }
        })
    $InputFileDataForm.Controls.Add($InputFileDataFormAttributeValueInput)
    #Выделить цветом
    $InputFileDataFormHighlightCheckbox = New-Object System.Windows.Forms.CheckBox
    $InputFileDataFormHighlightCheckbox.Width = 410
    $InputFileDataFormHighlightCheckbox.Text = "Выделить запись цветом"
    $InputFileDataFormHighlightCheckbox.Location = New-Object System.Drawing.Point(10,100) #x,y
    $InputFileDataFormHighlightCheckbox.Enabled = $true
    $InputFileDataFormHighlightCheckbox.Checked = $true
    $InputFileDataForm.Controls.Add($InputFileDataFormHighlightCheckbox)
    #Кнопка добавить
    $InputFileDataFormAddButton = New-Object System.Windows.Forms.Button
    $InputFileDataFormAddButton.Location = New-Object System.Drawing.Point(10,130) #x,y
    $InputFileDataFormAddButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $InputFileDataFormAddButton.Text = "Добавить"
    $InputFileDataFormAddButton.Add_Click({
    if ($InputFileDataFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
        Show-MessageBox -Message "Пожалуйста, укажите Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
    } else {
        if ($InputFileDataFormHighlightCheckbox.Checked -eq $true) {$script:HighlightChecboxStatus = $true} else {$script:HighlightChecboxStatus = $false}
        $script:VerNumber = "$($InputFileDataFormAttributeValueInput.Text)"
        $InputFileDataForm.Close()
    }
    })
    $InputFileDataForm.Controls.Add($InputFileDataFormAddButton)
    $InputFileDataForm.ActiveControl = $InputFileDataFormFileTypeLabel
    $InputFileDataForm.ShowDialog()
}

Function BulkImport ($ListOfSelectedFiles)
{
    Write-Host "Пакетный импорт начат"
    #Создать экземпляр приложения MS Word
    $WordReadData = New-Object -ComObject Word.Application
    #Создать экземпляр приложения MS Word
    $CurrentVersionWordReadData = New-Object -ComObject Word.Application
    #Сделать вызванное приложение невидемым
    $WordReadData.Visible = $false
    $CurrentVersionWordReadData.Visible = $false
    $ItemsOnTheList = @()
    $ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}
    $ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}
    $ListOfSelectedFiles | % {
        [string]$script:VerNumber = ""
        $script:SuspiciousAction = $false
        $script:IncorrectVersionDiscrepancy = $false
        $script:CurrentVersionDocumentExists = $false
        $FileNameWOExtension = [System.IO.Path]::GetFileNameWithoutExtension($_)
        $FileNameExtension = [System.IO.Path]::GetExtension($_)
        $CurrentFileNameExtension = [System.IO.Path]::GetExtension($_)
        Write-Host "Работаю с: " $FileNameWOExtension
        #Проверка на тип файла (документ)
        if ($_ -match '([A-Z0-9]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d') {
            #Проверка на дизайн гайд
            if ($_ -match '\d\d\.([a-z]{1})(DSG)\.\d\d') {
                BulkImport-InputFileDataForm -FileName $FileNameWOExtension -FileType "Документ" -FormTitle "Введите номер Изм., присвоенный DSG документу"
                if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                    if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                       try {$script:VerNumber = $script:VerNumber - 1} catch {$script:SuspiciousAction = $true}
                    }
                }
                if ($script:VerNumber -eq 0) {$script:VerNumber = "-"}
                BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $script:VerNumber -FileType "Документ" -HighlightFlag 1 -TestPathFullName "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))"
            #Проверка на эксель   
            } elseif ($([System.IO.Path]::GetExtension($_)) -eq '.xlsx' -or $([System.IO.Path]::GetExtension($_)) -eq '.xls') {
                BulkImport-InputFileDataForm -FileName $FileNameWOExtension -FileType "Документ" -FormTitle "Введите номер Изм., присвоенный Excel документу"
                if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                    if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                       try {$script:VerNumber = $script:VerNumber - 1} catch {$script:SuspiciousAction = $true}
                    }
                }
                if ($script:VerNumber -eq 0) {$script:VerNumber = "-"}
                BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $script:VerNumber -FileType "Документ" -HighlightFlag 1 -TestPathFullName "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))"
            #Проверки пройдены -- можно считать данные у файла
            } else {
                #Отбрасываем PDF файлы
                if ($([System.IO.Path]::GetExtension($_)) -ne '.pdf') {
                    $DocumentReadData = $WordReadData.Documents.Open($_)
                    #Код для спецификаци и списка проектной документации
                    if ($_ -match '\d\d\.([a-z]{1})(SPC)\.\d\d' -or $_ -match '\d\d\.([a-z]{1})(LPD)\.\d\d') {
                        [string]$ValueOfVersionNumber = try {($DocumentReadData.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                        if ($ValueOfVersionNumber -eq "error") {
                            BulkImport-InputFileDataForm -FileName $FileNameWOExtension -FileType "Документ" -FormTitle "Введите номер Изм., присвоенный публикуемому Word документу"
                            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                                #Проверяет существует ли в папке в текущем проекте файл стаким же обозначением и расширением.
                                if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                                    $script:CurrentVersionDocumentExists = $true
                                #Если не существует, то заменяем расширение, на соответсвующее.
                                } else {
                                    #Если проверялся docx файл, то проверяем doc
                                    if ($FileNameExtension -eq '.docx') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).doc") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.doc'}
                                    }
                                    #Если проверялся doc файл, то проверяем docx
                                    if ($FileNameExtension -eq '.doc') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).docx") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.docx'}
                                    }
                                }
                                if ($script:CurrentVersionDocumentExists -eq $true) {
                                   $CurrentVersionDocumentReadData = $CurrentVersionWordReadData.Documents.Open("$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)")
                                   [string]$ValueOfCurrentVersionNumber = try {($CurrentVersionDocumentReadData.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                                   if ($ValueOfCurrentVersionNumber -ne 'error') {
                                       if ($ValueOfCurrentVersionNumber -eq "") {[int]$ValueOfCurrentVersionNumber = 0} else {[int]$ValueOfCurrentVersionNumber = $ValueOfCurrentVersionNumber}
                                       if ($script:VerNumber -ne "-") {[int]$script:VerNumber = $script:VerNumber}
                                       if ($script:VerNumber -eq "-") {[int]$script:VerNumber = 0}
                                       #Write-Host "Current version: $ValueOfCurrentVersionNumber"
                                       #Write-Host "Imported version: $script:VerNumber"
                                       if ($ValueOfCurrentVersionNumber -ge $script:VerNumber -or ($script:VerNumber - $ValueOfCurrentVersionNumber) -ne 1) {$script:IncorrectVersionDiscrepancy = $true}
                                       $CurrentVersionDocumentReadData.Close([ref]0)
                                   }
                                   if ($script:VerNumber -eq 0) {[string]$script:VerNumber = "-"}
                                   try {$script:VerNumber = $script:VerNumber - 1} catch {$script:SuspiciousAction = $true}
                                }
                            }
                            if ($script:VerNumber -eq 0) {[string]$script:VerNumber = "-"}
                            BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $script:VerNumber -FileType "Документ" -HighlightFlag 1 -TestPathFullName "$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)" 
                        } else {
                            if ($ValueOfVersionNumber -eq "") {[int]$ValueOfVersionNumber = 0} else {[int]$ValueOfVersionNumber}
                            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                                #Проверяет существует ли в папке в текущем проекте файл стаким же обозначением и расширением.
                                if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                                    $script:CurrentVersionDocumentExists = $true
                                #Если не существует, то заменяем расширение, на соответсвующее.
                                } else {
                                    #Если проверялся docx файл, то проверяем doc
                                    if ($FileNameExtension -eq '.docx') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).doc") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.doc'}
                                    }
                                    #Если проверялся doc файл, то проверяем docx
                                    if ($FileNameExtension -eq '.doc') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).docx") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.docx'}
                                    }
                                }
                                if ($script:CurrentVersionDocumentExists -eq $true) {
                                   $CurrentVersionDocumentReadData = $CurrentVersionWordReadData.Documents.Open("$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)")
                                   [string]$ValueOfCurrentVersionNumber = try {($CurrentVersionDocumentReadData.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                                   if ($ValueOfCurrentVersionNumber -ne 'error') {
                                       if ($ValueOfCurrentVersionNumber -eq "") {[int]$ValueOfCurrentVersionNumber = 0} else {[int]$ValueOfCurrentVersionNumber = $ValueOfCurrentVersionNumber}
                                       #Write-Host "Current version: $ValueOfCurrentVersionNumber"
                                       #Write-Host "Imported version: $ValueOfVersionNumber"
                                       if ($ValueOfCurrentVersionNumber -ge $ValueOfVersionNumber -or ($ValueOfVersionNumber - $ValueOfCurrentVersionNumber) -ne 1) {$script:IncorrectVersionDiscrepancy = $true}
                                       $CurrentVersionDocumentReadData.Close([ref]0)
                                   }
                                   $ValueOfVersionNumber = $ValueOfVersionNumber - 1
                                }
                            }
                            if ($ValueOfVersionNumber -eq 0) {[string]$ValueOfVersionNumber = "-"}
                            BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $ValueOfVersionNumber -FileType "Документ" -HighlightFlag 0 -TestPathFullName "$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)"
                        }
                #Код для шаблонов в Word документах
                } else {
                        [string]$ValueOfVersionNumber = try {($DocumentReadData.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                        if ($ValueOfVersionNumber -eq "error") {
                            BulkImport-InputFileDataForm -FileName $FileNameWOExtension -FileType "Документ" -FormTitle "Введите номер Изм., присвоенный публикуемому Word документу"
                            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                                #Проверяет существует ли в папке в текущем проекте файл стаким же обозначением и расширением.
                                if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                                    $script:CurrentVersionDocumentExists = $true
                                #Если не существует, то заменяем расширение, на соответсвующее.
                                } else {
                                    #Если проверялся docx файл, то проверяем doc
                                    if ($FileNameExtension -eq '.docx') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).doc") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.doc'}
                                    }
                                    #Если проверялся doc файл, то проверяем docx
                                    if ($FileNameExtension -eq '.doc') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).docx") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.docx'}
                                    }
                                }
                                if ($script:CurrentVersionDocumentExists -eq $true) {
                                   $CurrentVersionDocumentReadData = $CurrentVersionWordReadData.Documents.Open("$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)")
                                   [string]$ValueOfCurrentVersionNumber = try {($CurrentVersionDocumentReadData.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                                   if ($ValueOfCurrentVersionNumber -ne 'error') {
                                       if ($ValueOfCurrentVersionNumber -eq "") {[int]$ValueOfCurrentVersionNumber = 0} else {[int]$ValueOfCurrentVersionNumber = $ValueOfCurrentVersionNumber}
                                       if ($script:VerNumber -ne "-") {[int]$script:VerNumber = $script:VerNumber}
                                       if ($script:VerNumber -eq "-") {[int]$script:VerNumber = 0}
                                       #Write-Host "Current version: $ValueOfCurrentVersionNumber"
                                       #Write-Host "Imported version: $script:VerNumber"
                                       if ($ValueOfCurrentVersionNumber -ge $script:VerNumber -or ($script:VerNumber - $ValueOfCurrentVersionNumber) -ne 1) {$script:IncorrectVersionDiscrepancy = $true}
                                       $CurrentVersionDocumentReadData.Close([ref]0)
                                   }
                                   if ($script:VerNumber -eq 0) {[string]$script:VerNumber = "-"}
                                   try {$script:VerNumber = $script:VerNumber - 1} catch {$script:SuspiciousAction = $true}
                                }
                            }
                            if ($script:VerNumber -eq 0) {[string]$script:VerNumber = "-"}
                            BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $script:VerNumber -FileType "Документ" -HighlightFlag 1 -TestPathFullName "$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)"
                        } else {
                            if ($ValueOfVersionNumber -eq "") {[int]$ValueOfVersionNumber = 0} else {[int]$ValueOfVersionNumber}
                            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                                #Проверяет существует ли в папке в текущем проекте файл стаким же обозначением и расширением.
                                if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                                    $script:CurrentVersionDocumentExists = $true
                                #Если не существует, то заменяем расширение, на соответсвующее.
                                } else {
                                    #Если проверялся docx файл, то проверяем doc
                                    if ($FileNameExtension -eq '.docx') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).doc") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.doc'}
                                    }
                                    #Если проверялся doc файл, то проверяем docx
                                    if ($FileNameExtension -eq '.doc') {
                                        if (Test-Path -Path "$($script:PathToCurrentVrsion)\$($FileNameWOExtension).docx") {$script:CurrentVersionDocumentExists = $true; $CurrentFileNameExtension = '.docx'}
                                    }
                                }
                                if ($script:CurrentVersionDocumentExists -eq $true) {
                                   $CurrentVersionDocumentReadData = $CurrentVersionWordReadData.Documents.Open("$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)")
                                   [string]$ValueOfCurrentVersionNumber = try {($CurrentVersionDocumentReadData.Tables.Item(1).Cell(7, 3).Range.Text).Trim([char]0x0007) -replace [char]13, ''} catch {"error"}
                                   if ($ValueOfCurrentVersionNumber -ne 'error') {
                                       if ($ValueOfCurrentVersionNumber -eq "") {[int]$ValueOfCurrentVersionNumber = 0} else {[int]$ValueOfCurrentVersionNumber = $ValueOfCurrentVersionNumber}
                                       #Write-Host "Current version: $ValueOfCurrentVersionNumber"
                                       #Write-Host "Imported version: $ValueOfVersionNumber"
                                       if ($ValueOfCurrentVersionNumber -ge $ValueOfVersionNumber -or ($ValueOfVersionNumber - $ValueOfCurrentVersionNumber) -ne 1) {$script:IncorrectVersionDiscrepancy = $true}
                                       $CurrentVersionDocumentReadData.Close([ref]0)
                                   }
                                   $ValueOfVersionNumber = $ValueOfVersionNumber - 1
                                }
                            }
                            if ($ValueOfVersionNumber -eq 0) {[string]$ValueOfVersionNumber = "-"}
                            BulkImportAdd-ItemToList -FileName $FileNameWOExtension -VersionNumber $ValueOfVersionNumber -FileType "Документ" -HighlightFlag 0 -TestPathFullName "$($script:PathToCurrentVrsion)\$($FileNameWOExtension)$($CurrentFileNameExtension)"
                        }
                    }
                    $DocumentReadData.Close([ref]0)
                }
            }
        #Проверка на тип файлам (программа)
        } else {
            #Если пользователь выбрал "Автоматическое распределение файлов по спискам", выполняется код ниже
            if ($BulkImportFormDistributeAutomatically.Checked -eq $true) {
                #Если файл существует в папке с текущей версией, то снимаем контрольную сумму ЗАМЕНЯЕМОГО файла
                if (Test-Path -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))") {
                    BulkImportAdd-ItemToList -FileName $([System.IO.Path]::GetFileName($_)) -VersionNumber (Get-FileHash -Path "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))" -Algorithm MD5).Hash -FileType "Программа" -HighlightFlag 0 -TestPathFullName "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))"
                #Если файла НЕ существует в папке с текущей версией, то снимаем контрольную сумму ЗАМЕНЯЮЩЕГО файла
                } else {
                    BulkImportAdd-ItemToList -FileName $([System.IO.Path]::GetFileName($_)) -VersionNumber (Get-FileHash -Path $_ -Algorithm MD5).Hash -FileType "Программа" -HighlightFlag 0 -TestPathFullName "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))"
                }
            #Если пользователь НЕ выбрал "Автоматическое распределение файлов по спискам", то просто снимаем контрольную сумму указанного файла
            } else {
                BulkImportAdd-ItemToList -FileName $([System.IO.Path]::GetFileName($_)) -VersionNumber (Get-FileHash -Path $_ -Algorithm MD5).Hash -FileType "Программа" -HighlightFlag 0 -TestPathFullName "$($script:PathToCurrentVrsion)\$([System.IO.Path]::GetFileName($_))"
            }
        }
    }
    $WordReadData.Quit()
    $CurrentVersionWordReadData.Quit()
    Write-Host "Пакетный импорт завершен"
}

Function Open-File ($Filter, $MultipleSelectionFlag)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = $Filter
    if ($MultipleSelectionFlag -eq $true) {$OpenFileDialog.Multiselect = $true}
    if ($MultipleSelectionFlag -eq $false) {$OpenFileDialog.Multiselect = $false}
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $OpenFileDialog.FileNames} else {return $null}
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

Function BulkImportForm ()
{
    $script:ImportFilesArray = @()
    $BulkImportForm = New-Object System.Windows.Forms.Form
    $BulkImportForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $BulkImportForm.ShowIcon = $false
    $BulkImportForm.AutoSize = $true
    $BulkImportForm.Text = "Пакетный импорт"
    $BulkImportForm.AutoSizeMode = "GrowAndShrink"
    $BulkImportForm.WindowState = "Normal"
    $BulkImportForm.SizeGripStyle = "Hide"
    $BulkImportForm.ShowInTaskbar = $true
    $BulkImportForm.StartPosition = "CenterScreen"
    $BulkImportForm.MinimizeBox = $false
    $BulkImportForm.MaximizeBox = $false
    #Кнопка обзор
    $BulkImportFormBrowseButton = New-Object System.Windows.Forms.Button
    $BulkImportFormBrowseButton.Location = New-Object System.Drawing.Point(10,10) #x,y
    $BulkImportFormBrowseButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $BulkImportFormBrowseButton.Text = "Обзор..."
    $BulkImportFormBrowseButton.TabStop = $false
    $BulkImportFormBrowseButton.Add_Click({
    $script:ImportFilesArray = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $true
    if ($script:ImportFilesArray.Count -eq 0) {
        $BulkImportFormBrowseButtonLabel.Text = "Выберите файлы для пакетного импорта"
    }
    if ($script:ImportFilesArray.Count -ne 0) {
        $BulkImportFormBrowseButtonLabel.Text = "Выбрано файлов: $($script:ImportFilesArray.Count)"
    }
    if ($script:ImportFilesArray.Count -ne 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $false) {
        $BulkImportFormApplyButton.Enabled = $true
    }
    if ($script:ImportFilesArray.Count -eq 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $false) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    if ($script:ImportFilesArray.Count -ne 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -eq $null) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    if ($script:ImportFilesArray.Count -ne 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -ne $null) {
        $BulkImportFormApplyButton.Enabled = $true
    }
    if ($script:ImportFilesArray.Count -eq 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -ne $null) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    })
    $BulkImportForm.Controls.Add($BulkImportFormBrowseButton)
    #Поле к кнопке Обзор
    $BulkImportFormBrowseButtonLabel = New-Object System.Windows.Forms.Label
    $BulkImportFormBrowseButtonLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $BulkImportFormBrowseButtonLabel.Width = 300
    $BulkImportFormBrowseButtonLabel.Text = "Выберите файлы для пакетного импорта"
    $BulkImportFormBrowseButtonLabel.TextAlign = "TopLeft"
    $BulkImportForm.Controls.Add($BulkImportFormBrowseButtonLabel)
    #Чекбокс 'Распределить файлы по спискам автоматически'
    $BulkImportFormDistributeAutomatically = New-Object System.Windows.Forms.CheckBox
    $BulkImportFormDistributeAutomatically.Width = 410
    $BulkImportFormDistributeAutomatically.Text = "Автоматически распределить файлы по спискам Выпустить и Заменить"
    $BulkImportFormDistributeAutomatically.Location = New-Object System.Drawing.Point(10,43) #x,y
    $BulkImportFormDistributeAutomatically.Enabled = $true
    $BulkImportFormDistributeAutomatically.Checked = $false
    $BulkImportFormDistributeAutomatically.Add_CheckStateChanged({
    if ($BulkImportFormDistributeAutomatically.Checked -eq $false) {
        $BulkImportFormRadioButtonRemove.Enabled = $true
        $BulkImportFormRadioButtonReplace.Enabled = $true
        $BulkImportFormRadioButtonAdd.Enabled = $true
        $BulkImportFormButtonGroupLabel.Enabled = $true
        $BulkImportFormBrowseCurrentVersionButtonLabel.Enabled = $false
        $BulkImportFormBrowseCurrentVersionButton.Enabled = $false
    } else {
        $BulkImportFormRadioButtonRemove.Enabled = $false
        $BulkImportFormRadioButtonReplace.Enabled = $false
        $BulkImportFormRadioButtonAdd.Enabled = $false
        $BulkImportFormButtonGroupLabel.Enabled = $false
        $BulkImportFormBrowseCurrentVersionButtonLabel.Enabled = $true
        $BulkImportFormBrowseCurrentVersionButton.Enabled = $true
    }
    if ($BulkImportFormDistributeAutomatically.Checked -eq $false -and $script:ImportFilesArray.Count -ne 0) {
        $BulkImportFormApplyButton.Enabled = $true
    }
    if ($BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:ImportFilesArray.Count -ne 0 -and $script:PathToCurrentVrsion -ne $null) {
        $BulkImportFormApplyButton.Enabled = $true
    }
    if ($BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:ImportFilesArray.Count -ne 0 -and $script:PathToCurrentVrsion -eq $null) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    })
    $BulkImportForm.Controls.Add($BulkImportFormDistributeAutomatically)
    #Кнопка обзор для текущей версии проекта
    $BulkImportFormBrowseCurrentVersionButton = New-Object System.Windows.Forms.Button
    $BulkImportFormBrowseCurrentVersionButton.Location = New-Object System.Drawing.Point(30,70) #x,y
    $BulkImportFormBrowseCurrentVersionButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $BulkImportFormBrowseCurrentVersionButton.Text = "Обзор..."
    $BulkImportFormBrowseCurrentVersionButton.TabStop = $false
    $BulkImportFormBrowseCurrentVersionButton.Enabled = $false
    $BulkImportFormBrowseCurrentVersionButton.Add_Click({
    $script:PathToCurrentVrsion = Select-Folder -Description "Укажите путь к папке с текущей версией проекта"
    if ($script:ImportFilesArray.Count -ne 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -eq $null) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    if ($script:ImportFilesArray.Count -ne 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -ne $null) {
        $BulkImportFormApplyButton.Enabled = $true
    }
    if ($script:ImportFilesArray.Count -eq 0 -and $BulkImportFormDistributeAutomatically.Checked -eq $true -and $script:PathToCurrentVrsion -ne $null) {
        $BulkImportFormApplyButton.Enabled = $false
    }
    if ($script:PathToCurrentVrsion -ne $null) {
        [string]$ThreeDirectories = "..."
        #$ThreeDirectories += "\$(Split-Path (Split-Path (Split-Path "$script:PathToCurrentVrsion" -Parent) -Parent) -Leaf)"
        $ThreeDirectories += "\$(Split-Path (Split-Path "$script:PathToCurrentVrsion" -Parent) -Leaf)"
        $ThreeDirectories += "\$((Get-Item "$script:PathToCurrentVrsion").Name)"
        $BulkImportFormBrowseCurrentVersionButtonLabel.Text = "Указанный путь: $ThreeDirectories"
    } else {
        $BulkImportFormBrowseCurrentVersionButtonLabel.Text = "Укажите путь к папке с текущей версией проекта"
    }
    })
    $BulkImportForm.Controls.Add($BulkImportFormBrowseCurrentVersionButton)
    #Поле к кнопке обзор для текущей версии проекта
    $BulkImportFormBrowseCurrentVersionButtonLabel = New-Object System.Windows.Forms.Label
    $BulkImportFormBrowseCurrentVersionButtonLabel.Location =  New-Object System.Drawing.Point(115,74) #x,y
    $BulkImportFormBrowseCurrentVersionButtonLabel.Width = 300
    if ($script:PathToCurrentVrsion -ne $null) {
        [string]$ThreeDirectories = "..."
        #$ThreeDirectories += "\$(Split-Path (Split-Path (Split-Path "$script:PathToCurrentVrsion" -Parent) -Parent) -Leaf)"
        $ThreeDirectories += "\$(Split-Path (Split-Path "$script:PathToCurrentVrsion" -Parent) -Leaf)"
        $ThreeDirectories += "\$((Get-Item "$script:PathToCurrentVrsion").Name)"
        $BulkImportFormBrowseCurrentVersionButtonLabel.Text = "Указанный путь: $ThreeDirectories"
    } else {
        $BulkImportFormBrowseCurrentVersionButtonLabel.Text = "Укажите путь к папке с текущей версией проекта"
    }
    $BulkImportFormBrowseCurrentVersionButtonLabel.TextAlign = "TopLeft"
    $BulkImportFormBrowseCurrentVersionButtonLabel.Enabled = $false
    $BulkImportForm.Controls.Add($BulkImportFormBrowseCurrentVersionButtonLabel)
    #Надпись к группе радиокнопок для выбора списка
    $BulkImportFormButtonGroupLabel = New-Object System.Windows.Forms.Label
    $BulkImportFormButtonGroupLabel.Location =  New-Object System.Drawing.Point(10,107) #x,y
    $BulkImportFormButtonGroupLabel.Width = 81
    $BulkImportFormButtonGroupLabel.Text = "Добавить в:"
    $BulkImportFormButtonGroupLabel.TextAlign = "TopLeft"
    $BulkImportForm.Controls.Add($BulkImportFormButtonGroupLabel)
    #Группа радиокнопок для выбора списка
    $BulkImportFormRadioButtonAdd = New-Object System.Windows.Forms.RadioButton
    $BulkImportFormRadioButtonAdd.Location = New-Object System.Drawing.Point(95,106)
    $BulkImportFormRadioButtonAdd.Size = New-Object System.Drawing.Size(250,20)
    $BulkImportFormRadioButtonAdd.Checked = $true 
    $BulkImportFormRadioButtonAdd.Text = "Выпустить"
    $BulkImportForm.Controls.Add($BulkImportFormRadioButtonAdd)
    $BulkImportFormRadioButtonReplace = New-Object System.Windows.Forms.RadioButton
    $BulkImportFormRadioButtonReplace.Location = New-Object System.Drawing.Point(95,126)
    $BulkImportFormRadioButtonReplace.Size = New-Object System.Drawing.Size(250,20)
    $BulkImportFormRadioButtonReplace.Checked = $false
    $BulkImportFormRadioButtonReplace.Text = "Заменить"
    $BulkImportForm.Controls.Add($BulkImportFormRadioButtonReplace)
    $BulkImportFormRadioButtonRemove = New-Object System.Windows.Forms.RadioButton
    $BulkImportFormRadioButtonRemove.Location = New-Object System.Drawing.Point(95,146)
    $BulkImportFormRadioButtonRemove.Size = New-Object System.Drawing.Size(250,20)
    $BulkImportFormRadioButtonRemove.Checked = $false
    $BulkImportFormRadioButtonRemove.Text = "Аннулировать"
    $BulkImportForm.Controls.Add($BulkImportFormRadioButtonRemove)
    #Кнопка Импорт
    $BulkImportFormApplyButton = New-Object System.Windows.Forms.Button
    $BulkImportFormApplyButton.Location = New-Object System.Drawing.Point(10,184) #x,y
    $BulkImportFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $BulkImportFormApplyButton.Text = "Начать"
    $BulkImportFormApplyButton.Enabled = $false
    $BulkImportFormApplyButton.Add_Click({
    $BulkImportForm.Close()
    BulkImport -ListOfSelectedFiles $script:ImportFilesArray
    $script:PathToCurrentVrsion = $null
    })
    $BulkImportForm.Controls.Add($BulkImportFormApplyButton)
    #Кнопка закрыть
    $BulkImportFormCancelButton = New-Object System.Windows.Forms.Button
    $BulkImportFormCancelButton.Location = New-Object System.Drawing.Point(100,184) #x,y
    $BulkImportFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $BulkImportFormCancelButton.Text = "Закрыть"
    $BulkImportFormCancelButton.Add_Click({
        $BulkImportForm.Close()
    })
    $BulkImportForm.Controls.Add($BulkImportFormCancelButton)
    $BulkImportForm.ShowDialog()
}

Function Add-HeaderToViewList ($ListView, $HeaderText, $Width)
{
    $ColumnHeader = New-Object System.Windows.Forms.ColumnHeader
    $ColumnHeader.Text = $HeaderText
    $ColumnHeader.Width = $Width
    $ListView.Columns.Add($ColumnHeader)
}

Function Move-ItemToAnotherList ($MoveFrom, $MoveTo) 
{
    $InheritedColor = New-Object System.Drawing.Color
    $InheritedColor = $MoVeFrom.Items[$MoVeFrom.SelectedIndices[0]].BackColor
    $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($MoVeFrom.Items[$MoVeFrom.SelectedIndices[0]].Text)")
    $ItemToAdd.SubItems.Add("$($MoVeFrom.Items[$MoVeFrom.SelectedIndices[0]].Subitems[1].Text)")
    $ItemToAdd.SubItems.Add("$($MoVeFrom.Items[$MoVeFrom.SelectedIndices[0]].Subitems[2].Text)")
    $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
    $ItemToAdd.BackColor = $InheritedColor
    $MoveTo.Items.Insert(0, $ItemToAdd)
    $MoveFrom.Items[$MoveFrom.SelectedIndices[0]].Remove()
}

Function Update-SelectedFileDetails ($FileName, $FileNameLabel, $FileAttribute, $FileAttributeLabel, $FileType, $FileTypeLabel)
{
    $FileNameLabel.Text = "Обозначение: $($FileName)"
    $FileAttributeLabel.Text = "Изм./MD5: $($FileAttribute)"
    $FileTypeLabel.Text = "Тип файла: $($FileType)"
}

Function Show-MessageBox ()
{ 
    param($Message, $Title, [ValidateSet("OK", "OKCancel", "YesNo")]$Type)
    Add-Type –AssemblyName System.Windows.Forms 
    if ($Type -eq "OK") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title")}  
    if ($Type -eq "OKCancel") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::OKCancel)}
    if ($Type -eq "YesNo") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::YesNo)}
}

Function Unselect-ItemsInOtherLists ($List1, $List2, $List3) 
{
    if ($List1.SelectedIndices.Count -gt 0) {
        for ($i = 0; $i -lt $List1.SelectedIndices.Count; $i++) {
            $List1.Items[$List1.SelectedIndices[$i]].Selected = $false
        }
    }
    if ($List2.SelectedIndices.Count -gt 0) {
        for ($i = 0; $i -lt $List2.SelectedIndices.Count; $i++) {
            $List2.Items[$List2.SelectedIndices[$i]].Selected = $false
        }
    }
    if ($List3.SelectedIndices.Count -gt 0) {
        for ($i = 0; $i -lt $List3.SelectedIndices.Count; $i++) {
            $List3.Items[$List3.SelectedIndices[$i]].Selected = $false
        }
    }
}

Function Update-ListCounters ($AddListCounter, $AddList, $ReplaceListCounter, $ReplaceList, $RemoveListCounter, $RemoveList, $TotalEntriesCounter)
{
    $TotalEntriesCounter.Text = "Всего файлов в списках: $($AddList.Items.Count + $ReplaceList.Items.Count + $RemoveList.Items.Count)"
    $AddListCounter.Text = "Выпустить ($($AddList.Items.Count)):"
    $ReplaceListCounter.Text = "Заменить ($($ReplaceList.Items.Count)):"
    $RemoveListCounter.Text = "Аннулировать ($($RemoveList.Items.Count)):"
}

Function Setup-OtherFields ()
{
    $SetupOtherFieldsForm = New-Object System.Windows.Forms.Form
    $SetupOtherFieldsForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $SetupOtherFieldsForm.ShowIcon = $false
    $SetupOtherFieldsForm.AutoSize = $true
    $SetupOtherFieldsForm.Text = "Остальные поля"
    $SetupOtherFieldsForm.AutoSizeMode = "GrowAndShrink"
    $SetupOtherFieldsForm.WindowState = "Normal"
    $SetupOtherFieldsForm.SizeGripStyle = "Hide"
    $SetupOtherFieldsForm.ShowInTaskbar = $true
    $SetupOtherFieldsForm.StartPosition = "CenterScreen"
    $SetupOtherFieldsForm.MinimizeBox = $false
    $SetupOtherFieldsForm.MaximizeBox = $false
    
    #Надпись к полю для ввода Причины
    $SetupOtherFieldsFormReasonLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormReasonLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $SetupOtherFieldsFormReasonLabel.Width = 140
    $SetupOtherFieldsFormReasonLabel.Text = "Причина (до 130 знаков):"
    $SetupOtherFieldsFormReasonLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormReasonLabel)
    #Поле для ввода Причины
    $SetupOtherFieldsFormReasonInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormReasonInput.Location = New-Object System.Drawing.Point(155,13) #x,y
    $SetupOtherFieldsFormReasonInput.Width = 450
    $SetupOtherFieldsFormReasonInput.Text = $script:GlobalReasonField
    $SetupOtherFieldsFormReasonInput.MaxLength = 130
    $SetupOtherFieldsFormReasonInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormReasonInput)
    #Надпись к полю для Указания о заделе
    $SetupOtherFieldsFormZadelLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormZadelLabel.Location =  New-Object System.Drawing.Point(10,45) #x,y
    $SetupOtherFieldsFormZadelLabel.Width = 140
    $SetupOtherFieldsFormZadelLabel.Text = "Указание о заделе:"
    $SetupOtherFieldsFormZadelLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormZadelLabel)
    #Поле для ввода обозначение
    $SetupOtherFieldsFormZadelInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormZadelInput.Location = New-Object System.Drawing.Point(155,43) #x,y
    $SetupOtherFieldsFormZadelInput.Width = 450
    $SetupOtherFieldsFormZadelInput.Text = $script:GlobalInStoreField
    $SetupOtherFieldsFormZadelInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormZadelInput)
    #Надпись к полю для Указания о внедрении
    $SetupOtherFieldsFormUsageStartLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormUsageStartLabel.Location =  New-Object System.Drawing.Point(10,75) #x,y
    $SetupOtherFieldsFormUsageStartLabel.Width = 140
    $SetupOtherFieldsFormUsageStartLabel.Text = "Указание о внедрении:"
    $SetupOtherFieldsFormUsageStartLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormUsageStartLabel)
    #Поле для ввода Указания о внедрении
    $SetupOtherFieldsFormUsageStartInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormUsageStartInput.Location = New-Object System.Drawing.Point(155,73) #x,y
    $SetupOtherFieldsFormUsageStartInput.Width = 450
    $SetupOtherFieldsFormUsageStartInput.Text = $script:GlobalStartUsageField
    $SetupOtherFieldsFormUsageStartInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormUsageStartInput)
    #Надпись к полю Применяемость
    $SetupOtherFieldsFormApplicableToLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormApplicableToLabel.Location =  New-Object System.Drawing.Point(10,105) #x,y
    $SetupOtherFieldsFormApplicableToLabel.Width = 140
    $SetupOtherFieldsFormApplicableToLabel.Text = "Применяемость:"
    $SetupOtherFieldsFormApplicableToLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormApplicableToLabel)
    #Поле для ввода Применяемость
    $SetupOtherFieldsFormApplicableToInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormApplicableToInput.Location = New-Object System.Drawing.Point(155,103) #x,y
    $SetupOtherFieldsFormApplicableToInput.Width = 450
    $SetupOtherFieldsFormApplicableToInput.Text = $script:GlobalApplicableToField
    $SetupOtherFieldsFormApplicableToInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormApplicableToInput)
    #Надпись к полю Разослать
    $SetupOtherFieldsFormSendToLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormSendToLabel.Location =  New-Object System.Drawing.Point(10,135) #x,y
    $SetupOtherFieldsFormSendToLabel.Width = 140
    $SetupOtherFieldsFormSendToLabel.Text = "Разослать:"
    $SetupOtherFieldsFormSendToLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormSendToLabel)
    #Поле для ввода Разослать
    $SetupOtherFieldsFormSendToInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormSendToInput.Location = New-Object System.Drawing.Point(155,133) #x,y
    $SetupOtherFieldsFormSendToInput.Width = 450
    $SetupOtherFieldsFormSendToInput.Text = $script:GlobalSendToField
    $SetupOtherFieldsFormSendToInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormSendToInput)
    #Надпись к полю Разослать
    $SetupOtherFieldsFormAppendixLabel = New-Object System.Windows.Forms.Label
    $SetupOtherFieldsFormAppendixLabel.Location =  New-Object System.Drawing.Point(10,165) #x,y
    $SetupOtherFieldsFormAppendixLabel.Width = 140
    $SetupOtherFieldsFormAppendixLabel.Text = "Приложение:"
    $SetupOtherFieldsFormAppendixLabel.TextAlign = "TopRight"
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormAppendixLabel)
    #Поле для ввода Разослать
    $SetupOtherFieldsFormAppendixInput = New-Object System.Windows.Forms.TextBox 
    $SetupOtherFieldsFormAppendixInput.Location = New-Object System.Drawing.Point(155,163) #x,y
    $SetupOtherFieldsFormAppendixInput.Width = 450
    $SetupOtherFieldsFormAppendixInput.Text = $script:GlobalAppendixField
    $SetupOtherFieldsFormAppendixInput.TabStop = $false
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormAppendixInput)
    #Кнопка Применить
    $SetupOtherFieldsFormApplyButton = New-Object System.Windows.Forms.Button
    $SetupOtherFieldsFormApplyButton.Location = New-Object System.Drawing.Point(10,193) #x,y
    $SetupOtherFieldsFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $SetupOtherFieldsFormApplyButton.Text = "Применить"
    $SetupOtherFieldsFormApplyButton.TabStop = $false
    $SetupOtherFieldsFormApplyButton.Add_Click({
    $script:GlobalReasonField = $SetupOtherFieldsFormReasonInput.Text
    $script:GlobalInStoreField = $SetupOtherFieldsFormZadelInput.Text
    $script:GlobalStartUsageField = $SetupOtherFieldsFormUsageStartInput.Text
    $script:GlobalApplicableToField = $SetupOtherFieldsFormApplicableToInput.Text
    $script:GlobalSendToField = $SetupOtherFieldsFormSendToInput.Text
    $script:GlobalAppendixField = $SetupOtherFieldsFormAppendixInput.Text
    $SetupOtherFieldsForm.Close()
    })
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormApplyButton)
    #Кнопка Закрыть
    $SetupOtherFieldsFormCloseButton = New-Object System.Windows.Forms.Button
    $SetupOtherFieldsFormCloseButton.Location = New-Object System.Drawing.Point(100,193) #x,y
    $SetupOtherFieldsFormCloseButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $SetupOtherFieldsFormCloseButton.Text = "Закрыть"
    $SetupOtherFieldsFormCloseButton.TabStop = $false
    $SetupOtherFieldsFormCloseButton.Add_Click({
    $SetupOtherFieldsForm.Close()
    })
    $SetupOtherFieldsForm.Controls.Add($SetupOtherFieldsFormCloseButton)
    $SetupOtherFieldsForm.ShowDialog()
}

Function Add-ItemToList ()
{
    $AddItemForm = New-Object System.Windows.Forms.Form
    $AddItemForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $AddItemForm.ShowIcon = $false
    $AddItemForm.AutoSize = $true
    $AddItemForm.Text = "Добавить запись"
    $AddItemForm.AutoSizeMode = "GrowAndShrink"
    $AddItemForm.WindowState = "Normal"
    $AddItemForm.SizeGripStyle = "Hide"
    $AddItemForm.ShowInTaskbar = $true
    $AddItemForm.StartPosition = "CenterScreen"
    $AddItemForm.MinimizeBox = $false
    $AddItemForm.MaximizeBox = $false
    #Надпись к поля для ввода обозначение
    $AddItemFormFileNameLabel = New-Object System.Windows.Forms.Label
    $AddItemFormFileNameLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $AddItemFormFileNameLabel.Width = 81
    $AddItemFormFileNameLabel.Text = "Обозначение:"
    $AddItemFormFileNameLabel.TextAlign = "TopLeft"
    $AddItemForm.Controls.Add($AddItemFormFileNameLabel)
    
    #Поле для ввода обозначение
    $AddItemFormFileNameInput = New-Object System.Windows.Forms.TextBox 
    $AddItemFormFileNameInput.Location = New-Object System.Drawing.Point(95,13) #x,y
    $AddItemFormFileNameInput.Width = 270
    $AddItemFormFileNameInput.Text = "Укажите обозначение..."
    $AddItemFormFileNameInput.ForeColor = "Gray"
    $AddItemFormFileNameInput.Add_GotFocus({
        if ($AddItemFormFileNameInput.Text -eq "Укажите обозначение...") {
            $AddItemFormFileNameInput.Text = ""
            $AddItemFormFileNameInput.ForeColor = "Black"
        }
        })
    $AddItemFormFileNameInput.Add_LostFocus({
        if ($AddItemFormFileNameInput.Text -eq "") {
            $AddItemFormFileNameInput.Text = "Укажите обозначение..."
            $AddItemFormFileNameInput.ForeColor = "Gray"
        }
        })
    $AddItemForm.Controls.Add($AddItemFormFileNameInput)
    
    #Надпись к списку для указания типа файла
    $AddItemFormFileTypeLabel = New-Object System.Windows.Forms.Label
    $AddItemFormFileTypeLabel.Location = New-Object System.Drawing.Point(10,45) #x,y
    $AddItemFormFileTypeLabel.Width = 81
    $AddItemFormFileTypeLabel.Text = "Тип файла:"
    $AddItemFormFileTypeLabel.TextAlign = "TopLeft"
    $AddItemForm.Controls.Add($AddItemFormFileTypeLabel)
    
    #Список содержащий доступные типы файлов
    $DataTypes = @("Документ","Программа")
    $AddItemFormFileTypeCombobox = New-Object System.Windows.Forms.ComboBox
    $AddItemFormFileTypeCombobox.Location = New-Object System.Drawing.Point(95,43) #x,y
    $AddItemFormFileTypeCombobox.DropDownStyle = "DropDownList"
    $DataTypes | % {$AddItemFormFileTypeCombobox.Items.add($_)}
    $AddItemFormFileTypeCombobox.SelectedIndex = 0
    $AddItemForm.Controls.Add($AddItemFormFileTypeCombobox)
    
    #Надпись к полю для ввода MD5 и Изм.
    $AddItemFormAttributeValueLabel = New-Object System.Windows.Forms.Label
    $AddItemFormAttributeValueLabel.Location =  New-Object System.Drawing.Point(10,75) #x,y
    $AddItemFormAttributeValueLabel.Width = 81
    $AddItemFormAttributeValueLabel.Text = "Изм./MD5:"
    $AddItemFormAttributeValueLabel.TextAlign = "TopLeft"
    $AddItemForm.Controls.Add($AddItemFormAttributeValueLabel)
    
    #Поле для ввода MD5 и Изм.
    $AddItemFormAttributeValueInput = New-Object System.Windows.Forms.TextBox 
    $AddItemFormAttributeValueInput.Location = New-Object System.Drawing.Point(95,73) #x,y
    $AddItemFormAttributeValueInput.Width = 270
    $AddItemFormAttributeValueInput.Text = "-"
    $AddItemFormAttributeValueInput.Add_GotFocus({
        if ($AddItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            $AddItemFormAttributeValueInput.Text = ""
            $AddItemFormAttributeValueInput.ForeColor = "Black"
        }
        })
    $AddItemFormAttributeValueInput.Add_LostFocus({
        if ($AddItemFormAttributeValueInput.Text -eq "") {
            $AddItemFormAttributeValueInput.Text = "Укажите Изм. или MD5..."
            $AddItemFormAttributeValueInput.ForeColor = "Gray"
        }
        })
    $AddItemForm.Controls.Add($AddItemFormAttributeValueInput)
    
    #Надпись к группе радиокнопок для выбора списка
    $AddItemFormRadioButtonGroupLabel = New-Object System.Windows.Forms.Label
    $AddItemFormRadioButtonGroupLabel.Location =  New-Object System.Drawing.Point(10,105) #x,y
    $AddItemFormRadioButtonGroupLabel.Width = 81
    $AddItemFormRadioButtonGroupLabel.Text = "Добавить в:"
    $AddItemFormRadioButtonGroupLabel.TextAlign = "TopLeft"
    $AddItemForm.Controls.Add($AddItemFormRadioButtonGroupLabel)
    
    #Группа радиокнопок для выбора списка
    $AddItemFormRadioButtonAdd = New-Object System.Windows.Forms.RadioButton
    $AddItemFormRadioButtonAdd.Location = New-Object System.Drawing.Point(95,104)
    $AddItemFormRadioButtonAdd.Size = New-Object System.Drawing.Size(250,20)
    $AddItemFormRadioButtonAdd.Checked = $true 
    $AddItemFormRadioButtonAdd.Text = "Выпустить"
    $AddItemForm.Controls.Add($AddItemFormRadioButtonAdd)
    $AddItemFormRadioButtonReplace = New-Object System.Windows.Forms.RadioButton
    $AddItemFormRadioButtonReplace.Location = New-Object System.Drawing.Point(95,124)
    $AddItemFormRadioButtonReplace.Size = New-Object System.Drawing.Size(250,20)
    $AddItemFormRadioButtonReplace.Checked = $false
    $AddItemFormRadioButtonReplace.Text = "Заменить"
    $AddItemForm.Controls.Add($AddItemFormRadioButtonReplace)
    $AddItemFormRadioButtonRemove = New-Object System.Windows.Forms.RadioButton
    $AddItemFormRadioButtonRemove.Location = New-Object System.Drawing.Point(95,144)
    $AddItemFormRadioButtonRemove.Size = New-Object System.Drawing.Size(250,20)
    $AddItemFormRadioButtonRemove.Checked = $false
    $AddItemFormRadioButtonRemove.Text = "Аннулировать"
    $AddItemForm.Controls.Add($AddItemFormRadioButtonRemove)
    
    #Кнопка добавить
    $AddItemFormAddButton = New-Object System.Windows.Forms.Button
    $AddItemFormAddButton.Location = New-Object System.Drawing.Point(10,180) #x,y
    $AddItemFormAddButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $AddItemFormAddButton.Text = "Добавить"
    $AddItemFormAddButton.Add_Click({
        if ($AddItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение и Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($AddItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -ne "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($AddItemFormFileNameInput.Text -ne "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } else {
            if ($AddItemFormFileNameInput.Text -match $script:BannedCharacters) {
                Show-MessageBox -Message "В обозначении запрещено использовать следующие символы:`r`n\ / ? % * : | < > """ -Title "Невозможно выполнить действие" -Type OK
            } else { 
                $ItemsOnTheList = @()
                $ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}
                $ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}
                $ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}
                if ($ItemsOnTheList -contains $AddItemFormFileNameInput.Text) {
                    Show-MessageBox -Message "Файл с указанным обозначением уже содержится в списках." -Title "Невозможно выполнить действие" -Type OK
                } else {
                    if ($AddItemFormFileTypeCombobox.SelectedItem -eq "Документ" -and $AddItemFormAttributeValueInput.Text.Length -gt 5) {
                        if ((Show-MessageBox -Message "Указанный Изм. содержит больше пяти символов. Возможно вы ошибочно указали MD5 или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указан подозрительный Изм." -Type YesNo) -eq "Yes") {
                            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($AddItemFormFileNameInput.Text)")
                            $ItemToAdd.SubItems.Add("$($AddItemFormAttributeValueInput.Text)")
                            $ItemToAdd.SubItems.Add("$($AddItemFormFileTypeCombobox.SelectedItem)")
                            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                            if ($AddItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                            if ($AddItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                            if ($AddItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                            Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
                            $AddItemForm.Close()
                        }
                    } elseif ($AddItemFormFileTypeCombobox.SelectedItem -eq "Программа" -and $AddItemFormAttributeValueInput.Text.Length -ne 32) {
                        if ((Show-MessageBox -Message "Указанная сумма MD5 некорректна. Возможно вы непольностью указали ее, ошибочно указали Изм. или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указана подозрительная MD5 сумма" -Type YesNo) -eq "Yes") {
                            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($AddItemFormFileNameInput.Text)")
                            $ItemToAdd.SubItems.Add("$($AddItemFormAttributeValueInput.Text)")
                            $ItemToAdd.SubItems.Add("$($AddItemFormFileTypeCombobox.SelectedItem)")
                            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                            if ($AddItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                            if ($AddItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                            if ($AddItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                            Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
                            $AddItemForm.Close()
                        }
                    } else {
                        $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($AddItemFormFileNameInput.Text)")
                        $ItemToAdd.SubItems.Add("$($AddItemFormAttributeValueInput.Text)")
                        $ItemToAdd.SubItems.Add("$($AddItemFormFileTypeCombobox.SelectedItem)")
                        $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                        if ($AddItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                        if ($AddItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                        if ($AddItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                        Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
                        $AddItemForm.Close()
                    }
                }
            }
        }
    })
    $AddItemForm.Controls.Add($AddItemFormAddButton)
    
    #Кнопка закрыть
    $AddItemFormCancelButton = New-Object System.Windows.Forms.Button
    $AddItemFormCancelButton.Location = New-Object System.Drawing.Point(100,180) #x,y
    $AddItemFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $AddItemFormCancelButton.Text = "Закрыть"
    $AddItemFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $AddItemFormCancelButton.Add_Click({
        $AddItemForm.Close()
    })
    $AddItemForm.Controls.Add($AddItemFormCancelButton)
    $AddItemForm.ActiveControl = $AddItemFormFileTypeLabel
    $AddItemForm.ShowDialog()
}

Function Edit-ItemOnList ($ListObject)
{
    $EditItemForm = New-Object System.Windows.Forms.Form
    $EditItemForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $EditItemForm.ShowIcon = $false
    $EditItemForm.AutoSize = $true
    $EditItemForm.Text = "Редактировать запись"
    $EditItemForm.AutoSizeMode = "GrowAndShrink"
    $EditItemForm.WindowState = "Normal"
    $EditItemForm.SizeGripStyle = "Hide"
    $EditItemForm.ShowInTaskbar = $true
    $EditItemForm.StartPosition = "CenterScreen"
    $EditItemForm.MinimizeBox = $false
    $EditItemForm.MaximizeBox = $false
    #Надпись к поля для ввода обозначение
    $EditItemFormFileNameLabel = New-Object System.Windows.Forms.Label
    $EditItemFormFileNameLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $EditItemFormFileNameLabel.Width = 81
    $EditItemFormFileNameLabel.Text = "Обозначение:"
    $EditItemFormFileNameLabel.TextAlign = "TopLeft"
    $EditItemForm.Controls.Add($EditItemFormFileNameLabel)
    
    #Поле для ввода обозначение
    $EditItemFormFileNameInput = New-Object System.Windows.Forms.TextBox 
    $EditItemFormFileNameInput.Location = New-Object System.Drawing.Point(95,13) #x,y
    $EditItemFormFileNameInput.Width = 270
    $EditItemFormFileNameInput.Text = $ListObject.Items[$ListObject.SelectedIndices[0]].Text
    $EditItemFormFileNameInput.ForeColor = "Black"
    $EditItemFormFileNameInput.Add_GotFocus({
        if ($EditItemFormFileNameInput.Text -eq "Укажите обозначение...") {
            $EditItemFormFileNameInput.Text = ""
            $EditItemFormFileNameInput.ForeColor = "Black"
        }
        })
    $EditItemFormFileNameInput.Add_LostFocus({
        if ($EditItemFormFileNameInput.Text -eq "") {
            $EditItemFormFileNameInput.Text = "Укажите обозначение..."
            $EditItemFormFileNameInput.ForeColor = "Gray"
        }
        })
    $EditItemForm.Controls.Add($EditItemFormFileNameInput)
    
    #Надпись к списку для указания типа файла
    $EditItemFormFileTypeLabel = New-Object System.Windows.Forms.Label
    $EditItemFormFileTypeLabel.Location = New-Object System.Drawing.Point(10,45) #x,y
    $EditItemFormFileTypeLabel.Width = 81
    $EditItemFormFileTypeLabel.Text = "Тип файла:"
    $EditItemFormFileTypeLabel.TextAlign = "TopLeft"
    $EditItemForm.Controls.Add($EditItemFormFileTypeLabel)
    
    #Список содержащий доступные типы файлов
    $DataTypes = @("Документ","Программа")
    $EditItemFormFileTypeCombobox = New-Object System.Windows.Forms.ComboBox
    $EditItemFormFileTypeCombobox.Location = New-Object System.Drawing.Point(95,43) #x,y
    $EditItemFormFileTypeCombobox.DropDownStyle = "DropDownList"
    $DataTypes | % {$EditItemFormFileTypeCombobox.Items.Add($_)}
    if ($ListObject.Items[$ListObject.SelectedIndices[0]].Subitems[2].Text -eq "Документ") {$EditItemFormFileTypeCombobox.SelectedIndex = 0} else {$EditItemFormFileTypeCombobox.SelectedIndex = 1}
    $EditItemForm.Controls.Add($EditItemFormFileTypeCombobox)
    
    #Надпись к полю для ввода MD5 и Изм.
    $EditItemFormAttributeValueLabel = New-Object System.Windows.Forms.Label
    $EditItemFormAttributeValueLabel.Location =  New-Object System.Drawing.Point(10,75) #x,y
    $EditItemFormAttributeValueLabel.Width = 81
    $EditItemFormAttributeValueLabel.Text = "Изм./MD5:"
    $EditItemFormAttributeValueLabel.TextAlign = "TopLeft"
    $EditItemForm.Controls.Add($EditItemFormAttributeValueLabel)
    
    #Поле для ввода MD5 и Изм.
    $EditItemFormAttributeValueInput = New-Object System.Windows.Forms.TextBox 
    $EditItemFormAttributeValueInput.Location = New-Object System.Drawing.Point(95,73) #x,y
    $EditItemFormAttributeValueInput.Width = 270
    $EditItemFormAttributeValueInput.Text = $ListObject.Items[$ListObject.SelectedIndices[0]].Subitems[1].Text
    $EditItemFormAttributeValueInput.ForeColor = "Black"
    $EditItemFormAttributeValueInput.Add_GotFocus({
        if ($EditItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            $EditItemFormAttributeValueInput.Text = ""
            $EditItemFormAttributeValueInput.ForeColor = "Black"
        }
        })
    $EditItemFormAttributeValueInput.Add_LostFocus({
        if ($EditItemFormAttributeValueInput.Text -eq "") {
            $EditItemFormAttributeValueInput.Text = "Укажите Изм. или MD5..."
            $EditItemFormAttributeValueInput.ForeColor = "Gray"
        }
        })
    $EditItemForm.Controls.Add($EditItemFormAttributeValueInput)
    
    #Кнопка применить
    $EditItemFormAddButton = New-Object System.Windows.Forms.Button
    $EditItemFormAddButton.Location = New-Object System.Drawing.Point(10,109) #x,y
    $EditItemFormAddButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EditItemFormAddButton.Text = "Применить"
    $EditItemFormAddButton.Add_Click({
        if ($EditItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение и Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($EditItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -ne "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($EditItemFormFileNameInput.Text -ne "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } else {
            if ($EditItemFormFileNameInput.Text -match $script:BannedCharacters) {
                Show-MessageBox -Message "В обозначении запрещено использовать следующие символы:`r`n\ / ? % * : | < > """ -Title "Невозможно выполнить действие" -Type OK
            } else {
                $ItemsOnTheList = @()
                $ListObject.Items | % {$ItemsOnTheList += $_.Text}
                if ($ItemsOnTheList -contains $EditItemFormFileNameInput.Text -and $EditItemFormFileNameInput.Text -ne $ListObject.Items[$ListObject.SelectedIndices[0]].Text) {
                    Show-MessageBox -Message "Файл с указанным обозначением уже содержится в списке." -Title "Невозможно выполнить действие" -Type OK
                } else {
                    if ($EditItemFormFileTypeCombobox.SelectedItem -eq "Документ" -and $EditItemFormAttributeValueInput.Text.Length -gt 5) {
                        if ((Show-MessageBox -Message "Указанный Изм. содержит больше пяти символов. Возможно вы ошибочно указали MD5 или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указан подозрительный Изм." -Type YesNo) -eq "Yes") {
                            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                            $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                            $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                            $ItemToAdd.BackColor = $ListObject.Items[$ListObject.SelectedIndices[0]].BackColor
                            $ListObject.Items.Insert($ListObject.Items[$ListObject.SelectedIndices[0]].Index, $ItemToAdd)
                            $ListObject.Items[$ListObject.SelectedIndices[0]].Remove()
                            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                            $EditItemForm.Close()
                        }
                    } elseif ($EditItemFormFileTypeCombobox.SelectedItem -eq "Программа" -and $EditItemFormAttributeValueInput.Text.Length -ne 32) {
                        if ((Show-MessageBox -Message "Указанная сумма MD5 некорректна. Возможно вы непольностью указали ее, ошибочно указали Изм. или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указана подозрительная MD5 сумма" -Type YesNo) -eq "Yes") {
                            $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                            $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                            $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                            $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                            $ItemToAdd.BackColor = $ListObject.Items[$ListObject.SelectedIndices[0]].BackColor
                            $ListObject.Items.Insert($ListObject.Items[$ListObject.SelectedIndices[0]].Index, $ItemToAdd)
                            $ListObject.Items[$ListObject.SelectedIndices[0]].Remove()
                            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                            $EditItemForm.Close()
                        }
                    } else {
                        $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                        $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                        $ItemToAdd.BackColor = $ListObject.Items[$ListObject.SelectedIndices[0]].BackColor
                        $ListObject.Items.Insert($ListObject.Items[$ListObject.SelectedIndices[0]].Index, $ItemToAdd)
                        $ListObject.Items[$ListObject.SelectedIndices[0]].Remove()
                        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                        $EditItemForm.Close()
                    }
                }
            }
        }
    })
    $EditItemForm.Controls.Add($EditItemFormAddButton)
    
    #Кнопка закрыть
    $EditItemFormCancelButton = New-Object System.Windows.Forms.Button
    $EditItemFormCancelButton.Location = New-Object System.Drawing.Point(100,109) #x,y
    $EditItemFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EditItemFormCancelButton.Text = "Закрыть"
    $EditItemFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $EditItemFormCancelButton.Add_Click({
    $EditItemForm.Close()
    })
    $EditItemForm.Controls.Add($EditItemFormCancelButton)
    $EditItemForm.ActiveControl = $EditItemFormFileTypeLabel
    $EditItemForm.ShowDialog()
}

Function Clear-Lists ()
{
    #Окно Очистить списки
    $ClearListsForm = New-Object System.Windows.Forms.Form
    $ClearListsForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $ClearListsForm.ShowIcon = $false
    $ClearListsForm.AutoSize = $true
    $ClearListsForm.Text = "Очистить списки"
    $ClearListsForm.AutoSizeMode = "GrowAndShrink"
    $ClearListsForm.WindowState = "Normal"
    $ClearListsForm.SizeGripStyle = "Hide"
    $ClearListsForm.ShowInTaskbar = $true
    $ClearListsForm.StartPosition = "CenterScreen"
    $ClearListsForm.MinimizeBox = $false
    $ClearListsForm.MaximizeBox = $false
    #Чекбокс Очистить список 'Выпустить'
    $CheckboxClearAddList = New-Object System.Windows.Forms.CheckBox
    $CheckboxClearAddList.Width = 350
    $CheckboxClearAddList.Text = "Очистить список 'Выпустить'"
    $CheckboxClearAddList.Location = New-Object System.Drawing.Point(10,15) #x,y
    $CheckboxClearAddList.Enabled = $true
    $CheckboxClearAddList.Checked = $true
    $CheckboxClearAddList.Add_CheckStateChanged({})
    $ClearListsForm.Controls.Add($CheckboxClearAddList)
    #Чекбокс Очистить список 'Заменить'
    $CheckboxClearReplaceList = New-Object System.Windows.Forms.CheckBox
    $CheckboxClearReplaceList.Width = 350
    $CheckboxClearReplaceList.Text = "Очистить список 'Заменить'"
    $CheckboxClearReplaceList.Location = New-Object System.Drawing.Point(10,45) #x,y
    $CheckboxClearReplaceList.Enabled = $true
    $CheckboxClearReplaceList.Checked = $true
    $CheckboxClearReplaceList.Add_CheckStateChanged({})
    $ClearListsForm.Controls.Add($CheckboxClearReplaceList)
    #Чекбокс Очистить список 'Аннулировать'
    $CheckboxClearRemoveList = New-Object System.Windows.Forms.CheckBox
    $CheckboxClearRemoveList.Width = 350
    $CheckboxClearRemoveList.Text = "Очистить список 'Аннулировать'"
    $CheckboxClearRemoveList.Location = New-Object System.Drawing.Point(10,75) #x,y
    $CheckboxClearRemoveList.Enabled = $true
    $CheckboxClearRemoveList.Checked = $true
    $CheckboxClearRemoveList.Add_CheckStateChanged({})
    $ClearListsForm.Controls.Add($CheckboxClearRemoveList)
    #Кнопка применить
    $ClearListsFormAddButton = New-Object System.Windows.Forms.Button
    $ClearListsFormAddButton.Location = New-Object System.Drawing.Point(10,109) #x,y
    $ClearListsFormAddButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClearListsFormAddButton.Text = "Применить"
    $ClearListsFormAddButton.Add_Click({
       if ($CheckboxClearAddList.Checked -eq $true) {$ListViewAdd.Items.Clear()}
       if ($CheckboxClearReplaceList.Checked -eq $true) {$ListViewReplace.Items.Clear()}
       if ($CheckboxClearRemoveList.Checked -eq $true) {$ListViewRemove.Items.Clear()}
       Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
       $ClearListsForm.Close()
    })
    $ClearListsForm.Controls.Add($ClearListsFormAddButton)
    #Кнопка закрыть
    $ClearListsFormCancelButton = New-Object System.Windows.Forms.Button
    $ClearListsFormCancelButton.Location = New-Object System.Drawing.Point(100,109) #x,y
    $ClearListsFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClearListsFormCancelButton.Text = "Закрыть"
    $ClearListsFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ClearListsFormCancelButton.Add_Click({
    $ClearListsForm.Close()
    })
    $ClearListsForm.Controls.Add($ClearListsFormCancelButton)
    $ClearListsForm.ShowDialog()
}

Function Discard-Coloring ()
{
    #Окно Очистить списки
    $DiscardColoringForm = New-Object System.Windows.Forms.Form
    $DiscardColoringForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $DiscardColoringForm.ShowIcon = $false
    $DiscardColoringForm.AutoSize = $true
    $DiscardColoringForm.Text = "Отменить выделение цветом в списках"
    $DiscardColoringForm.AutoSizeMode = "GrowAndShrink"
    $DiscardColoringForm.WindowState = "Normal"
    $DiscardColoringForm.SizeGripStyle = "Hide"
    $DiscardColoringForm.ShowInTaskbar = $true
    $DiscardColoringForm.StartPosition = "CenterScreen"
    $DiscardColoringForm.MinimizeBox = $false
    $DiscardColoringForm.MaximizeBox = $false
    #Чекбокс Очистить список 'Выпустить'
    $CheckboxDiscardColoringInAddList = New-Object System.Windows.Forms.CheckBox
    $CheckboxDiscardColoringInAddList.Width = 350
    $CheckboxDiscardColoringInAddList.Text = "Отменить выделение цветом в списке 'Выпустить'"
    $CheckboxDiscardColoringInAddList.Location = New-Object System.Drawing.Point(10,15) #x,y
    $CheckboxDiscardColoringInAddList.Enabled = $true
    $CheckboxDiscardColoringInAddList.Checked = $true
    $CheckboxDiscardColoringInAddList.Add_CheckStateChanged({})
    $DiscardColoringForm.Controls.Add($CheckboxDiscardColoringInAddList)
    #Чекбокс Очистить список 'Заменить'
    $CheckboxDiscardColoringInReplaceList = New-Object System.Windows.Forms.CheckBox
    $CheckboxDiscardColoringInReplaceList.Width = 350
    $CheckboxDiscardColoringInReplaceList.Text = "Отменить выделение цветом в списке 'Заменить'"
    $CheckboxDiscardColoringInReplaceList.Location = New-Object System.Drawing.Point(10,45) #x,y
    $CheckboxDiscardColoringInReplaceList.Enabled = $true
    $CheckboxDiscardColoringInReplaceList.Checked = $true
    $CheckboxDiscardColoringInReplaceList.Add_CheckStateChanged({})
    $DiscardColoringForm.Controls.Add($CheckboxDiscardColoringInReplaceList)
    #Чекбокс Очистить список 'Аннулировать'
    $CheckboxDiscardColoringInRemoveList = New-Object System.Windows.Forms.CheckBox
    $CheckboxDiscardColoringInRemoveList.Width = 350
    $CheckboxDiscardColoringInRemoveList.Text = "Отменить выделение цветом в списке 'Аннулировать'"
    $CheckboxDiscardColoringInRemoveList.Location = New-Object System.Drawing.Point(10,75) #x,y
    $CheckboxDiscardColoringInRemoveList.Enabled = $true
    $CheckboxDiscardColoringInRemoveList.Checked = $true
    $CheckboxDiscardColoringInRemoveList.Add_CheckStateChanged({})
    $DiscardColoringForm.Controls.Add($CheckboxDiscardColoringInRemoveList)
    #Кнопка применить
    $DiscardColoringFormAddButton = New-Object System.Windows.Forms.Button
    $DiscardColoringFormAddButton.Location = New-Object System.Drawing.Point(10,109) #x,y
    $DiscardColoringFormAddButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $DiscardColoringFormAddButton.Text = "Применить"
    $DiscardColoringFormAddButton.Add_Click({
       if ($CheckboxDiscardColoringInAddList.Checked -eq $true) {foreach ($Item in $ListViewAdd.Items) {$Item.BackColor = [System.Drawing.Color]::White}}
       if ($CheckboxDiscardColoringInReplaceList.Checked -eq $true) {foreach ($Item in $ListViewReplace.Items) {$Item.BackColor = [System.Drawing.Color]::White}}
       if ($CheckboxDiscardColoringInRemoveList.Checked -eq $true) {foreach ($Item in $ListViewRemove.Items) {$Item.BackColor = [System.Drawing.Color]::White}}
       $DiscardColoringForm.Close()
    })
    $DiscardColoringForm.Controls.Add($DiscardColoringFormAddButton)
    #Кнопка закрыть
    $DiscardColoringFormCancelButton = New-Object System.Windows.Forms.Button
    $DiscardColoringFormCancelButton.Location = New-Object System.Drawing.Point(100,109) #x,y
    $DiscardColoringFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $DiscardColoringFormCancelButton.Text = "Закрыть"
    $DiscardColoringFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $DiscardColoringFormCancelButton.Add_Click({
    $DiscardColoringForm.Close()
    })
    $DiscardColoringForm.Controls.Add($DiscardColoringFormCancelButton)
    $DiscardColoringForm.ShowDialog()
}

Function Disable-AllExceptEditing ($BooleanRest, $BooleanEditing) 
{
    $GetPropertyListBoxBlackList.Enabled = $BooleanRest
    $GetPropertyButtonAddItem.Enabled = $BooleanRest
    $GetPropertyInputboxAddItem.Enabled = $BooleanRest
    $GetPropertyButtonDeleteItem.Enabled = $BooleanRest
    $GetPropertyLabelButtonDelete.Enabled = $BooleanRest
    $GetPropertyInputboxEditItem.Enabled = $BooleanEditing
    $GetPropertyButtonApplyItem.Enabled = $BooleanEditing
    $GetPropertyButtonCancelItem.Enabled = $BooleanEditing
    $GetPropertyButtonEditItem.Enabled = $BooleanRest
    $GetPropertyButtonClear.Enabled = $BooleanRest
    $GetPropertyLabelButtonClear.Enabled = $BooleanRest
    #$ManageCustomListsCloseButton.Enabled = $BooleanRest
    $ManageCustomListsSaveButton.Enabled = $BooleanRest
}

Function Populate-List ($List, $PathToXml)
{
    $ServiceXml = New-Object System.Xml.XmlDocument
    $ServiceXml.Load($PathToXml)
    $InnerTexts = $ServiceXml.SelectNodes("//name")
    Foreach ($InnerText in $InnerTexts) {
        $List.Items.Add($InnerText.InnerText)
    }
}

Function Generate-XmlList ($List, [ValidateSet("Departments", "Employees", "Projects", "Developers")]$ListType)
{
    $XmlList = New-Object System.Xml.XmlDocument
    $XmlList.CreateXmlDeclaration("1.0","UTF-8",$null)
    $XmlList.AppendChild($XmlList.CreateXmlDeclaration("1.0","UTF-8",$null))
$CommentForXml = @"
Автоматически сгенерированный список
Сгенерирован: $(Get-Date)
"@
    $XmlList.AppendChild($XmlList.CreateComment($CommentForXml))
    if ($ListType -eq "Departments") {$RootElement = $XmlList.CreateNode("element","departments",$null)}
    if ($ListType -eq "Employees") {$RootElement = $XmlList.CreateNode("element","employees",$null)}
    if ($ListType -eq "Projects") {$RootElement = $XmlList.CreateNode("element","projects",$null)}
    if ($ListType -eq "Developers") {$RootElement = $XmlList.CreateNode("element","developers",$null)}
    $XmlList.AppendChild($RootElement) | Out-Null
    Foreach ($ListItem in $List) {
        $ElementName = $XmlList.CreateNode("element","name",$null)
        $ElementName.InnerText = $ListItem
        if ($ListType -eq "Departments") {$XmlList.SelectSingleNode("/departments").AppendChild($ElementName)}
        if ($ListType -eq "Employees") {$XmlList.SelectSingleNode("/employees").AppendChild($ElementName)}
        if ($ListType -eq "Projects") {$XmlList.SelectSingleNode("/projects").AppendChild($ElementName)}
        if ($ListType -eq "Developers") {$XmlList.SelectSingleNode("/developers").AppendChild($ElementName)}
    }
    if ($ListType -eq "Departments") {$XmlList.Save("$PSScriptRoot\Отделы.xml")}
    if ($ListType -eq "Employees") {$XmlList.Save("$PSScriptRoot\Сотрудники.xml")}
    if ($ListType -eq "Projects") {$XmlList.Save("$PSScriptRoot\Проекты.xml")}
    if ($ListType -eq "Developers") {$XmlList.Save("$PSScriptRoot\Разработчики.xml")}
}

Function Edit-ElementFromExistingList ($EditedElement, $NewValue, [ValidateSet("Departments", "Employees", "Projects", "Developers")]$ListType)
{
    if ($ListType -eq "Departments") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Отделы.xml")
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").Attributes.RemoveAll()
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").InnerXml = "$NewValue"
        $XmlEditElement.Save("$PSScriptRoot\Отделы.xml")
    }
    if ($ListType -eq "Employees") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Сотрудники.xml")
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").Attributes.RemoveAll()
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").InnerXml = "$NewValue"
        $XmlEditElement.Save("$PSScriptRoot\Сотрудники.xml")
    }
    if ($ListType -eq "Projects") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Проекты.xml")
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").Attributes.RemoveAll()
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").InnerXml = "$NewValue"
        $XmlEditElement.Save("$PSScriptRoot\Проекты.xml")
    }
    if ($ListType -eq "Developers") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Разработчики.xml")
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").Attributes.RemoveAll()
        $XmlEditElement.SelectSingleNode("//name[.='$EditedElement']").InnerXml = "$NewValue"
        $XmlEditElement.Save("$PSScriptRoot\Разработчики.xml")
    }
}

Function Add-NewElementsToExistingList ($List, [ValidateSet("Departments", "Employees", "Projects", "Developers")]$ListType) 
{
    if ($ListType -eq "Departments") {
        $XmlAddNewElement = New-Object System.Xml.XmlDocument
        $XmlAddNewElement.Load("$PSScriptRoot\Отделы.xml")
        Foreach ($ListItem in $List) {
            if ($XmlAddNewElement.SelectNodes("//name[.='$ListItem']").Count -eq 0) {
                $NewElement = $XmlAddNewElement.CreateNode("element","name",$null)
                $NewElement.InnerXml = "$ListItem"
                $XmlAddNewElement.SelectSingleNode("/departments").AppendChild($NewElement)
            }
        }
        $XmlAddNewElement.Save("$PSScriptRoot\Отделы.xml")
    }
    if ($ListType -eq "Projects") {
        $XmlAddNewElement = New-Object System.Xml.XmlDocument
        $XmlAddNewElement.Load("$PSScriptRoot\Проекты.xml")
        Foreach ($ListItem in $List) {
            if ($XmlAddNewElement.SelectNodes("//name[.='$ListItem']").Count -eq 0) {
                $NewElement = $XmlAddNewElement.CreateNode("element","name",$null)
                $NewElement.InnerXml = "$ListItem"
                $XmlAddNewElement.SelectSingleNode("/projects").AppendChild($NewElement)
            }
        }
        $XmlAddNewElement.Save("$PSScriptRoot\Проекты.xml")
    }
    if ($ListType -eq "Employees") {
        $XmlAddNewElement = New-Object System.Xml.XmlDocument
        $XmlAddNewElement.Load("$PSScriptRoot\Сотрудники.xml")
        Foreach ($ListItem in $List) {
            if ($XmlAddNewElement.SelectNodes("//name[.='$ListItem']").Count -eq 0) {
                $NewElement = $XmlAddNewElement.CreateNode("element","name",$null)
                $NewElement.InnerXml = "$ListItem"
                $XmlAddNewElement.SelectSingleNode("/employees").AppendChild($NewElement)
            }
        }
        $XmlAddNewElement.Save("$PSScriptRoot\Сотрудники.xml")
    }
    if ($ListType -eq "Developers") {
        $XmlAddNewElement = New-Object System.Xml.XmlDocument
        $XmlAddNewElement.Load("$PSScriptRoot\Разработчики.xml")
        Foreach ($ListItem in $List) {
            if ($XmlAddNewElement.SelectNodes("//name[.='$ListItem']").Count -eq 0) {
                $NewElement = $XmlAddNewElement.CreateNode("element","name",$null)
                $NewElement.InnerXml = "$ListItem"
                $XmlAddNewElement.SelectSingleNode("/developers").AppendChild($NewElement)
            }
        }
        $XmlAddNewElement.Save("$PSScriptRoot\Разработчики.xml")
    }
}

Function Remove-ElementFromExistingList ($ElementToDelete, [ValidateSet("Departments", "Employees", "Projects", "Developers")]$ListType)
{
    if ($ListType -eq "Departments") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Отделы.xml")
        $XmlEditElement.SelectSingleNode("/departments").RemoveChild($XmlEditElement.SelectSingleNode("//name[.='$ElementToDelete']"))
        $XmlEditElement.Save("$PSScriptRoot\Отделы.xml")
    }
    if ($ListType -eq "Employees") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Сотрудники.xml")
        $XmlEditElement.SelectSingleNode("/employees").RemoveChild($XmlEditElement.SelectSingleNode("//name[.='$ElementToDelete']"))
        $XmlEditElement.Save("$PSScriptRoot\Сотрудники.xml")
    }
    if ($ListType -eq "Projects") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Проекты.xml")
        $XmlEditElement.SelectSingleNode("/projects").RemoveChild($XmlEditElement.SelectSingleNode("//name[.='$ElementToDelete']"))
        $XmlEditElement.Save("$PSScriptRoot\Проекты.xml")
    }
    if ($ListType -eq "Developers") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Разработчики.xml")
        $XmlEditElement.SelectSingleNode("/developers").RemoveChild($XmlEditElement.SelectSingleNode("//name[.='$ElementToDelete']"))
        $XmlEditElement.Save("$PSScriptRoot\Разработчики.xml")
    }
}

Function Remove-AllItemsFromExistingList ([ValidateSet("Departments", "Employees", "Projects", "Developers")]$ListType)
{
    if ($ListType -eq "Departments") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Отделы.xml")
        $XmlEditElement.SelectSingleNode("/departments").RemoveAll()
        $XmlEditElement.Save("$PSScriptRoot\Отделы.xml")
    }
    if ($ListType -eq "Employees") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Сотрудники.xml")
        $XmlEditElement.SelectSingleNode("/employees").RemoveAll()
        $XmlEditElement.Save("$PSScriptRoot\Сотрудники.xml")
    }
    if ($ListType -eq "Projects") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Проекты.xml")
        $XmlEditElement.SelectSingleNode("/projects").RemoveAll()
        $XmlEditElement.Save("$PSScriptRoot\Проекты.xml")
    }
    if ($ListType -eq "Developers") {
        $XmlEditElement = New-Object System.Xml.XmlDocument
        $XmlEditElement.Load("$PSScriptRoot\Разработчики.xml")
        $XmlEditElement.SelectSingleNode("/developers").RemoveAll()
        $XmlEditElement.Save("$PSScriptRoot\Разработчики.xml")
    }
}

Function Manage-CustomLists ($PathToLost, [ValidateSet("Departments", "Employees", "Projects", "RegisterProjects", "Developers")]$ListType)
{
    $ManageCustomListsForm = New-Object System.Windows.Forms.Form
    $ManageCustomListsForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ManageCustomListsForm.ShowIcon = $false
    $ManageCustomListsForm.AutoSize = $true
    if ($ListType -eq "Departments") {$ManageCustomListsForm.Text = "Редактировать список отделов"}
    if ($ListType -eq "Employees") {$ManageCustomListsForm.Text = "Редактировать список сотрудников"}
    if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {$ManageCustomListsForm.Text = "Редактировать список проектов"}
    if ($ListType -eq "Developers") {$ManageCustomListsForm.Text = "Редактировать список разработчиков"}
    $ManageCustomListsForm.AutoSizeMode = "GrowAndShrink"
    $ManageCustomListsForm.WindowState = "Normal"
    $ManageCustomListsForm.SizeGripStyle = "Hide"
    $ManageCustomListsForm.ShowInTaskbar = $true
    $ManageCustomListsForm.StartPosition = "CenterScreen"
    $ManageCustomListsForm.MinimizeBox = $false
    $ManageCustomListsForm.MaximizeBox = $false
    #Надпись к списку, который содержит список отделов/сотрудников компании
    $GetPropertyLabelBlacklistListBox = New-Object System.Windows.Forms.Label
    $GetPropertyLabelBlacklistListBox.Location =  New-Object System.Drawing.Point(10,10) #x,y
    $GetPropertyLabelBlacklistListBox.Width = 250
    $GetPropertyLabelBlacklistListBox.Height = 13
    if ($ListType -eq "Departments") {$GetPropertyLabelBlacklistListBox.Text = "Список отделов компании:"}
    if ($ListType -eq "Employees") {$GetPropertyLabelBlacklistListBox.Text = "Список сотрудников компании:"}
    if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {$GetPropertyLabelBlacklistListBox.Text = "Список проектов компании:"}
    if ($ListType -eq "Developers") {$GetPropertyLabelBlacklistListBox.Text = "Список разработчиков:"}
    $ManageCustomListsForm.Controls.Add($GetPropertyLabelBlacklistListBox)
    #Список отделов/сотрудников компании
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(10,25) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,260) #width,height
    if ($ListType -eq "Departments") {if (Test-Path "$PSScriptRoot\Отделы.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Отделы.xml"}}
    if ($ListType -eq "Employees") {if (Test-Path "$PSScriptRoot\Сотрудники.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Сотрудники.xml"}}
    if ($ListType -eq "Projects"  -or $ListType -eq "RegisterProjects") {if (Test-Path "$PSScriptRoot\Проекты.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Проекты.xml"}}
    if ($ListType -eq "Developers") {if (Test-Path "$PSScriptRoot\Разработчики.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Разработчики.xml"}}
    $GetPropertyListBoxBlackList.Add_SelectedIndexChanged({
        if ($GetPropertyListBoxBlackList.SelectedIndex -ne -1) {
            $GetPropertyInputboxEditItem.Text = $GetPropertyListBoxBlackList.SelectedItem
        }
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyListBoxBlackList)
    #Кнопка добавить
    $GetPropertyButtonAddItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonAddItem.Location = New-Object System.Drawing.Point(235,25) #x,y
    $GetPropertyButtonAddItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $GetPropertyButtonAddItem.Text = "Добавить"
    $GetPropertyButtonAddItem.Add_Click({
        if ($ListType -eq "Departments") {
            if ($GetPropertyInputboxAddItem.Text -ne "Укажите название отдела...") {
                if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                    Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанный отдел ($($GetPropertyInputboxAddItem.Text))."
                } else {
                    $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                    $GetPropertyInputboxAddItem.Text = "Укажите название отдела..."
                    $GetPropertyInputboxAddItem.ForeColor = "Gray"
                    if (Test-Path "$PSScriptRoot\Отделы.xml") {Add-NewElementsToExistingList -List $GetPropertyListBoxBlackList.Items -ListType Departments} else {Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Departments}
                    $ComboboxDepartmentName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $ComboboxDepartmentName.Items.Add($ItemInList)
                    }
                }
            }
        }
        if ($ListType -eq "Employees") {
            if ($GetPropertyInputboxAddItem.Text -ne "Укажите ФИО...") {
                if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанное ФИО ($($GetPropertyInputboxAddItem.Text))."
                } else {
                    $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                    $GetPropertyInputboxAddItem.Text = "Укажите ФИО..."
                    $GetPropertyInputboxAddItem.ForeColor = "Gray"
                    if (Test-Path "$PSScriptRoot\Сотрудники.xml") {Add-NewElementsToExistingList -List $GetPropertyListBoxBlackList.Items -ListType Employees} else {Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Employees}
                        $ComboboxCheckedBy.Items.Clear()
                        $ComboboxCreatedBy.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $ComboboxCheckedBy.Items.Add($ItemInList)
                        $ComboboxCreatedBy.Items.Add($ItemInList)
                    }               
                }
            }
        }
        if ($ListType -eq "Projects") {
            if ($GetPropertyInputboxAddItem.Text -ne "Укажите название проекта...") {
                if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанный проект ($($GetPropertyInputboxAddItem.Text))."
                } else {
                    $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                    $GetPropertyInputboxAddItem.Text = "Укажите название проекта..."
                    $GetPropertyInputboxAddItem.ForeColor = "Gray"                   
                    $EmailSubjectAccessPathInput.Text = ""
                    if (Test-Path "$PSScriptRoot\Проекты.xml") {Add-NewElementsToExistingList -List $GetPropertyListBoxBlackList.Items -ListType Projects} else {Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Projects}
                    $EmailSubjectComboboxProjectName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $EmailSubjectComboboxProjectName.Items.Add($ItemInList)
                    }
                    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
                }
            }
        }
        if ($ListType -eq "RegisterProjects") {
            if ($GetPropertyInputboxAddItem.Text -ne "Укажите название проекта...") {
                if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                    Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанный проект ($($GetPropertyInputboxAddItem.Text))."
                } else {
                    $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                    $GetPropertyInputboxAddItem.Text = "Укажите название проекта..."
                    $GetPropertyInputboxAddItem.ForeColor = "Gray"
                    if (Test-Path "$PSScriptRoot\Проекты.xml") {Add-NewElementsToExistingList -List $GetPropertyListBoxBlackList.Items -ListType Projects} else {Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Projects}
                    $UpdateRegisterFormComboboxProjectName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $UpdateRegisterFormComboboxProjectName.Items.Add($ItemInList)
                    }                  
                }
            }
        }
        if ($ListType -eq "Developers") {
            if ($GetPropertyInputboxAddItem.Text -ne "Укажите название разработчика...") {
                if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                    Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанного разработчика ($($GetPropertyInputboxAddItem.Text))."
                } else {
                    $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                    $GetPropertyInputboxAddItem.Text = "Укажите название разработчика..."
                    $GetPropertyInputboxAddItem.ForeColor = "Gray"
                    if (Test-Path "$PSScriptRoot\Разработчики.xml") {Add-NewElementsToExistingList -List $GetPropertyListBoxBlackList.Items -ListType Developers} else {Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Developers}
                    $UpdateRegisterFormComboboxDeveloperName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $UpdateRegisterFormComboboxDeveloperName.Items.Add($ItemInList)
                    }
                }
            }
        }
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonAddItem)
    #Поле ввода для указания названия отдела
    $GetPropertyInputboxAddItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxAddItem.Location = New-Object System.Drawing.Size(350,26) #x,y
    $GetPropertyInputboxAddItem.Width = 190
    if ($ListType -eq "Departments") {$GetPropertyInputboxAddItem.Text = "Укажите название отдела..."}
    if ($ListType -eq "Employees") {$GetPropertyInputboxAddItem.Text = "Укажите ФИО..."}
    if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {$GetPropertyInputboxAddItem.Text = "Укажите название проекта..."}
    if ($ListType -eq "Developers") {$GetPropertyInputboxAddItem.Text = "Укажите название разработчика..."}
    $GetPropertyInputboxAddItem.ForeColor = "Gray"
    $GetPropertyInputboxAddItem.Add_GotFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "Укажите название отдела..." -or $GetPropertyInputboxAddItem.Text -eq "Укажите ФИО..." -or $GetPropertyInputboxAddItem.Text -eq "Укажите название проекта..." -or  $GetPropertyInputboxAddItem.Text -eq "Укажите название разработчика...") {
            $GetPropertyInputboxAddItem.Text = ""
            $GetPropertyInputboxAddItem.ForeColor = "Black"
        }
        })
    $GetPropertyInputboxAddItem.Add_LostFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "") {
            if ($ListType -eq "Departments") {$GetPropertyInputboxAddItem.Text = "Укажите название отдела..."}
            if ($ListType -eq "Employees") {$GetPropertyInputboxAddItem.Text = "Укажите ФИО..."}
            if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {$GetPropertyInputboxAddItem.Text = "Укажите название проекта..."}
            if ($ListType -eq "Developers") {$GetPropertyInputboxAddItem.Text = "Укажите название разработчика..."}
            $GetPropertyInputboxAddItem.ForeColor = "Gray"
        }
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyInputboxAddItem)
    #Кнопка редактировать
    $GetPropertyButtonEditItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonEditItem.Location = New-Object System.Drawing.Point(235,53) #x,y
    $GetPropertyButtonEditItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $GetPropertyButtonEditItem.Text = "Редактировать"
    $GetPropertyButtonEditItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -ne "Выберите запись из списка...") {
            Disable-AllExceptEditing -BooleanRest $false -BooleanEditing $true
        }
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonEditItem)
    #Поле ввода для редактирования названия отдела
    $GetPropertyInputboxEditItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxEditItem.Location = New-Object System.Drawing.Size(350,54) #x,y
    $GetPropertyInputboxEditItem.Width = 190
    $GetPropertyInputboxEditItem.Enabled = $false
    $GetPropertyInputboxEditItem.Text = "Выберите запись из списка..."
    $ManageCustomListsForm.Controls.Add($GetPropertyInputboxEditItem)
    #Кнопка применить
    $GetPropertyButtonApplyItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonApplyItem.Location = New-Object System.Drawing.Point(350,76) #x,y
    $GetPropertyButtonApplyItem.Size = New-Object System.Drawing.Point(80,22) #width,height
    $GetPropertyButtonApplyItem.Text = "Применить"
    $GetPropertyButtonApplyItem.Enabled = $false
    $GetPropertyButtonApplyItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -eq $GetPropertyListBoxBlackList.SelectedItem) {
            Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
        } else {
            if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxEditItem.Text)) {
                if ($ListType -eq "Departments") {Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанный отдел ($($GetPropertyInputboxEditItem.Text))."}
                if ($ListType -eq "Employees") {Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанное ФИО ($($GetPropertyInputboxEditItem.Text))."}
                if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанный проект ($($GetPropertyInputboxEditItem.Text))."}
                if ($ListType -eq "Developers") {Show-MessageBox -Title "Запись уже существует в списке" -Type OK -Message "Список уже содержит указанного разработчика ($($GetPropertyInputboxEditItem.Text))."}
            } else {
                if ($ListType -eq "Departments") {Edit-ElementFromExistingList -EditedElement $GetPropertyListBoxBlackList.SelectedItem -NewValue $GetPropertyInputboxEditItem.Text -ListType Departments}
                if ($ListType -eq "Employees") {Edit-ElementFromExistingList -EditedElement $GetPropertyListBoxBlackList.SelectedItem -NewValue $GetPropertyInputboxEditItem.Text -ListType Employees}
                if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {Edit-ElementFromExistingList -EditedElement $GetPropertyListBoxBlackList.SelectedItem -NewValue $GetPropertyInputboxEditItem.Text -ListType Projects}
                if ($ListType -eq "Developers") {Edit-ElementFromExistingList -EditedElement $GetPropertyListBoxBlackList.SelectedItem -NewValue $GetPropertyInputboxEditItem.Text -ListType Developers}
                $SelectedIndex = $GetPropertyListBoxBlackList.SelectedIndex
                $GetPropertyListBoxBlackList.Items.Insert($GetPropertyListBoxBlackList.SelectedIndex, ($GetPropertyInputboxEditItem.Text).Trim(' '))
                $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
                $GetPropertyListBoxBlackList.SelectedIndex = $SelectedIndex
                $GetPropertyInputboxEditItem.Text = ($GetPropertyInputboxEditItem.Text).Trim(' ')
                Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
                if ($ListType -eq "Departments") {
                    $ComboboxDepartmentName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $ComboboxDepartmentName.Items.Add($ItemInList)
                    }
                }
                if ($ListType -eq "Employees") {
                    $ComboboxCheckedBy.Items.Clear()
                    $ComboboxCreatedBy.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $ComboboxCheckedBy.Items.Add($ItemInList)
                        $ComboboxCreatedBy.Items.Add($ItemInList)
                    }
                }
                if ($ListType -eq "Projects") {
                    $EmailSubjectAccessPathInput.Text = ""
                    $EmailSubjectComboboxProjectName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $EmailSubjectComboboxProjectName.Items.Add($ItemInList)
                    }
                    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
                }
                if ($ListType -eq "RegisterProjects") {
                    $UpdateRegisterFormComboboxProjectName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $UpdateRegisterFormComboboxProjectName.Items.Add($ItemInList)
                    }  
                }
                if ($ListType -eq "Developers") {
                    $UpdateRegisterFormComboboxDeveloperName.Items.Clear()
                    Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                        $UpdateRegisterFormComboboxDeveloperName.Items.Add($ItemInList)
                    }  
                } 
            }
        }
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonApplyItem)
    #Кнопка отмена
    $GetPropertyButtonCancelItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonCancelItem.Location = New-Object System.Drawing.Point(432,76) #x,y
    $GetPropertyButtonCancelItem.Size = New-Object System.Drawing.Point(80,22) #width,height
    $GetPropertyButtonCancelItem.Text = "Отмена"
    $GetPropertyButtonCancelItem.Enabled = $false
    $GetPropertyButtonCancelItem.Add_Click({
        $GetPropertyInputboxEditItem.Text = $GetPropertyListBoxBlackList.SelectedItem
        Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false    
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonCancelItem)
    #Кнопка удалить
    $GetPropertyButtonDeleteItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonDeleteItem.Location = New-Object System.Drawing.Point(235,104) #x,y
    $GetPropertyButtonDeleteItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $GetPropertyButtonDeleteItem.Text = "Удалить"
    $GetPropertyButtonDeleteItem.Add_Click({
        if ($ListType -eq "Departments") {Remove-ElementFromExistingList -ElementToDelete $GetPropertyListBoxBlackList.SelectedItem -ListType Departments}
        if ($ListType -eq "Employees") {Remove-ElementFromExistingList -ElementToDelete $GetPropertyListBoxBlackList.SelectedItem -ListType Employees}
        if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {Remove-ElementFromExistingList -ElementToDelete $GetPropertyListBoxBlackList.SelectedItem -ListType Projects}
        if ($ListType -eq "Developers") {Remove-ElementFromExistingList -ElementToDelete $GetPropertyListBoxBlackList.SelectedItem -ListType Developers}
        $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
        $GetPropertyInputboxEditItem.Text = "Выберите запись из списка..."
        if ($ListType -eq "Departments") {
            $ComboboxDepartmentName.Items.Clear()
            Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                $ComboboxDepartmentName.Items.Add($ItemInList)
            }
        }
        if ($ListType -eq "Employees") {
            $ComboboxCheckedBy.Items.Clear()
            $ComboboxCreatedBy.Items.Clear()
            Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                $ComboboxCheckedBy.Items.Add($ItemInList)
                $ComboboxCreatedBy.Items.Add($ItemInList)
            }
        }
        if ($ListType -eq "Projects") {
            $EmailSubjectAccessPathInput.Text = ""
            $EmailSubjectComboboxProjectName.Items.Clear()
            Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                $EmailSubjectComboboxProjectName.Items.Add($ItemInList)
            }
            UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
        }
        if ($ListType -eq "RegisterProjects") {
            $UpdateRegisterFormComboboxProjectName.Items.Clear()
            Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                $UpdateRegisterFormComboboxProjectName.Items.Add($ItemInList)
            }
        }
        if ($ListType -eq "Developers") {
            $UpdateRegisterFormComboboxDeveloperName.Items.Clear()
            Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                $UpdateRegisterFormComboboxDeveloperName.Items.Add($ItemInList)
            }  
        }  
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonDeleteItem)
    #Надпись для кнопки удалить
    $GetPropertyLabelButtonDelete = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonDelete.Location =  New-Object System.Drawing.Point(350,107) #x,y
    $GetPropertyLabelButtonDelete.Size =  New-Object System.Drawing.Point(180,15) #width,height
    if ($ListType -eq "Departments") {$GetPropertyLabelButtonDelete.Text = "Удалить отдел из списка"}
    if ($ListType -eq "Employees") {$GetPropertyLabelButtonDelete.Text = "Удалить сотрудника из списка"}
    if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {$GetPropertyLabelButtonDelete.Text = "Удалить проект из списка"}
    if ($ListType -eq "Developers") {$GetPropertyLabelButtonDelete.Text = "Удалить разработчика из списка"}
    $ManageCustomListsForm.Controls.Add($GetPropertyLabelButtonDelete)
    #Кнопка очистить
    $GetPropertyButtonClear = New-Object System.Windows.Forms.Button
    $GetPropertyButtonClear.Location = New-Object System.Drawing.Point(235,132) #x,y
    $GetPropertyButtonClear.Size = New-Object System.Drawing.Point(110,22) #width,height
    $GetPropertyButtonClear.Text = "Очистить список"
    $GetPropertyButtonClear.Add_Click({
        $ClickResult = Show-MessageBox -Title "Подтвердите действие" -Type YesNo -Message "Вы уверены, что хотите удалить все записи из списка?"
            if ($ClickResult -eq "Yes") {
            if ($ListType -eq "Departments") {Remove-AllItemsFromExistingList -ListType Departments}
            if ($ListType -eq "Employees") {Remove-AllItemsFromExistingList -ListType Employees}
            if ($ListType -eq "Projects" -or $ListType -eq "RegisterProjects") {Remove-AllItemsFromExistingList -ListType Projects}
            if ($ListType -eq "Developers") {Remove-AllItemsFromExistingList -ListType Developers}
            $GetPropertyListBoxBlackList.Items.Clear()
            if ($ListType -eq "Departments") {
                $ComboboxDepartmentName.Items.Clear()
                Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                    $ComboboxDepartmentName.Items.Add($ItemInList)
                }
            }
            if ($ListType -eq "Employees") {
                $ComboboxCheckedBy.Items.Clear()
                $ComboboxCreatedBy.Items.Clear()
                Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                    $ComboboxCheckedBy.Items.Add($ItemInList)
                    $ComboboxCreatedBy.Items.Add($ItemInList)
                }
            }
            if ($ListType -eq "Projects") {
                $EmailSubjectAccessPathInput.Text = ""
                $EmailSubjectComboboxProjectName.Items.Clear()
                Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                    $EmailSubjectComboboxProjectName.Items.Add($ItemInList)
                }
                UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
            }
            if ($ListType -eq "RegisterProjects") {
                $UpdateRegisterFormComboboxProjectName.Items.Clear()
                Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                    $UpdateRegisterFormComboboxProjectName.Items.Add($ItemInList)
                }
            }
            if ($ListType -eq "Developers") {
                $UpdateRegisterFormComboboxDeveloperName.Items.Clear()
                Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
                    $UpdateRegisterFormComboboxDeveloperName.Items.Add($ItemInList)
                }  
            }
        }
    })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonClear)
    #Надпись для кнопки очистить
    $GetPropertyLabelButtonClear = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonClear.Location =  New-Object System.Drawing.Point(350,135) #x,y
    $GetPropertyLabelButtonClear.Size =  New-Object System.Drawing.Point(180,15) #width,height
    $GetPropertyLabelButtonClear.Text = "Удалить все записи из списка"
    $ManageCustomListsForm.Controls.Add($GetPropertyLabelButtonClear)
    #Кнопка Сохранить
    $ManageCustomListsSaveButton = New-Object System.Windows.Forms.Button
    $ManageCustomListsSaveButton.Location = New-Object System.Drawing.Point(235,255) #x,y
    $ManageCustomListsSaveButton.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ManageCustomListsSaveButton.Text = "Закрыть"
    $ManageCustomListsSaveButton.Add_Click({
    $ManageCustomListsForm.Close()
    })
    $ManageCustomListsForm.Controls.Add($ManageCustomListsSaveButton)
    <#Кнопка Закрыть
    $ManageCustomListsCloseButton = New-Object System.Windows.Forms.Button
    $ManageCustomListsCloseButton.Location = New-Object System.Drawing.Point(350,255) #x,y
    $ManageCustomListsCloseButton.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ManageCustomListsCloseButton.Text = "Закрыть"
    $ManageCustomListsCloseButton.Add_Click({$ManageCustomListsForm.Close()})
    $ManageCustomListsForm.Controls.Add($ManageCustomListsCloseButton)#>
    $ManageCustomListsForm.ShowDialog()
}

Function Apply-FormattingInListTable($TableObject, $WordApp)
{
    #Сделать границы таблицы видимыми
    $TableObject.Borders.Enable = $true
    #Ширина таблицы (19,4 см)
    $TableObject.Columns.Item(1).Width = $WordApp.CentimetersToPoints(19.4)
    #Отсутуп от левого края в ячейках
    $TableObject.LeftPadding = $WordApp.CentimetersToPoints(0.05)
    #Отсутуп от правого края в ячейках
    $TableObject.RightPadding = $WordApp.CentimetersToPoints(0.05)
    #Установить вертикальное выравнивание по центру для всех ячеек
    $TableObject.Cell(1, 1).VerticalAlignment = 1
    #Настройка шрифта в таблице
    $TableObject.Cell(1, 1).Range.Font.Name = "Arial"
    $TableObject.Cell(1, 1).Range.Font.Size = 9
    #Интервал после (0 пт) для каждой ячейки таблицы
    $TableObject.Range.ParagraphFormat.SpaceAfter = 2
    $TableObject.Range.ParagraphFormat.SpaceBefore = 2
    $TableObject.Range.ParagraphFormat.LeftIndent = $WordApp.CentimetersToPoints(0.05)
    $TableObject.Range.ParagraphFormat.RightIndent = $WordApp.CentimetersToPoints(0.05)
    #Междустрочный интервал (одинарный) для каждой ячейки таблицы
    $TableObject.Range.ParagraphFormat.LineSpacingRule = 0
    #Установить высоту на минимум
    $TableObject.Cell(1, 1).HeightRule = 1
    $TableObject.Rows.Height = $word.CentimetersToPoints(0.5)
    #Разбить таблицу на 4 столбца и установить их ширину
    $TableObject.Cell(1, 1).Split(1, 4)
    $TableObject.Cell(1, 1).Column.Width = $WordApp.CentimetersToPoints(1.7)
    $TableObject.Cell(1, 2).Column.Width = $WordApp.CentimetersToPoints(1.7)
    $TableObject.Cell(1, 3).Column.Width = $WordApp.CentimetersToPoints(8)
    $TableObject.Cell(1, 4).Column.Width = $WordApp.CentimetersToPoints(8)
    #Добавить надписи в таблицу
    $TableObject.Cell(1, 1).Range.Text = "Поз."
    $TableObject.Cell(1, 2).Range.Text = "Изм."
    $TableObject.Cell(1, 3).Range.Text = "Обозначение"
    $TableObject.Cell(1, 4).Range.Text = "Примечание"
    #Установить выравниевание по центру для заголовка
    $TableObject.Rows.Item(1).Range.ParagraphFormat.Alignment = 1
    #Добавить строку для входных данных
    $TableObject.Rows.Add()
    #Запретить переносить ячейку на селдующую страницу
    $TableObject.Rows.Item(2).AllowBreakAcrossPages = $false
    #Отформатировать строку для входных данных
    $TableObject.Cell(2, 3).Range.ParagraphFormat.Alignment = 0
}

<#Function Generate-UpdateNotification ($NotificationName)
{
Kill -Name WINWORD -ErrorAction SilentlyContinue
Write-Host "Генерация шаблона извещения..."
Start-Sleep -Seconds 3
#Создать экземпляр приложения MS Word
$word = New-Object -ComObject Word.Application
#Создать документ MS Word
$document = $word.Documents.Add()
#Сделать вызванное приложение невидемым
$word.Visible = $false
Start-Sleep -Seconds 15
#НАСТРОЙКА ПОЛЕЙ ДОКУМЕНТА
#Левое поле (сантиметры)
$document.PageSetup.LeftMargin = $word.CentimetersToPoints(1)
#Правое поле (сантиметры)
$document.PageSetup.RightMargin = $word.CentimetersToPoints(1)
#Верхнее поле (сантиметры)
$document.PageSetup.TopMargin = $word.CentimetersToPoints(0.5)
#Нижнее поле (сантиметры)
$document.PageSetup.BottomMargin = $word.CentimetersToPoints(0.5)
#Верхний колонтитул
$document.PageSetup.HeaderDistance = $word.CentimetersToPoints(1)
#Нижний колонтитул
$document.PageSetup.FooterDistance = $word.CentimetersToPoints(1)

#НАСТРОЙКА ТАБЛИЦЫ
#Добавить таблицу
$document.Tables.Add($word.Selection.Range, 1, 1)
$table = $document.Tables.Item(1)
#Сделать границы таблицы видимыми
$table.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$table.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$table.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$table.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$table.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$table.Cell(1, 1).Range.Font.Name = "Arial"
$table.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$table.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$table.Range.ParagraphFormat.LineSpacingRule = 0

#ВЕРСТКА ТАБЛИЦЫ
#Разбить первую строку на 4 колонки
$table.Cell(1, 1).Split(1, 4)
$table.Cell(1, 1).Column.Width = $word.CentimetersToPoints(3)
$table.Cell(1, 2).Column.Width = $word.CentimetersToPoints(3)
$table.Cell(1, 3).Column.Width = $word.CentimetersToPoints(6.7)
$table.Cell(1, 4).Column.Width = $word.CentimetersToPoints(6.7)
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(1, 1).Merge($table.Cell(2, 1))
$table.Cell(1, 2).Merge($table.Cell(2, 2))

#Добавить строку с датами и информации о количестве страниц
$table.Rows.Add()
$table.Cell(3, 4).Split(1, 3)
$table.Cell(3, 4).SetWidth($word.CentimetersToPoints(2.7), 1)
$document.Range([ref]$table.Cell(3, 1).Range.Start, [ref]$table.Cell(3, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsBelow()
$document.Application.Selection.InsertRowsBelow()
$document.Application.Selection.InsertRowsBelow()
$Document.Application.Selection.Move()
$table.Cell(3, 2).Merge($table.Cell(3, 4))
$table.Cell(4, 2).Merge($table.Cell(4, 4))
$table.Cell(5, 3).Merge($table.Cell(5, 4))
$table.Cell(6, 3).Merge($table.Cell(6, 4))
$table.Cell(5, 1).Merge($table.Cell(5, 2))
$table.Cell(6, 1).Merge($table.Cell(6, 2))
$table.Cell(5, 3).Merge($table.Cell(5, 4))
$table.Cell(6, 3).Merge($table.Cell(6, 4))
$table.Cell(5, 1).Merge($table.Cell(6, 1))
$table.Cell(5, 2).Merge($table.Cell(6, 2))
$table.Cell(7, 1).Merge($table.Cell(7, 2))
$table.Cell(7, 2).Merge($table.Cell(7, 3))
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(12, 1).SetWidth($word.CentimetersToPoints(1.5), 1)
$table.Rows.Add()
$table.Rows.Add()
$table.Cell(12, 2).Merge($table.Cell(13, 2))
$table.Cell(14, 1).Merge($table.Cell(14, 2))

for ($i = 0; $i -lt 29; $i++) {
$table.Rows.Add()
}

#Вставить текст
$table.Cell(1, 3).Range.Text = "Извещение"
$table.Cell(1, 4).Range.Text = "Обозначение изменяемого документа"
$table.Cell(3, 1).Range.Text = "Дата выпуска"
$table.Cell(3, 2).Range.Text = "Срок внесения изменений"
$table.Cell(3, 3).Range.Text = "Лист"
$table.Cell(3, 4).Range.Text = "Листов"
$table.Cell(5, 1).Range.Text = "Причина"
$table.Cell(5, 3).Range.Text = "Код"
$table.Cell(7, 1).Range.Text = "Указание о заделе"
$table.Cell(8, 1).Range.Text = "Указание о внедрении"
$table.Cell(9, 1).Range.Text = "Применяемость"
$table.Cell(10, 1).Range.Text = "Разослать"
$table.Cell(11, 1).Range.Text = "Приложение"
$table.Cell(12, 1).Range.Text = "Изм."
$table.Cell(13, 1).Range.Text = "-"
$table.Cell(12, 2).Range.Text = "Содержание изменения"
#Добавить логотип компании и выровнять по центру
$table.Cell(1, 1).Range.InlineShapes.AddPicture("$PSScriptRoot\logo.jpg", $false, $true)
$table.Cell(1, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(1, 1).Range.ParagraphFormat.SpaceBefore = 1

#ФОРМАТИРОВАНИЕ ТАБЛИЦЫ ПОСЛЕ ВЕРСТКИ
#Высота строки (0,5 см) для всех ячеек
$table.Rows.Height = $word.CentimetersToPoints(0.5)
#Включить постоянную высоту для строк таблицы
$table.Rows.HeightRule = 2
$table.Cell(2, 3).Height = $word.CentimetersToPoints(1)
$table.Cell(7, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(8, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(9, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(10, 1).Height = $word.CentimetersToPoints(1)
$table.Cell(11, 1).Height = $word.CentimetersToPoints(1)
#Настроить выравнивание в ячейках
$table.Cell(1, 2).Range.ParagraphFormat.Alignment = 1
$table.Cell(5, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(7, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(8, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(9, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(10, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(11, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(12, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(3, 3).Range.ParagraphFormat.Alignment = 1
$table.Cell(3, 4).Range.ParagraphFormat.Alignment = 1
$table.Cell(5, 3).Range.ParagraphFormat.Alignment = 1
$table.Cell(13, 1).Range.ParagraphFormat.Alignment = 1
$table.Cell(12, 2).Range.ParagraphFormat.Alignment = 1
#Вставить надписть ГОСТ
$document.Range([ref]$table.Cell(1, 1).Range.Start, [ref]$table.Cell(1, 4).Range.Start).Select()
$document.Application.Selection.InsertRowsAbove()
$table.Cell(1, 1).Merge($table.Cell(1, 4))
$table.Cell(1, 1).Range.ParagraphFormat.Alignment = 2
$table.Cell(1, 1).Range.Text = "ГОСТ 2.503-90 Форма 1"
$table.Cell(1, 1).Borders.Item(-2).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-1).LineStyle = 0
$table.Cell(1, 1).Borders.Item(-4).LineStyle = 0

#ДОБАВИТЬ НАДПИСЬ В ВЕРХНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
$document.PageSetup.DifferentFirstPageHeaderFooter = -1
$document.Sections.Item(1).Headers.Item(2).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Конфиденциально"
$shapeTop = $document.Sections.Item(1).Headers.Item(2).Shapes.Item(1)
try {$shapeTop.Height = $word.CentimetersToPoints(0.8)} catch {Out-Null}
Start-Sleep -Seconds 3
$shapeTop.Height = $word.CentimetersToPoints(0.8)
$shapeTop.Width = $word.CentimetersToPoints(8.5)
$shapeTop.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeTop.TextFrame.TextRange.Font.Size = 16
$shapeTop.TextFrame.TextRange.Font.Name = "Arial"
$shapeTop.TextFrame.TextRange.Font.Bold = $true
$shapeTop.TextFrame.TextRange.Font.ColorIndex = 1
$shapeTop.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeTop.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeTop.RelativeHorizontalPosition = 1
$shapeTop.Left = $word.CentimetersToPoints(11.8)
$shapeTop.RelativeVerticalPosition = 1
$shapeTop.Top = $word.CentimetersToPoints(0.4)
$shapeTop.Fill.Visible = 0
$shapeTop.Line.Weight = 1
$shapeTop.Line.Visible = 0
$shapeTop.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeTop.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeTop.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#НИЖНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
$footer = $document.Sections.Item(1).Footers.Item(2)
$footer.Range.Tables.Add($footer.Range, 1, 1)
$footerTable = $footer.Range.Tables.Item(1)
$footerTable.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$footerTable.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$footerTable.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$footerTable.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$footerTable.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$footerTable.Cell(1, 1).Range.Font.Name = "Arial"
$footerTable.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$footerTable.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$footerTable.Range.ParagraphFormat.LineSpacingRule = 0
#Разбить таблицу на восемь колонок и задать их ширину
$footerTable.Cell(1, 1).Split(1, 8)
$footerTable.Cell(1, 1).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 2).Column.Width = $word.CentimetersToPoints(3.5)
$footerTable.Cell(1, 3).Column.Width = $word.CentimetersToPoints(2.2)
$footerTable.Cell(1, 4).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 5).Column.Width = $word.CentimetersToPoints(2)
$footerTable.Cell(1, 6).Column.Width = $word.CentimetersToPoints(3.5)
$footerTable.Cell(1, 7).Column.Width = $word.CentimetersToPoints(2.2)
$footerTable.Cell(1, 8).Column.Width = $word.CentimetersToPoints(2)
#Добавить еще четырке строки в таблицу
$footerTable.Rows.Add()
$footerTable.Rows.Add()
$footerTable.Rows.Add()
$footerTable.Rows.Add()
#Добавить надписи в таблицу
$footerTable.Cell(1, 2).Range.Text = "Фамилия"
$footerTable.Cell(1, 3).Range.Text = "Подпись"
$footerTable.Cell(1, 4).Range.Text = "Дата"
$footerTable.Cell(1, 6).Range.Text = "Фамилия"
$footerTable.Cell(1, 7).Range.Text = "Подпись"
$footerTable.Cell(1, 8).Range.Text = "Дата"
$footerTable.Cell(2, 1).Range.Text = "Составил"
$footerTable.Cell(3, 1).Range.Text = "Проверил"
$footerTable.Cell(4, 1).Range.Text = "Т. контр."
$footerTable.Cell(5, 1).Range.Text = "Изменение внес"
$footerTable.Cell(2, 5).Range.Text = "Н. контр."
$footerTable.Cell(4, 5).Range.Text = "Утвердил"
$footerTable.Rows.Item(1).Range.ParagraphFormat.Alignment = 1
#Объединить строку "Изменения внес"
$footerTable.Cell(5, 1).Merge($footerTable.Cell(5, 8))
#Высота строки (0,5 см) для всех ячеек
$footerTable.Rows.Height = $word.CentimetersToPoints(0.5)
#Включить постоянную высоту для строк таблицы
$footerTable.Rows.HeightRule = 2

#ДОБАВИТЬ НАДПИСЬ В НИЖНИЙ КОЛОНТИТУЛ НА ПЕРВОЙ СТРАНИЦЕ
$document.Sections.Item(1).Headers.Item(2).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Коммерческая тайна"
$shapeBot = $document.Sections.Item(1).Headers.Item(2).Shapes.Item(2)
$shapeBot.Height = $word.CentimetersToPoints(0.8)
$shapeBot.Width = $word.CentimetersToPoints(8.5)
$shapeBot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeBot.TextFrame.TextRange.Font.Size = 16
$shapeBot.TextFrame.TextRange.Font.Name = "Arial"
$shapeBot.TextFrame.TextRange.Font.Bold = $true
$shapeBot.TextFrame.TextRange.Font.ColorIndex = 1
$shapeBot.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeBot.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeBot.RelativeHorizontalPosition = 1
$shapeBot.Left = $word.CentimetersToPoints(11.8)
$shapeBot.RelativeVerticalPosition = 1
$shapeBot.Top = $word.CentimetersToPoints(28.4)
$shapeBot.Fill.Visible = 0
$shapeBot.Line.Weight = 1
$shapeBot.Line.Visible = 0
$shapeBot.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeBot.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeBot.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeBot.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#ВТОРАЯ СТРАНИЦА
#Верхний колонтитул

#Надпись в верхнем колонтитуле
$document.Sections.Item(1).Headers.Item(1).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Конфиденциально"
$shapeTopPageTwo = $document.Sections.Item(1).Headers.Item(1).Shapes.Item(3)
$shapeTopPageTwo.Height = $word.CentimetersToPoints(0.8)
$shapeTopPageTwo.Width = $word.CentimetersToPoints(8.5)
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeTopPageTwo.TextFrame.TextRange.Font.Size = 16
$shapeTopPageTwo.TextFrame.TextRange.Font.Name = "Arial"
$shapeTopPageTwo.TextFrame.TextRange.Font.Bold = $true
$shapeTopPageTwo.TextFrame.TextRange.Font.ColorIndex = 1
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeTopPageTwo.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeTopPageTwo.RelativeHorizontalPosition = 1
$shapeTopPageTwo.RelativeVerticalPosition = 1
$shapeTopPageTwo.Left = $word.CentimetersToPoints(11.8)
$shapeTopPageTwo.Top = $word.CentimetersToPoints(0.4)
$shapeTopPageTwo.Fill.Visible = 0
$shapeTopPageTwo.Line.Weight = 1
$shapeTopPageTwo.Line.Visible = 0
$shapeTopPageTwo.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeTopPageTwo.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeTopPageTwo.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeTopPageTwo.TextFrame.MarginTop = $word.CentimetersToPoints(0)

#Таблица в верхнем колонтитуле
$header = $document.Sections.Item(1).Headers.Item(1).Range
$header.Collapse(0)
$document.Sections.Item(1).Headers.Item(1).Range.Tables.Add($header, 1, 1)
$headerTable = $document.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1)
$headerTable.Borders.Enable = $true
#Ширина таблицы (19,4 см)
$headerTable.Columns.Item(1).Width = $word.CentimetersToPoints(19.4)
#Отсутуп от левого края в ячейках
$headerTable.LeftPadding = $word.CentimetersToPoints(0.05)
#Отсутуп от правого края в ячейках
$headerTable.RightPadding = $word.CentimetersToPoints(0.05)
#Установить вертикальное выравнивание по центру для всех ячеек
$headerTable.Cell(1, 1).VerticalAlignment = 1
#Настройка шрифта в таблице
$headerTable.Cell(1, 1).Range.Font.Name = "Arial"
$headerTable.Cell(1, 1).Range.Font.Size = 9
#Интервал после (0 пт) для каждой ячейки таблицы
$headerTable.Range.ParagraphFormat.SpaceAfter = 0
#Междустрочный интервал (одинарный) для каждой ячейки таблицы
$headerTable.Range.ParagraphFormat.LineSpacingRule = 0
#Вставить надписть ГОСТ
$headerTable.Cell(1, 1).Range.Text = "ГОСТ 2.503-90 Форма 1"
$headerTable.Rows.Add()
$headerTable.Cell(1, 1).Range.ParagraphFormat.Alignment = 2
$headerTable.Cell(1, 1).Borders.Item(-2).LineStyle = 0
$headerTable.Cell(1, 1).Borders.Item(-1).LineStyle = 0
$headerTable.Cell(1, 1).Borders.Item(-4).LineStyle = 0
#Собрать сотальную часть
$headerTable.Cell(2, 1).Split(1, 4)
$headerTable.Cell(2, 1).SetWidth($word.CentimetersToPoints(2.5), 1)
$headerTable.Cell(2, 2).SetWidth($word.CentimetersToPoints(13.9), 1)
$headerTable.Rows.Add()
$headerTable.Cell(4, 2).Merge($headerTable.Cell(4, 4))
$headerTable.Cell(4, 1).SetWidth($word.CentimetersToPoints(1.5), 1)
$headerTable.Rows.Add()
$headerTable.Cell(3, 2).Merge($headerTable.Cell(4, 2))
#Высота строки (0,5 см) для всех ячеек
$headerTable.Rows.Height = $word.CentimetersToPoints(0.5)
#Настройка выравнивания в ячейках
$headerTable.Cell(2, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(2, 4).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(3, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(4, 1).Range.ParagraphFormat.Alignment = 1
$headerTable.Cell(3, 2).Range.ParagraphFormat.Alignment = 1
#Включить постоянную высоту для строк таблицы
$headerTable.Rows.HeightRule = 2
#Добавить текст в таблице верхнего колонтитула второй страницы
$headerTable.Cell(2, 1).Range.Text = "Извещение"
$headerTable.Cell(2, 3).Range.Text = "Лист"
$headerTable.Cell(3, 1).Range.Text = "Изм."
$headerTable.Cell(4, 1).Range.Text = "-"
$headerTable.Cell(3, 2).Range.Text = "Содержание изменения"
#Надпись в нижнем колонтитуле
$document.Sections.Item(1).Headers.Item(1).Shapes.AddShape(1, 10, 10, 200, 20).TextFrame.TextRange.Text = "Коммерческая тайна"
$shapeBotPageTwo = $document.Sections.Item(1).Headers.Item(1).Shapes.Item(4)
$shapeBotPageTwo.Height = $word.CentimetersToPoints(0.8)
$shapeBotPageTwo.Width = $word.CentimetersToPoints(8.5)
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.Alignment = 2
$shapeBotPageTwo.TextFrame.TextRange.Font.Size = 16
$shapeBotPageTwo.TextFrame.TextRange.Font.Name = "Arial"
$shapeBotPageTwo.TextFrame.TextRange.Font.Bold = $true
$shapeBotPageTwo.TextFrame.TextRange.Font.ColorIndex = 1
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
$shapeBotPageTwo.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = 0
$shapeBotPageTwo.RelativeHorizontalPosition = 1
$shapeBotPageTwo.Left = $word.CentimetersToPoints(11.8)
$shapeBotPageTwo.RelativeVerticalPosition = 1
$shapeBotPageTwo.Top = $word.CentimetersToPoints(28.4)
$shapeBotPageTwo.Fill.Visible = 0
$shapeBotPageTwo.Line.Weight = 1
$shapeBotPageTwo.Line.Visible = 0
$shapeBotPageTwo.TextFrame.MarginBottom = $word.CentimetersToPoints(0)
$shapeBotPageTwo.TextFrame.MarginLeft = $word.CentimetersToPoints(0.1)
$shapeBotPageTwo.TextFrame.MarginRight = $word.CentimetersToPoints(0.1)
$shapeBotPageTwo.TextFrame.MarginTop = $word.CentimetersToPoints(0)

for ($i = 0; $i -lt 30; $i++) {
$table.Cell(15, 1).Delete()
}

$document.Paragraphs.Item(52).Range.Font.Name = "Arial"
$document.Paragraphs.Item(52).Range.Font.Size = 9
$document.Paragraphs.Item(52).Range.ParagraphFormat.SpaceAfter = 0
$document.Paragraphs.Item(52).Range.ParagraphFormat.LineSpacingRule = 0
$document.Paragraphs.Item(52).Range = [char]10 + [char]10 + "Заменить:" + [char]10 + [char]10 + [char]10 + [char]10 + "Аннулировать:" + [char]10 + [char]10 + [char]10 + [char]10 + "Выпустить:" + [char]10
#Сделать надписи к спискам жирными
$document.Paragraphs.Item(54).Range.Font.Bold = $true
$document.Paragraphs.Item(58).Range.Font.Bold = $true
$document.Paragraphs.Item(62).Range.Font.Bold = $true
#Вставить таблицы
$document.Paragraphs.Item(55).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Paragraphs.Item(60).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Paragraphs.Item(65).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
#Отформатировать таблицы
Apply-FormattingInListTable -TableObject $document.Tables.Item(2) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(3) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(4) -WordApp $word
Start-Sleep -Seconds 2
$document.SaveAs([ref]"$PSScriptRoot\$NotificationName.docx")
Start-Sleep -Seconds 2
$document.Close()
Start-Sleep -Seconds 2
$word.Quit()
Start-Sleep -Seconds 3
Kill -Name WINWORD -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
}#>

Function Add-Data ($TableObject, $List)
{
    $RowCounter = 2
    Foreach ($Item in $List.Items) {
        $TableObject.Rows.Add()
        $TableObject.Cell($RowCounter, 1).Range.Text = $script:CounterForWordItems
        $script:CounterForWordItems += 1
        if ($Item.Subitems[2].Text -eq "Программа") {
            $TableObject.Cell($RowCounter, 2).Range.Text = "-"
            $TableObject.Cell($RowCounter, 3).Range.Text = "$($Item.Text)"
            $TableObject.Cell($RowCounter, 4).Range.Text = "$($Item.Subitems[1].Text)"
        } else {
            $TableObject.Cell($RowCounter, 2).Range.Text = "$($Item.Subitems[1].Text)"
            $TableObject.Cell($RowCounter, 3).Range.Text = "$($Item.Text)"
        }
        $RowCounter += 1
    }
}

Function Move-ListsToWordDocument ($NotificationName)
{
Kill -Name WINWORD -ErrorAction SilentlyContinue
Write-Host "Заполняю шаблон..."
Start-Sleep -Seconds 3
$script:counter = 1
#Создать экземпляр приложения MS Word
$WordToPopulate = New-Object -ComObject Word.Application
#Создать документ MS Word
$DocumentToPopulate = $WordToPopulate.Documents.Open("$PSScriptRoot\Template.docx")
#Сделать вызванное приложение невидемым
$WordToPopulate.Visible = $false
Start-Sleep -Seconds 5
#Заполнить таблицы
Write-Host "Заполняю таблицу Заменить..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(2) -List $ListViewReplace
Write-Host "Заполняю таблицу Аннулировать..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(3) -List $ListViewRemove
Write-Host "Заполняю таблицу Выпустить..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(4) -List $ListViewAdd
Write-Host "Обновляю поля в Word документе..."
#Обращаемся к элементам таблицы
$HeaderTablePopulate = $DocumentToPopulate.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1)
$FooterFirstPageTablePopulate = $DocumentToPopulate.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1)
$TablePopulate = $DocumentToPopulate.Tables.Item(1)

#Вставить название отдела
$TablePopulate.Cell(2, 2).Range.Text = "$($ComboboxDepartmentName.SelectedItem)"
$TablePopulate.Cell(2, 2).Range.Font.Bold = $true

#Вставить номер извещения
$TablePopulate.Cell(3, 3).Range.Text = "$($UpdateNotificationNumberInput.Text)"
$TablePopulate.Cell(3, 3).Range.Font.Bold = $true
$HeaderTablePopulate.Cell(2, 2).Range.Text = "$($UpdateNotificationNumberInput.Text)"
$HeaderTablePopulate.Cell(2, 2).Range.Font.Bold = $true

#Вставить "См. содержание изменения"
$TablePopulate.Cell(3, 4).Range.Text = "См. содержание изменения"
$TablePopulate.Cell(3, 4).Range.Font.Bold = $true

#Вставить дату выпуска
$TablePopulate.Cell(5, 1).Range.Text = "$($CalendarIssueDateInput.Text)"
$TablePopulate.Cell(5, 1).Range.Font.Bold = $true
$FooterFirstPageTablePopulate.Cell(2, 4).Range.Text = "$($CalendarIssueDateInput.Text)"
$FooterFirstPageTablePopulate.Cell(3, 4).Range.Text = "$($CalendarIssueDateInput.Text)"

#Вставить дату срока внесения изменений
$TablePopulate.Cell(5, 2).Range.Text = "$($CalendarApplyUpdatesUntilInput.Text)"
$TablePopulate.Cell(5, 2).Range.Font.Bold = $true

#Вставить код
$TablePopulate.Cell(7, 3).Range.Text = "$($ComboboxCodes.SelectedItem)"
$TablePopulate.Cell(7, 3).Range.Font.Bold = $true
$TablePopulate.Cell(7, 3).Range.ParagraphFormat.Alignment = 1

#Вставить причину
$TablePopulate.Cell(6, 2).Range.Text = "$script:GlobalReasonField"
$TablePopulate.Cell(6, 2).Range.Font.Bold = $true

#Вставить указание о заделе
$TablePopulate.Cell(8, 2).Range.Text = "$script:GlobalInStoreField"
$TablePopulate.Cell(8, 2).Range.Font.Bold = $true

#Вставить указание о внедрении
$TablePopulate.Cell(9, 2).Range.Text = "$script:GlobalStartUsageField"
$TablePopulate.Cell(9, 2).Range.Font.Bold = $true

#Вставить применяемость
$TablePopulate.Cell(10, 2).Range.Text = "$script:GlobalApplicableToField"
$TablePopulate.Cell(10, 2).Range.Font.Bold = $true

#Вставить разослать
$TablePopulate.Cell(11, 2).Range.Text = "$script:GlobalSendToField"
$TablePopulate.Cell(11, 2).Range.Font.Bold = $true

#Вставить приложение
$TablePopulate.Cell(12, 2).Range.Text = "$script:GlobalAppendixField"
$TablePopulate.Cell(12, 2).Range.Font.Bold = $true

#Вставить поля
$TablePopulate.Cell(5, 4).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 26)
$TablePopulate.Cell(5, 4).Range.ParagraphFormat.Alignment = 1

#Вставить ФИО
$FooterFirstPageTablePopulate.Cell(2, 2).Range.Text = "$($ComboboxCreatedBy.SelectedItem)"
$FooterFirstPageTablePopulate.Cell(3, 2).Range.Text = "$($ComboboxCheckedBy.SelectedItem)"

$HeaderTablePopulate.Cell(2, 4).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 33)

$TablePopulate.Cell(5, 3).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 33)
$TablePopulate.Cell(5, 3).Range.ParagraphFormat.Alignment = 1
Start-Sleep -Seconds 2
#Обновить поля
$DocumentToPopulate.Fields.Update()
$Wholestory = $DocumentToPopulate.Range()
$TotalPages = $Wholestory.Information(4)
$TablePopulate.Cell(5, 4).Range.Text = $TotalPages
$DocumentToPopulate.SaveAs([ref]"$PSScriptRoot\$NotificationName.docx")
Start-Sleep -Seconds 3
$DocumentToPopulate.Close()
Start-Sleep -Seconds 3
$WordToPopulate.Quit()
Start-Sleep -Seconds 3
Kill -Name WINWORD -ErrorAction SilentlyContinue
Write-Host "Извещение сгенерировано. Скрипт закончил работу."
}

Function Create-HtmlReportForErrors ([array]$Errors) 
{
Add-Content "$PSScriptRoot\Ошибки.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Ошибки</title>
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
<h4>Следующие ошибки должны быть исправлены перед началом процедуры внесения изменений по ИИ:</h4>
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
        <td class=""Item_Number"">$($ErrorCounter)</td>
        <td class=""Details"">$($_)</td>
    </tr>" -Encoding UTF8
    $ErrorCounter += 1
}

Add-Content "$PSScriptRoot\Ошибки.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
}

Function Create-HtmlReportForUpdateResults ([array]$Errors)
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
    width: 3%;
    text-align: left;
}
.File {
    width:27%;
    text-align: left;
}
.Action {
    width:15%;
    text-align: left;
}
.Details {
    width:55%;
    text-align: left;
}
</style>
</head>
<body>
<div>
<h4>Отчет о действиях, выполненных во время внесения изменений по ИИ:</h4>
<div>
</div>
<table style=""width:100%"">
    <tr>
        <th class=""Item_Number"">№</th>
        <th class=""File"">Файл</th>
        <th class=""Action"">Операция</th>
        <th class=""Details"">Действия/Результат</th>
    </tr>" -Encoding UTF8

$ErrorCounter = 1
for ($i = 0; $i -lt $Errors[0].Length; $i++) {
    Add-Content "$PSScriptRoot\Отчет.html" "    <tr>
        <td class=""Item_Number"">$($ErrorCounter)</td>
        <td class=""File"">$($Errors[0][$i])</td>
        <td class=""Action"">$($Errors[1][$i])</td>
        <td class=""Details"">$($Errors[2][$i])</td>
    </tr>" -Encoding UTF8
    $ErrorCounter += 1
}

Add-Content "$PSScriptRoot\Отчет.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
}

Function Copy-FileToFolder ($FileFullName, $Destination, [ValidateSet("Backup", "Default", "Notification")]$Type)
{
    $ExceptionCatched = $false
    try {
        [System.IO.File]::Copy($FileFullName, "$Destination\$([System.IO.Path]::GetFileName($FileFullName))", $true)
    } catch [Exception] {
        $ExceptionCatched = $true
        $ExceptionText = $_.Exception.Message
    }
    if ($Type -eq "Backup") {if ($ExceptionCatched -eq $true) {$script:AggregatingString += "<br>Создать резервную копию текущей версии файла: <font color=""red""><b>Ошибка</b></font> ($ExceptionText)"} else {$script:AggregatingString += "<br>Создать резервную копию текущей версии файла: <font color=""green""><b>Успешно</b></font>"}}
    if ($Type -eq "Default") {if ($ExceptionCatched -eq $true) {$script:AggregatingString += "<br>Скопировать публикуемый файл в папку с текущей версией проекта: <font color=""red""><b>Ошибка</b></font> ($ExceptionText)"} else {$script:AggregatingString += "<br>Скопировать публикуемый файл в папку с текущей версией проекта: <font color=""green""><b>Успешно</b></font>"}}
    if ($Type -eq "Notification") {if ($ExceptionCatched -eq $true) {$script:AggregatingString += "<br>Скопировать ИИ в архивную папку: <font color=""red""><b>Ошибка</b></font> ($ExceptionText)"} else {$script:AggregatingString += "<br>Скопировать ИИ в архивную папку: <font color=""green""><b>Успешно</b></font>"}}
}

Function Move-FileToFolder ($FileFullName, $Destination)
{
    $ExceptionCatched = $false
    try {
        [System.IO.File]::Move($FileFullName, "$Destination\$([System.IO.Path]::GetFileName($FileFullName))")
    } catch [Exception] {
        $ExceptionCatched = $true
        $ExceptionText = $_.Exception.Message
    }
    if ($ExceptionCatched -eq $true) {$script:AggregatingString += "<br>Переместить текущую версию файла в архивную папку: <font color=""red""><b>Ошибка</b></font> ($ExceptionText)"} else {$script:AggregatingString += "<br>Переместить текущую версию файла в архивную папку: <font color=""green""><b>Успешно</b></font>"} 
}

Function Apply-Changes ($BackupFlag)
{
    [array]$ErrorsForHtmlReport = @(), @(), @()
    #Замена файлов
    Write-Host "Начал работу..."
    Write-Host "Вношу изменения из списка Заменить"
    $ListViewReplace.Items | % {
        #Делаем резервную копию файла и переносим его в архивную папку
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "Работаю с документом $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*" | % {
                #Резервное копирование
                if ($BackupFlag -eq $true) {Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToBackupFolder -Type Backup}
                #Перемещает файл из папки с текущим релизом в архивную папку
                Move-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToArchiveFolder
                Start-Sleep -Seconds 1
                #Копируем из папки с публикуемыми файлами в папку с текущей версией
                Copy-FileToFolder -FileFullName "$script:MakeChangesPathToFilesBeingPublished\$([System.IO.Path]::GetFileName($_))" -Destination $script:MakeChangesPathToCurrentVersion -Type Default
                #Собираем данные для HTML отчета
                $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
                $ErrorsForHtmlReport[1] += "Заменить"
                $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
                $script:AggregatingString = ""
            } 
        }
        if ($_.SubItems[2].Text -eq "Программа") {
        Write-Host "Работаю с программой $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)" | % {
                #Резервное копирование
                if ($BackupFlag -eq $true) {Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToBackupFolder -Type Backup}
                #Перемещает файл из папки с текущим релизом в архивную папку
                Move-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToArchiveFolder
                Start-Sleep -Seconds 1
                #Копируем из папки с публикуемыми файлами в папку с текущей версией
                Copy-FileToFolder -FileFullName "$script:MakeChangesPathToFilesBeingPublished\$([System.IO.Path]::GetFileName($_))" -Destination $script:MakeChangesPathToCurrentVersion -Type Default
                #Собираем данные для HTML отчета
                $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
                $ErrorsForHtmlReport[1] += "Заменить"
                $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
                $script:AggregatingString = ""
            }
        } 
    }
    Write-Host "Вношу изменения из списка Аннулировать"
    $ListViewRemove.Items | % {
        #Делаем резервную копию файла и переносим его в архиную папку
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "Работаю с документом $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*" | % {
                #Резервное копирование
                if ($BackupFlag -eq $true) {Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToBackupFolder -Type Backup}
                Start-Sleep -Seconds 1
                #Перемещает файл из папки с текущим релизом в архивную папку
                Move-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToArchiveFolder
                #Собираем данные для HTML отчета
                $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
                $ErrorsForHtmlReport[1] += "Аннулировать"
                $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
                $script:AggregatingString = ""
            }
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            Write-Host "Работаю с программой $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)" | % {
                #Резервное копирование
                if ($BackupFlag -eq $true) {Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToBackupFolder -Type Backup}
                Start-Sleep -Seconds 1
                #Перемещает файл из папки с текущим релизом в архивную папку
                Move-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToArchiveFolder
                #Собираем данные для HTML отчета
                $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
                $ErrorsForHtmlReport[1] += "Аннулировать"
                $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
                $script:AggregatingString = ""
            }
        }
    }
    Write-Host "Вношу изменения из списка Выпустить"
    $ListViewAdd.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            Write-Host "Работаю с документом $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text).*" | % {
            Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToCurrentVersion -Type Default
            #Собираем данные для HTML отчета
            $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
            $ErrorsForHtmlReport[1] += "Выпустить"
            $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
            $script:AggregatingString = ""
            }
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            Write-Host "Работаю с программой $($_.Text)"
            Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text)" | % {
            Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToCurrentVersion -Type Default
            #Собираем данные для HTML отчета
            $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
            $ErrorsForHtmlReport[1] += "Выпустить"
            $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
            $script:AggregatingString = ""
            }
        }
    }
    #Переносим ИИ во всех форматах в архивную папку
    Write-Host "Архивация ИИ..."
    Get-ChildItem -Path "$script:MakeChangesPathToFilesBeingPublished\$($UpdateNotificationNumberInput.Text).*" | % {
        Copy-FileToFolder -FileFullName $_ -Destination $script:MakeChangesPathToArchiveFolder -Type Notification
        #Собираем данные для HTML отчета
        $ErrorsForHtmlReport[0] += [System.IO.Path]::GetFileName($_)
        $ErrorsForHtmlReport[1] += "Архивация ИИ"
        $ErrorsForHtmlReport[2] += $script:AggregatingString.Trim("<br>")
        $script:AggregatingString = ""
    }
    if (Test-Path -Path "$PSScriptRoot\Отчет.html") {Remove-Item -Path "$PSScriptRoot\Отчет.html"}
    Start-Sleep -Seconds 3
    Create-HtmlReportForUpdateResults -Errors $ErrorsForHtmlReport
    Write-Host "Изменения внесены..."
}

Function Check-Conditions ($BackupFlag)
{
    $ErrorDetectedFlag = $false
    $ErrorsToBePublished = @()
    #ПРОВЕРКИ
    #Папка для резервной копии пуста
    if ($BackupFlag -eq $true) {
        if ((Get-ChildItem -Path $script:MakeChangesPathToBackupFolder).Count -gt 0) {$ErrorsToBePublished += "Папка для резервного копирования содержит файлы. Данная папка должна быть пустой перед началом процедуры внесения изменений по ИИ."; $ErrorDetectedFlag = $true}
    }
    #Архивная папка пуста
    if ((Get-ChildItem -Path $script:MakeChangesPathToArchiveFolder).Count -gt 0) {$ErrorsToBePublished += "Архивная папка содержит файлы. Данная папка должна быть пустой перед началом процедуры внесения изменений по ИИ."; $ErrorDetectedFlag = $true}
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
            if ((Test-Path -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text).*") -eq $false) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Test-Path -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text)") -eq $false) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }  
    }
    #Присутствуют ли файлы из списка в папке с текущей версией проекта?
    $ListViewAdd.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*") -eq $true) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но папка с текущей версией проекта уже содержит данный документ."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)") -eq $true) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Выпустить</b>, но папка с текущей версией проекта уже содержит данную программу."; $ErrorDetectedFlag = $true}
        }  
    }
    #Проверки для списка Заменить.
    #Все файлы из списка присутствуют в папке с публикуемыми файлами?
    $ListViewReplace.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Test-Path -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text).*") -eq $false) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Test-Path -Path "$script:MakeChangesPathToFilesBeingPublished\$($_.Text)") -eq $false) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с публикуемыми файлами."; $ErrorDetectedFlag = $true}
        }  
    }
    #Присутствуют ли файлы из списка в папке с текущей версией проекта?
    $ListViewReplace.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*") -eq $false) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)") -eq $false) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Заменить</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }  
    }
    #Проверки для списка Аннулировать.
    #Присутствуют ли файлы из списка в папке с текущей версией проекта?
    $ListViewRemove.Items | % {
        if ($_.SubItems[2].Text -eq "Документ") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text).*") -eq $false) {$ErrorsToBePublished += "Документ <i>$($_.Text)</i> находится в списке <b>Аннулировать</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }
        if ($_.SubItems[2].Text -eq "Программа") {
            if ((Test-Path -Path "$script:MakeChangesPathToCurrentVersion\$($_.Text)") -eq $false) {$ErrorsToBePublished += "Программа <i>$($_.Text)</i> находится в списке <b>Аннулировать</b>, но отсутствует в папке с текущей версией проекта."; $ErrorDetectedFlag = $true}
        }  
    }
    #ПРОВЕРКИ
    if ($ErrorDetectedFlag -eq $true) {
        if ((Test-Path -Path "$PSScriptRoot\Ошибки.html") -eq $true) {Remove-Item -Path "$PSScriptRoot\Ошибки.html" -Force}
        Create-HtmlReportForErrors -Errors $ErrorsToBePublished
    }
    if ($ErrorDetectedFlag -eq $false) {return $false} else {return $true}
}

Function Check-FileNameUniqueness ()
{
    Add-Content "$PSScriptRoot\Проверка уникальности обозначений.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Проверка комплектности</title>
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
.Designation {
    width:47%;
    text-align: left;
}
.Check_results {
    width:50%;
    text-align: left;
}
</style>
</head>
<body>
<div>
<h4>Проверка уникальности обозначений из списка Выпустить:</h4>
<div>
</div>
<table style=""width:100%"">
    <tr>
        <th class=""Item_number"">№</th>
        <th class=""Designation"">Обозначение</th>
        <th class=""Check_results"">Результат проверки</th>
    </tr>" -Encoding UTF8
    $Register = New-Object -ComObject Excel.Application
    $Register.Visible = $false
    $RegisterWorkbook = $Register.WorkBooks.Open($script:PathToRegister)
    $RegisterWorksheet = $RegisterWorkbook.Worksheets.Item(1)
    if ($RegisterWorksheet.AutoFilterMode -eq $true) {$RegisterWorksheet.ShowAllData()}
    $RegisterLastRow = $RegisterWorksheet.Cells.Item($RegisterWorksheet.Rows.Count, "E").End(-4162).Row
    $SearchRange = $RegisterWorksheet.Range("C2:E$($RegisterLastRow)")
    $RegisterRowCounter = 1
    $ListViewAdd.Items | % {
    if (($SearchRange.Find([string]$_.Text, [Type]::Missing, -4163, 1)) -ne $null) {
        Add-Content "$PSScriptRoot\Проверка уникальности обозначений.html" "    <tr>
<td class=""Item_Number"">$($RegisterRowCounter)</td>
<td class=""Designation"">$($_.Text)</td>
<td class=""Check_results""><font color=""red""><b>Не уникально</b></font></td>
    </tr>" -Encoding UTF8
        $RegisterRowCounter += 1
        } else {
        Add-Content "$PSScriptRoot\Проверка уникальности обозначений.html" "    <tr>
<td class=""Item_Number"">$($RegisterRowCounter)</td>
<td class=""Designation"">$($_.Text)</td>
<td class=""Check_results""><font color=""green""><b>Уникально</b></font></td>
    </tr>" -Encoding UTF8
        $RegisterRowCounter += 1
        }
    }
    $RegisterWorkbook.Close($false)
    $Register.Quit()
    Add-Content "$PSScriptRoot\Проверка уникальности обозначений.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
Invoke-Item "$PSScriptRoot\Проверка уникальности обозначений.html"
}

Function UniquenessCheckForm ()
{
    $UniquenessCheckForm = New-Object System.Windows.Forms.Form
    $UniquenessCheckForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $UniquenessCheckForm.ShowIcon = $false
    $UniquenessCheckForm.AutoSize = $true
    $UniquenessCheckForm.Text = "Проверить уникальность обозначений в списке Выпустить"
    $UniquenessCheckForm.AutoSizeMode = "GrowAndShrink"
    $UniquenessCheckForm.WindowState = "Normal"
    $UniquenessCheckForm.SizeGripStyle = "Hide"
    $UniquenessCheckForm.ShowInTaskbar = $true
    $UniquenessCheckForm.StartPosition = "CenterScreen"
    $UniquenessCheckForm.MinimizeBox = $false
    $UniquenessCheckForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор
    $UniquenessCheckFormRegister = New-Object System.Windows.Forms.Button
    $UniquenessCheckFormRegister.Location = New-Object System.Drawing.Point(10,10) #x,y
    $UniquenessCheckFormRegister.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UniquenessCheckFormRegister.Text = "Обзор..."
    $UniquenessCheckFormRegister.TabStop = $false
    $UniquenessCheckFormRegister.Add_Click({
        $script:PathToRegister = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:PathToRegister -ne $null) {
            $UniquenessCheckFormApplyButton.Enabled = $true
            $UniquenessCheckFormRegisterLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToRegister -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($UniquenessCheckFormRegisterLabel, $script:PathToRegister)
            #Write-Host $script:MakeChangesPathToFilesBeingPublished
        }
    })
    $UniquenessCheckForm.Controls.Add($UniquenessCheckFormRegister)
    #Поле к кнопке Обзор
    $UniquenessCheckFormRegisterLabel = New-Object System.Windows.Forms.Label
    $UniquenessCheckFormRegisterLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $UniquenessCheckFormRegisterLabel.Width = 600
    $UniquenessCheckFormRegisterLabel.Text = "Укажите путь к файлу учета программ и проектной документации"
    $UniquenessCheckFormRegisterLabel.TextAlign = "TopLeft"
    $UniquenessCheckForm.Controls.Add($UniquenessCheckFormRegisterLabel)
    #Кнопка Начать
    $UniquenessCheckFormApplyButton = New-Object System.Windows.Forms.Button
    $UniquenessCheckFormApplyButton.Location = New-Object System.Drawing.Point(10,50) #x,y
    $UniquenessCheckFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UniquenessCheckFormApplyButton.Text = "Начать"
    $UniquenessCheckFormApplyButton.Enabled = $false
    $UniquenessCheckFormApplyButton.Add_Click({
        if (Test-Path -Path "$PSScriptRoot\Проверка уникальности обозначений.html") {
            if ((Show-MessageBox -Message "Отчет Проверка уникальности обозначений.html уже существует в папке.`r`n`r`nНажмите Да, чтобы продолжить (отчет будет перезаписан).`r`nНажмите Нет, чтобы приостановить проверку." -Title "Отчет уже существует" -Type YesNo) -eq "Yes") {
            Remove-Item -Path "$PSScriptRoot\Проверка уникальности обозначений.html"
            Check-FileNameUniqueness
            $UniquenessCheckForm.Close()
            }
        } else {
            Check-FileNameUniqueness
            $UniquenessCheckForm.Close()
        }
    })
    $UniquenessCheckForm.Controls.Add($UniquenessCheckFormApplyButton)
    #Кнопка закрыть
    $UniquenessCheckFormCancelButton = New-Object System.Windows.Forms.Button
    $UniquenessCheckFormCancelButton.Location = New-Object System.Drawing.Point(100,50) #x,y
    $UniquenessCheckFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UniquenessCheckFormCancelButton.Text = "Закрыть"
    $UniquenessCheckFormCancelButton.Add_Click({
        $UniquenessCheckForm.Close()
    })
    $UniquenessCheckForm.Controls.Add($UniquenessCheckFormCancelButton)
    if ($script:PathToRegister -ne $null) {
            $UniquenessCheckFormApplyButton.Enabled = $true
            $UniquenessCheckFormRegisterLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToRegister -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($UniquenessCheckFormRegisterLabel, $script:PathToRegister)
            #Write-Host $script:MakeChangesPathToFilesBeingPublished
    }
    $UniquenessCheckForm.ShowDialog()
}

Function Apply-ColoringToRegister ()
{
    Write-Host "Применяю заливку к автоматически заполненным строкам в файле учета программ и проектной документации..."
    $RegisterApplyColoring = New-Object -ComObject Excel.Application
    $RegisterApplyColoring.Visible = $true
    $RegisterWorkbookApplyColoring = $RegisterApplyColoring.WorkBooks.Open($script:PathToRegisterColoring)
    $RegisterWorksheetApplyColoring = $RegisterWorkbookApplyColoring.Worksheets.Item(1)
    if ($RegisterWorksheetApplyColoring.AutoFilterMode -eq $true) {$RegisterWorksheetApplyColoring.ShowAllData()}
    $RegisterLastRowApplyColoring = $RegisterWorksheetApplyColoring.Cells.Item($RegisterWorksheetApplyColoring.Rows.Count, "A").End(-4162).Row
    $RegisterWorksheetApplyColoring.Range("A2:A$($RegisterLastRowApplyColoring)").AutoFilter(1, 139, 8)
    $LoopRangeApplyColoring = $RegisterWorksheetApplyColoring.Range("A2:A$($RegisterLastRowApplyColoring)")
    Foreach ($Cell in $LoopRangeApplyColoring.SpecialCells(12)) {
        if ($RegisterWorksheetApplyColoring.Cells.Item($Cell.Row, 2).Interior.Color -eq 16436871) {
        $RegisterWorksheetApplyColoring.Range($RegisterWorksheetApplyColoring.Cells.Item($Cell.Row, 1), $RegisterWorksheetApplyColoring.Cells.Item($Cell.Row, 14)).Interior.Color = 14336204
        } else {
        $RegisterWorksheetApplyColoring.Cells.Item($Cell.Row, 1).Interior.Color = 16777215
        }
    }
    Write-Host "ЗАЛИВКА ПРИМЕНИНА. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
    Write-Host "ЗАЛИВКА ПРИМЕНИНА. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
    Write-Host "ЗАЛИВКА ПРИМЕНИНА. МОЖНО РАБОТАТЬ С ФАЙЛОМ УЧЕТА."
}

Function ApplyColoringToRegisterForm ()
{
    $ApplyColoringToRegisterForm = New-Object System.Windows.Forms.Form
    $ApplyColoringToRegisterForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ApplyColoringToRegisterForm.ShowIcon = $false
    $ApplyColoringToRegisterForm.AutoSize = $true
    $ApplyColoringToRegisterForm.Text = "Применить заливку в файле учета программ и проектной документации"
    $ApplyColoringToRegisterForm.AutoSizeMode = "GrowAndShrink"
    $ApplyColoringToRegisterForm.WindowState = "Normal"
    $ApplyColoringToRegisterForm.SizeGripStyle = "Hide"
    $ApplyColoringToRegisterForm.ShowInTaskbar = $true
    $ApplyColoringToRegisterForm.StartPosition = "CenterScreen"
    $ApplyColoringToRegisterForm.MinimizeBox = $false
    $ApplyColoringToRegisterForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор
    $ApplyColoringToRegisterFormRegister = New-Object System.Windows.Forms.Button
    $ApplyColoringToRegisterFormRegister.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ApplyColoringToRegisterFormRegister.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyColoringToRegisterFormRegister.Text = "Обзор..."
    $ApplyColoringToRegisterFormRegister.TabStop = $false
    $ApplyColoringToRegisterFormRegister.Add_Click({
        $script:PathToRegisterColoring = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:PathToRegisterColoring -ne $null) {
            $ApplyColoringToRegisterFormApplyButton.Enabled = $true
            $ApplyColoringToRegisterFormRegisterLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToRegisterColoring -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyColoringToRegisterFormRegisterLabel, $script:PathToRegisterColoring)
            #Write-Host $script:MakeChangesPathToFilesBeingPublished
        }
    })
    $ApplyColoringToRegisterForm.Controls.Add($ApplyColoringToRegisterFormRegister)
    #Поле к кнопке Обзор
    $ApplyColoringToRegisterFormRegisterLabel = New-Object System.Windows.Forms.Label
    $ApplyColoringToRegisterFormRegisterLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $ApplyColoringToRegisterFormRegisterLabel.Width = 600
    $ApplyColoringToRegisterFormRegisterLabel.Text = "Укажите путь к файлу учета программ и проектной документации"
    $ApplyColoringToRegisterFormRegisterLabel.TextAlign = "TopLeft"
    $ApplyColoringToRegisterForm.Controls.Add($ApplyColoringToRegisterFormRegisterLabel)
    #Кнопка Начать
    $ApplyColoringToRegisterFormApplyButton = New-Object System.Windows.Forms.Button
    $ApplyColoringToRegisterFormApplyButton.Location = New-Object System.Drawing.Point(10,50) #x,y
    $ApplyColoringToRegisterFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyColoringToRegisterFormApplyButton.Text = "Начать"
    $ApplyColoringToRegisterFormApplyButton.Enabled = $false
    $ApplyColoringToRegisterFormApplyButton.Add_Click({
    Apply-ColoringToRegister
    $ApplyColoringToRegisterForm.Close()
    })
    $ApplyColoringToRegisterForm.Controls.Add($ApplyColoringToRegisterFormApplyButton)
    #Кнопка закрыть
    $ApplyColoringToRegisterFormCancelButton = New-Object System.Windows.Forms.Button
    $ApplyColoringToRegisterFormCancelButton.Location = New-Object System.Drawing.Point(100,50) #x,y
    $ApplyColoringToRegisterFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyColoringToRegisterFormCancelButton.Text = "Закрыть"
    $ApplyColoringToRegisterFormCancelButton.Add_Click({
        $ApplyColoringToRegisterForm.Close()
    })
    $ApplyColoringToRegisterForm.Controls.Add($ApplyColoringToRegisterFormCancelButton)
    if ($script:PathToRegisterColoring -ne $null) {
            $ApplyColoringToRegisterFormApplyButton.Enabled = $true
            $ApplyColoringToRegisterFormRegisterLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToRegisterColoring -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyColoringToRegisterFormRegisterLabel, $script:PathToRegisterColoring)
            #Write-Host $script:MakeChangesPathToFilesBeingPublished
    }
    $ApplyColoringToRegisterForm.ShowDialog()
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
            #Write-Host $script:MakeChangesPathToFilesBeingPublished
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
            #Write-Host $script:MakeChangesPathToCurrentVersion
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
            #Write-Host $script:MakeChangesPathToArchiveFolder
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
            #Write-Host $script:MakeChangesPathToBackupFolder
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
                    $CheckResult = Check-Conditions -BackupFlag $MakeChangesBackupFlag
                    if ($CheckResult -eq $true) {$ApplyChangesFormApplyButton.Enabled = $false; Invoke-Item "$PSScriptRoot\Ошибки.html"} else {Show-MessageBox -Message "Ошибки не обнаружены.`r`n`r`nТеперь вы можете начать процедуру внесения изменений по ИИ." -Title "Ошибки не обнаружены" -Type OK ;$ApplyChangesFormApplyButton.Enabled = $true}
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
        Invoke-Item "$PSScriptRoot\Отчет.html"
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

Function ClientReleaseForm ()
{
    $script:ImportFilesArray = @()
    $ClientReleaseForm = New-Object System.Windows.Forms.Form
    $ClientReleaseForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ClientReleaseForm.ShowIcon = $false
    $ClientReleaseForm.AutoSize = $true
    $ClientReleaseForm.Text = "Релиз для клиента"
    $ClientReleaseForm.AutoSizeMode = "GrowAndShrink"
    $ClientReleaseForm.WindowState = "Normal"
    $ClientReleaseForm.SizeGripStyle = "Hide"
    $ClientReleaseForm.ShowInTaskbar = $true
    $ClientReleaseForm.StartPosition = "CenterScreen"
    $ClientReleaseForm.MinimizeBox = $false
    $ClientReleaseForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор для файла
    $ClientReleaseFormBrowseFileButton = New-Object System.Windows.Forms.Button
    $ClientReleaseFormBrowseFileButton.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ClientReleaseFormBrowseFileButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClientReleaseFormBrowseFileButton.Text = "Обзор..."
    $ClientReleaseFormBrowseFileButton.TabStop = $false
    $ClientReleaseFormBrowseFileButton.Add_Click({
    $script:SelectedWordFile = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:SelectedWordFile -ne $null) {
            $ClientReleaseFormBrowseButtonFileLabel.Text = "Указанная спецификация: $(Split-Path -Path $script:SelectedWordFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ClientReleaseFormBrowseButtonFileLabel, $script:SelectedWordFile)
        } else {
            $ClientReleaseFormBrowseButtonFileLabel.Text = "Выберите спецификацию, которая содержит ссылки на архивы с исходным кодом"
        }
    })
    $ClientReleaseForm.Controls.Add($ClientReleaseFormBrowseFileButton)
    #Поле к кнопке Обзор для файла
    $ClientReleaseFormBrowseButtonFileLabel = New-Object System.Windows.Forms.Label
    $ClientReleaseFormBrowseButtonFileLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $ClientReleaseFormBrowseButtonFileLabel.Width = 725
    $ClientReleaseFormBrowseButtonFileLabel.Text = "Выберите спецификацию, которая содержит ссылки на архивы с исходным кодом"
    $ClientReleaseFormBrowseButtonFileLabel.TextAlign = "TopLeft"
    $ClientReleaseForm.Controls.Add($ClientReleaseFormBrowseButtonFileLabel)
    #Кнопка обзор для папки
    $ClientReleaseFormBrowseFolderButton = New-Object System.Windows.Forms.Button
    $ClientReleaseFormBrowseFolderButton.Location = New-Object System.Drawing.Point(10,42) #x,y
    $ClientReleaseFormBrowseFolderButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClientReleaseFormBrowseFolderButton.Text = "Обзор..."
    $ClientReleaseFormBrowseFolderButton.TabStop = $false
    $ClientReleaseFormBrowseFolderButton.Add_Click({
        $script:SelectedClientFolder = Select-Folder -Description "Выберите папку, в которой необходимо удалить архивы с исходным кодом и исходные документы"
        if ($script:SelectedClientFolder -ne $null) {
            $ClientReleaseFormBrowseButtonFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:SelectedClientFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ClientReleaseFormBrowseButtonFolderLabel, $script:SelectedClientFolder)
        } else {
            $ClientReleaseFormBrowseButtonFolderLabel.Text = "Выберите папку, в которой необходимо удалить архивы с исходным кодом и исходные документы"
        }
    })
    $ClientReleaseForm.Controls.Add($ClientReleaseFormBrowseFolderButton)
    #Поле к кнопке Обзор для папки
    $ClientReleaseFormBrowseButtonFolderLabel = New-Object System.Windows.Forms.Label
    $ClientReleaseFormBrowseButtonFolderLabel.Location =  New-Object System.Drawing.Point(95,46) #x,y
    $ClientReleaseFormBrowseButtonFolderLabel.Width = 725
    $ClientReleaseFormBrowseButtonFolderLabel.Text = "Выберите папку, в которой необходимо удалить архивы с исходным кодом и исходные документы"
    $ClientReleaseFormBrowseButtonFolderLabel.TextAlign = "TopLeft"
    $ClientReleaseForm.Controls.Add($ClientReleaseFormBrowseButtonFolderLabel)
    #Обновление полей
    if ($script:SelectedWordFile -ne $null) {
        $ClientReleaseFormBrowseButtonFileLabel.Text = "Указанная спецификация: $(Split-Path -Path $script:SelectedWordFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ClientReleaseFormBrowseButtonFileLabel, $script:SelectedWordFile)
    }
    if ($script:SelectedClientFolder -ne $null) {
        $ClientReleaseFormBrowseButtonFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:SelectedClientFolder -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($ClientReleaseFormBrowseButtonFolderLabel, $script:SelectedClientFolder)
    }
    #Чекбокс 'Удалить файлы MS Word'
    $ClientReleaseFormDeleteMsOfficeFiles = New-Object System.Windows.Forms.CheckBox
    $ClientReleaseFormDeleteMsOfficeFiles.Width = 410
    $ClientReleaseFormDeleteMsOfficeFiles.Text = "Удалить файлы приложения MS Word"
    $ClientReleaseFormDeleteMsOfficeFiles.Location = New-Object System.Drawing.Point(10,72) #x,y
    $ClientReleaseFormDeleteMsOfficeFiles.Enabled = $true
    $ClientReleaseFormDeleteMsOfficeFiles.Checked = $true
    $ClientReleaseFormDeleteMsOfficeFiles.Add_CheckStateChanged({})
    $ClientReleaseForm.Controls.Add($ClientReleaseFormDeleteMsOfficeFiles)
    #Чекбокс 'Удалить файлы MS Excel'
    $ClientReleaseFormDeleteMsExcelFiles = New-Object System.Windows.Forms.CheckBox
    $ClientReleaseFormDeleteMsExcelFiles.Width = 410
    $ClientReleaseFormDeleteMsExcelFiles.Text = "Удалить файлы приложения MS Excel"
    $ClientReleaseFormDeleteMsExcelFiles.Location = New-Object System.Drawing.Point(10,97) #x,y
    $ClientReleaseFormDeleteMsExcelFiles.Enabled = $true
    $ClientReleaseFormDeleteMsExcelFiles.Checked = $true
    $ClientReleaseFormDeleteMsExcelFiles.Add_CheckStateChanged({})
    $ClientReleaseForm.Controls.Add($ClientReleaseFormDeleteMsExcelFiles)
    #Кнопка Начать
    $ClientReleaseFormApplyButton = New-Object System.Windows.Forms.Button
    $ClientReleaseFormApplyButton.Location = New-Object System.Drawing.Point(10,134) #x,y
    $ClientReleaseFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClientReleaseFormApplyButton.Text = "Начать"
    $ClientReleaseFormApplyButton.Enabled = $true
    $ClientReleaseFormApplyButton.Add_Click({
    if ($script:SelectedClientFolder -eq $null -and $script:SelectedWordFile -eq $null) {
        Show-MessageBox -Message "Не указана спецификация, содержащая список архивов с исходным кодом, и папка, в которой необходимо удалить архивы с исходным кодом и исходные документы." -Title "Невозможно выполнить операцию" -Type OK
    } elseif ($script:SelectedWordFile -eq $null) {
        Show-MessageBox -Message "Не указана спецификация, содержащая список архивов с исходным кодом." -Title "Невозможно выполнить операцию" -Type OK
    } elseif ($script:SelectedClientFolder -eq $null) {
        Show-MessageBox -Message "Не указана папка, в которой необходимо удалить архивы с исходным кодом и исходные документы." -Title "Невозможно выполнить операцию." -Type OK
    } else {
        if ((Show-MessageBox -Message "Перед началом операции убедитесь в том, что у вас нет открытых Word документов.`r`nВо время работы скрипт закроет все Word документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
        if ($ClientReleaseFormDeleteMsOfficeFiles.Checked -eq $true) {$DeleteWordFlag = $true} else {$DeleteWordFlag = $false}
        if ($ClientReleaseFormDeleteMsExcelFiles.Checked -eq $true) {$DeleteExcelFlag = $true} else {$DeleteExcelFlag = $false}
        Create-ClientVersion -PathToSpecification $script:SelectedWordFile -PathToClientFolder $script:SelectedClientFolder -DeleteWordFlag $DeleteWordFlag -DeleteExcelFlag $DeleteExcelFlag
        $ClientReleaseForm.Close()
        }
    }
    })
    $ClientReleaseForm.Controls.Add($ClientReleaseFormApplyButton)
    #Кнопка закрыть
    $ClientReleaseFormCancelButton = New-Object System.Windows.Forms.Button
    $ClientReleaseFormCancelButton.Location = New-Object System.Drawing.Point(100,134) #x,y
    $ClientReleaseFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ClientReleaseFormCancelButton.Text = "Закрыть"
    $ClientReleaseFormCancelButton.Add_Click({
        $ClientReleaseForm.Close()
    })
    $ClientReleaseForm.Controls.Add($ClientReleaseFormCancelButton)
    $ClientReleaseForm.ShowDialog()
}

Function Create-ClientVersion ($PathToSpecification, $PathToClientFolder, $DeleteWordFlag, $DeleteExcelFlag)
{
    Kill -Name WINWORD -ErrorAction SilentlyContinue
    Write-Host "Выполняю операцию..."
    Start-Sleep -Seconds 1
    $script:counter = 1
    Write-Host "Открываю спецификацию..."
    #Создать экземпляр приложения MS Word
    $WordApp = New-Object -ComObject Word.Application
    #Создать документ MS Word
    $WordDocumentSpecification = $WordApp.Documents.Open($PathToSpecification)
    #Сделать вызванное приложение невидемым
    $WordApp.Visible = $false
    $CatchedArchives = @()
    Write-Host "Собираю ссылки на архивы с исходным кодом..."
    $WordDocumentSpecification.Tables.Item(1).Rows | % {
        if ($_.Cells.Count -eq 7 -and ((($_.Cells.Item(5).Range.Text).Trim([char]0x0007)).Trim(' ') -replace [char]13, '') -match "архив исходных кодов") {
            $CatchedArchives += (($_.Cells.Item(4).Range.Text).Trim([char]0x0007)).Trim(' ')  -replace [char]13, ''
        }
    }
    $WordDocumentSpecification.Close([ref]0)
    $WordApp.Quit()
    Kill -Name WINWORD -ErrorAction SilentlyContinue
    Write-Host "Удаляю архивы, указанные в спецификации, из указанной папки..."
    Start-Sleep -Seconds 2
    $CatchedArchives | % {
       Remove-Item -Path "$($PathToClientFolder)\$_" -ErrorAction SilentlyContinue
    }
    Write-Host "Удаляю документы приложения MS Word из указанной папки..."
    Start-Sleep -Seconds 2
    if ($DeleteWordFlag -eq $true) {Get-ChildItem -Path "$PathToClientFolder\*.docx" | % {Remove-Item $_ -ErrorAction SilentlyContinue}}
    if ($DeleteWordFlag -eq $true) {Get-ChildItem -Path "$PathToClientFolder\*.doc" | % {Remove-Item $_ -ErrorAction SilentlyContinue}}
    Write-Host "Удаляю документы приложения MS Excel из указанной папки..."
    Start-Sleep -Seconds 2
    if ($DeleteExcelFlag -eq $true) {Get-ChildItem -Path "$PathToClientFolder\*.xlsx" | % {Remove-Item $_ -ErrorAction SilentlyContinue}}
    if ($DeleteExcelFlag -eq $true) {Get-ChildItem -Path "$PathToClientFolder\*.xls" | % {Remove-Item $_ -ErrorAction SilentlyContinue}}
    Write-Host "Операция выполнена. Релиз для клиента готов."
}

Function UpdateSubjectTemplate ($Department, $Conjugation, $NotificationNumber, $Project, $TemplateInputField)
{
    $TemplateString = ""
    $TemplateString = "$Department "
    if ($Conjugation -eq $null) {$TemplateString += "<спряжение> извещение об изменении "} else {$TemplateString += "$Conjugation извещение об изменении "}
    if ($NotificationNumber -eq "") {$TemplateString += "<номер извещения> для проекта "} else {$TemplateString += "$NotificationNumber для проекта "}
    if ($Project -eq $null) {$TemplateString += "<название проекта>."} else {$TemplateString += "$Project."}
    #Write-Host "$TemplateString"
    $TemplateInputField.Text = $TemplateString
}

Function CreateLetterForm ()
{
    $CreateLetterForm = New-Object System.Windows.Forms.Form
    $CreateLetterForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    $CreateLetterForm.ShowIcon = $false
    $CreateLetterForm.AutoSize = $true
    $CreateLetterForm.Text = "Создать письмо"
    $CreateLetterForm.AutoSizeMode = "GrowAndShrink"
    $CreateLetterForm.WindowState = "Normal"
    $CreateLetterForm.SizeGripStyle = "Hide"
    $CreateLetterForm.ShowInTaskbar = $true
    $CreateLetterForm.StartPosition = "CenterScreen"
    $CreateLetterForm.MinimizeBox = $false
    $CreateLetterForm.MaximizeBox = $false
    #Группа элементов Настройка переменных письма
    $EmailSubjectSettingsGroup = New-Object System.Windows.Forms.GroupBox
    $EmailSubjectSettingsGroup.Location = New-Object System.Drawing.Point(10,10) #x,y
    $EmailSubjectSettingsGroup.Size = New-Object System.Drawing.Point(785,210) #width,height
    $EmailSubjectSettingsGroup.Text = "Настройка переменных письма"
    $CreateLetterForm.Controls.Add($EmailSubjectSettingsGroup)
    #Отдел
    $EmailSubjectDepartmentLabel = New-Object System.Windows.Forms.Label
    $EmailSubjectDepartmentLabel.Location =  New-Object System.Drawing.Point(10,25) #x,y
    $EmailSubjectDepartmentLabel.Width = 105
    $EmailSubjectDepartmentLabel.Text = "Отдел:"
    $EmailSubjectDepartmentLabel.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectDepartmentLabel)
    #Список, содержащий доступные названия отделов
    $EmailSubjectComboboxDepartmentName = New-Object System.Windows.Forms.ComboBox
    $EmailSubjectComboboxDepartmentName.Location = New-Object System.Drawing.Point(120,23) #x,y
    $EmailSubjectComboboxDepartmentName.DropDownStyle = "DropDownList"
    $EmailSubjectComboboxDepartmentName.Width = 200
    $ComboboxDepartmentName.Items | % {$EmailSubjectComboboxDepartmentName.Items.Add($_)}
    $EmailSubjectComboboxDepartmentName.SelectedItem = $ComboboxDepartmentName.SelectedItem
    $EmailSubjectComboboxDepartmentName.Add_SelectedIndexChanged({
    if (Test-Path "$PSScriptRoot\Отделы.xml") {
        $XmlPopulateDepartmentMetadata = New-Object System.Xml.XmlDocument
        $XmlPopulateDepartmentMetadata.Load("$PSScriptRoot\Отделы.xml")
        if ($XmlPopulateDepartmentMetadata.SelectNodes("//name[.='$($EmailSubjectComboboxDepartmentName.SelectedItem)' and @conjugation]").Count -eq 1) {
            $EmailSubjectComboboxConjugation.SelectedItem = "$($XmlPopulateDepartmentMetadata.SelectSingleNode("//name[.='$($EmailSubjectComboboxDepartmentName.SelectedItem)' and @conjugation]").Attributes.GetNamedItem("conjugation").Value)"
        }
    }
    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
    })
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectComboboxDepartmentName)
    #Спряжение
    $EmailSubjectConjugationLabel = New-Object System.Windows.Forms.Label
    $EmailSubjectConjugationLabel.Location =  New-Object System.Drawing.Point(10,55) #x,y
    $EmailSubjectConjugationLabel.Width = 105
    $EmailSubjectConjugationLabel.Text = "Спряжение:"
    $EmailSubjectConjugationLabel.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectConjugationLabel)
    #Список, содержащий спряжения
    $ConjugationsArray = @("выпустил", "выпустило")
    $EmailSubjectComboboxConjugation = New-Object System.Windows.Forms.ComboBox
    $EmailSubjectComboboxConjugation.Location = New-Object System.Drawing.Point(120,53) #x,y
    $EmailSubjectComboboxConjugation.DropDownStyle = "DropDownList"
    $EmailSubjectComboboxConjugation.Width = 200
    $ConjugationsArray | % {$EmailSubjectComboboxConjugation.Items.Add($_)}
    $EmailSubjectComboboxConjugation.Add_SelectedIndexChanged({
    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
    })
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectComboboxConjugation)
    #Надпись к полю Номер извещения
    $EmailSubjectNotificationNumber = New-Object System.Windows.Forms.Label
    $EmailSubjectNotificationNumber.Location =  New-Object System.Drawing.Point(10,85) #x,y
    $EmailSubjectNotificationNumber.Width = 105
    $EmailSubjectNotificationNumber.Text = "Номер извещения:"
    $EmailSubjectNotificationNumber.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectNotificationNumber)
    #Поле для ввода Номера извещения
    $EmailSubjectNotificationNumberInput = New-Object System.Windows.Forms.MaskedTextBox
    $EmailSubjectNotificationNumberInput.Location = New-Object System.Drawing.Point(120,83) #x,y
    $EmailSubjectNotificationNumberInput.Width = 62
    $EmailSubjectNotificationNumberInput.Mask = "00-00-0000"
    $EmailSubjectNotificationNumberInput.ForeColor = "Black"
    $EmailSubjectNotificationNumberInput.Text = "$($UpdateNotificationNumberInput.Text)"
    $EmailSubjectNotificationNumberInput.Add_TextChanged({
    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
    })
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectNotificationNumberInput)
    #Надпись к списку для указания названия проекта
    $EmailSubjectProjectLabel = New-Object System.Windows.Forms.Label
    $EmailSubjectProjectLabel.Location = New-Object System.Drawing.Point(10,115) #x,y
    $EmailSubjectProjectLabel.Width = 105
    $EmailSubjectProjectLabel.Text = "Проект:"
    $EmailSubjectProjectLabel.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectProjectLabel)
    #Список содержащий доступные названия проектов
    $EmailSubjectComboboxProjectName = New-Object System.Windows.Forms.ComboBox
    $EmailSubjectComboboxProjectName.Location = New-Object System.Drawing.Point(120,113) #x,y
    $EmailSubjectComboboxProjectName.DropDownStyle = "DropDownList"
    $EmailSubjectComboboxProjectName.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Проекты.xml") {Populate-List -List $EmailSubjectComboboxProjectName -PathToXml "$PSScriptRoot\Проекты.xml"}
    $EmailSubjectComboboxProjectName.Add_SelectedIndexChanged({
    $EmailSubjectAccessPathInput.Text = ""
    if (Test-Path "$PSScriptRoot\Проекты.xml") {
        $XmlPopulateProjectMetadata = New-Object System.Xml.XmlDocument
        $XmlPopulateProjectMetadata.Load("$PSScriptRoot\Проекты.xml")
        if ($XmlPopulateProjectMetadata.SelectNodes("//name[.='$($EmailSubjectComboboxProjectName.SelectedItem)' and @access_path]").Count -eq 1) {
            $EmailSubjectAccessPathInput.Text = "$($XmlPopulateProjectMetadata.SelectSingleNode("//name[.='$($EmailSubjectComboboxProjectName.SelectedItem)' and @access_path]").Attributes.GetNamedItem("access_path").Value)"
        }
    }
    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
    })
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectComboboxProjectName)
    #Кнопка Редактирования списка отделов
    $EmailSubjectButtonEditListOfProjects = New-Object System.Windows.Forms.Button
    $EmailSubjectButtonEditListOfProjects.Location = New-Object System.Drawing.Point(325,112) #x,y
    $EmailSubjectButtonEditListOfProjects.Size = New-Object System.Drawing.Point(22,23) #width,height
    $EmailSubjectButtonEditListOfProjects.Text = "..."
    $EmailSubjectButtonEditListOfProjects.Add_Click({Manage-CustomLists -ListType Projects})
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectButtonEditListOfProjects)
    #Адрес доступа (лейбл)
    $EmailSubjectAccessPathLabel = New-Object System.Windows.Forms.Label
    $EmailSubjectAccessPathLabel.Location =  New-Object System.Drawing.Point(10,145) #x,y
    $EmailSubjectAccessPathLabel.Width = 105
    $EmailSubjectAccessPathLabel.Text = "Адрес доступа:"
    $EmailSubjectAccessPathLabel.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectAccessPathLabel)
    #Поле адрес доступа
    $EmailSubjectAccessPathInput = New-Object System.Windows.Forms.TextBox 
    $EmailSubjectAccessPathInput.Location = New-Object System.Drawing.Point(120,143) #x,y
    $EmailSubjectAccessPathInput.Width = 565
    $EmailSubjectAccessPathInput.Text = ""
    $EmailSubjectAccessPathInput.Enabled = $true
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectAccessPathInput)
    #Кнопка обзор для адреса доступа
    $EmailSubjectAccessPathBrowseFolderButton = New-Object System.Windows.Forms.Button
    $EmailSubjectAccessPathBrowseFolderButton.Location = New-Object System.Drawing.Point(690,142) #x,y
    $EmailSubjectAccessPathBrowseFolderButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EmailSubjectAccessPathBrowseFolderButton.Text = "Обзор..."
    $EmailSubjectAccessPathBrowseFolderButton.TabStop = $false
    $EmailSubjectAccessPathBrowseFolderButton.Add_Click({
        $script:SelectedAccessPath = Select-Folder -Description "Укажите адрес доступа."
        if ($script:SelectedAccessPath -ne $null) {
            $EmailSubjectAccessPathInput.Text  = "$script:SelectedAccessPath"
        } else {
            $EmailSubjectAccessPathInput = ""
        }
    })
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectAccessPathBrowseFolderButton)
    #Шаблон
    $EmailSubjectTemplateLabel = New-Object System.Windows.Forms.Label
    $EmailSubjectTemplateLabel.Location =  New-Object System.Drawing.Point(10,175) #x,y
    $EmailSubjectTemplateLabel.Width = 105
    $EmailSubjectTemplateLabel.Text = "Шаблон:"
    $EmailSubjectTemplateLabel.TextAlign = "TopRight"
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectTemplateLabel)
    #Поле шаблона
    $EmailSubjectTemplateInput = New-Object System.Windows.Forms.TextBox 
    $EmailSubjectTemplateInput.Location = New-Object System.Drawing.Point(120,173) #x,y
    $EmailSubjectTemplateInput.Width = 650
    <#"$($ComboboxDepartmentName.SelectedItem) <спряжение> извещение об изменении <номер извещения> для проекта <название проекта>"#>
    $EmailSubjectTemplateInput.Enabled = $false
    $EmailSubjectSettingsGroup.Controls.Add($EmailSubjectTemplateInput)
    #Группа элементов Настройка списка рассылки
    $RecipientSettingsGroup = New-Object System.Windows.Forms.GroupBox
    $RecipientSettingsGroup.Location = New-Object System.Drawing.Point(10,230) #x,y
    $RecipientSettingsGroup.Size = New-Object System.Drawing.Point(785,55) #width,height
    $RecipientSettingsGroup.Text = "Настройка списка рассылки"
    $CreateLetterForm.Controls.Add($RecipientSettingsGroup)
    #Чекбокс 'Использовать стандартный список рассылки'
    $EmailSubjectFormDefaulRecipients = New-Object System.Windows.Forms.CheckBox
    $EmailSubjectFormDefaulRecipients.Width = 410
    $EmailSubjectFormDefaulRecipients.Text = "Использовать стандартный список рассылки"
    $EmailSubjectFormDefaulRecipients.Location = New-Object System.Drawing.Point(10,20) #x,y
    $EmailSubjectFormDefaulRecipients.Enabled = $true
    $EmailSubjectFormDefaulRecipients.Checked = $true
    $EmailSubjectFormDefaulRecipients.Add_CheckStateChanged({})
    $RecipientSettingsGroup.Controls.Add($EmailSubjectFormDefaulRecipients)
    #Кнопка Начать
    $EmailSubjectFormApplyButton = New-Object System.Windows.Forms.Button
    $EmailSubjectFormApplyButton.Location = New-Object System.Drawing.Point(10,310) #x,y
    $EmailSubjectFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EmailSubjectFormApplyButton.Text = "Начать"
    $EmailSubjectFormApplyButton.Enabled = $true
    $EmailSubjectFormApplyButton.Add_Click({
        $TextInMessage = "Не указаны или некорректно указаны следующие параметры:`r`n"
        $ErrorPresent = $false
        if ($EmailSubjectComboboxConjugation.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано спряжение."}
        if ($EmailSubjectNotificationNumberInput.Text -eq '  -  -') {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан номер извещения."}
        if ($EmailSubjectNotificationNumberInput.Text -ne '  -  -') {if ($EmailSubjectNotificationNumberInput.Text -notmatch '\d\d-\d\d-\d\d\d\d') {$ErrorPresent = $true; $TextInMessage += "`r`nНомер извещения указан неполностью, либо содержит недопустимые символы."}}
        if ($EmailSubjectComboboxProjectName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран проект."}
        if ($EmailSubjectAccessPathInput.Text -eq "") {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан адрес доступа."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать генерацию электронного письма" -Type OK
        } else {
            Save-MetaData -DataType Projects -SelectedItem $EmailSubjectComboboxProjectName.SelectedItem
            Save-MetaData -DataType Departments -SelectedItem $EmailSubjectComboboxDepartmentName.SelectedItem
            Build-OutlookMessage -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -AccessPath $EmailSubjectAccessPathInput.Text
            $CreateLetterForm.Close()
        }
    })
    $CreateLetterForm.Controls.Add($EmailSubjectFormApplyButton)
    #Кнопка закрыть
    $EmailSubjectFormCancelButton = New-Object System.Windows.Forms.Button
    $EmailSubjectFormCancelButton.Location = New-Object System.Drawing.Point(100,310) #x,y
    $EmailSubjectFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $EmailSubjectFormCancelButton.Text = "Закрыть"
    $EmailSubjectFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $EmailSubjectFormCancelButton.Add_Click({
        $CreateLetterForm.Close()
    })
    $CreateLetterForm.Controls.Add($EmailSubjectFormCancelButton)
    if (Test-Path "$PSScriptRoot\Отделы.xml") {
        $XmlPopulateDepartmentMetadata = New-Object System.Xml.XmlDocument
        $XmlPopulateDepartmentMetadata.Load("$PSScriptRoot\Отделы.xml")
        if ($XmlPopulateDepartmentMetadata.SelectNodes("//name[.='$($EmailSubjectComboboxDepartmentName.SelectedItem)' and @conjugation]").Count -eq 1) {
            $EmailSubjectComboboxConjugation.SelectedItem = "$($XmlPopulateDepartmentMetadata.SelectSingleNode("//name[.='$($EmailSubjectComboboxDepartmentName.SelectedItem)' and @conjugation]").Attributes.GetNamedItem("conjugation").Value)"
        }
    }
    UpdateSubjectTemplate -Department $EmailSubjectComboboxDepartmentName.SelectedItem -Conjugation $EmailSubjectComboboxConjugation.SelectedItem -NotificationNumber $EmailSubjectNotificationNumberInput.Text -Project $EmailSubjectComboboxProjectName.SelectedItem -TemplateInputField $EmailSubjectTemplateInput
    $CreateLetterForm.ShowDialog()
}

Function UpdateRegisterForm ()
{
    $UpdateRegisterForm = New-Object System.Windows.Forms.Form
    $UpdateRegisterForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $UpdateRegisterForm.ShowIcon = $false
    $UpdateRegisterForm.AutoSize = $true
    $UpdateRegisterForm.Text = "Внести изменения в файл учета программ и ПД"
    $UpdateRegisterForm.AutoSizeMode = "GrowAndShrink"
    $UpdateRegisterForm.WindowState = "Normal"
    $UpdateRegisterForm.SizeGripStyle = "Hide"
    $UpdateRegisterForm.ShowInTaskbar = $true
    $UpdateRegisterForm.StartPosition = "CenterScreen"
    $UpdateRegisterForm.MinimizeBox = $false
    $UpdateRegisterForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор для файла
    $UpdateRegisterFormBrowseFileButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormBrowseFileButton.Location = New-Object System.Drawing.Point(10,10) #x,y
    $UpdateRegisterFormBrowseFileButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormBrowseFileButton.Text = "Обзор..."
    $UpdateRegisterFormBrowseFileButton.TabStop = $false
    $UpdateRegisterFormBrowseFileButton.Add_Click({
    $script:SelectedRegister = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:SelectedRegister -ne $null) {
            $UpdateRegisterFormBrowseButtonFileLabel.Text = "Указанный файл учета: $(Split-Path -Path $script:SelectedRegister -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($UpdateRegisterFormBrowseButtonFileLabel, $script:SelectedRegister)
        } else {
            $UpdateRegisterFormBrowseButtonFileLabel.Text = "Выберите файл учета, в который необходимо внести изменения"
        }
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseFileButton)
    #Поле к кнопке Обзор для файла
    $UpdateRegisterFormBrowseButtonFileLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBrowseButtonFileLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $UpdateRegisterFormBrowseButtonFileLabel.Width = 725
    $UpdateRegisterFormBrowseButtonFileLabel.Text = "Выберите файл учета, в который необходимо внести изменения"
    $UpdateRegisterFormBrowseButtonFileLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseButtonFileLabel)
    #Кнопка обзор для папки
    $UpdateRegisterFormBrowseFolderButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormBrowseFolderButton.Location = New-Object System.Drawing.Point(10,42) #x,y
    $UpdateRegisterFormBrowseFolderButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormBrowseFolderButton.Text = "Обзор..."
    $UpdateRegisterFormBrowseFolderButton.TabStop = $false
    $UpdateRegisterFormBrowseFolderButton.Add_Click({
        $script:SelectedFolderWithFilesBeingPublished = Select-Folder -Description "Выберите папку с спецификациями, в которых упоминаются файлы из извещения об изменении"
        if ($script:SelectedFolderWithFilesBeingPublished -ne $null) {
            $UpdateRegisterFormBrowseButtonFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:SelectedFolderWithFilesBeingPublished -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($UpdateRegisterFormBrowseButtonFolderLabel, $script:SelectedFolderWithFilesBeingPublished)
        } else {
            $UpdateRegisterFormBrowseButtonFolderLabel.Text = "Выберите папку с спецификациями, в которых упоминаются файлы из извещения об изменении"
        }
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseFolderButton)
    #Поле к кнопке Обзор для папки
    $UpdateRegisterFormBrowseButtonFolderLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormBrowseButtonFolderLabel.Location =  New-Object System.Drawing.Point(95,46) #x,y
    $UpdateRegisterFormBrowseButtonFolderLabel.Width = 725
    $UpdateRegisterFormBrowseButtonFolderLabel.Text = "Выберите папку с спецификациями, в которых упоминаются файлы из извещения об изменении"
    $UpdateRegisterFormBrowseButtonFolderLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormBrowseButtonFolderLabel)
    #Обновление полей
    if ($script:SelectedRegister -ne $null) {
        $UpdateRegisterFormBrowseButtonFileLabel.Text = "Указанный файл учета: $(Split-Path -Path $script:SelectedRegister -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($UpdateRegisterFormBrowseButtonFileLabel, $script:SelectedRegister)
    }
    if ($script:SelectedFolderWithFilesBeingPublished -ne $null) {
        $UpdateRegisterFormBrowseButtonFolderLabel.Text = "Указанная папка: $(Split-Path -Path $script:SelectedFolderWithFilesBeingPublished -Leaf). Наведите курсором, чтобы увидеть полный путь."
        $ToolTip.SetToolTip($UpdateRegisterFormBrowseButtonFolderLabel, $script:SelectedFolderWithFilesBeingPublished)
    }
    #Надпись к списку для указания названия проекта
    $UpdateRegisterFormProjectLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormProjectLabel.Location = New-Object System.Drawing.Point(10,78) #x,y
    $UpdateRegisterFormProjectLabel.Width = 75
    $UpdateRegisterFormProjectLabel.Text = "Код проекта:"
    $UpdateRegisterFormProjectLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormProjectLabel)
    #Список содержащий доступные названия проектов
    $UpdateRegisterFormComboboxProjectName = New-Object System.Windows.Forms.ComboBox
    $UpdateRegisterFormComboboxProjectName.Location = New-Object System.Drawing.Point(85,75) #-3x,y
    $UpdateRegisterFormComboboxProjectName.DropDownStyle = "DropDownList"
    $UpdateRegisterFormComboboxProjectName.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Проекты.xml") {Populate-List -List $UpdateRegisterFormComboboxProjectName -PathToXml "$PSScriptRoot\Проекты.xml"}
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormComboboxProjectName)
    #Кнопка Редактирования списка отделов
    $UpdateRegisterFormButtonEditListOfProjects = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormButtonEditListOfProjects.Location = New-Object System.Drawing.Point(290,74) #-4x,y
    $UpdateRegisterFormButtonEditListOfProjects.Size = New-Object System.Drawing.Point(22,23) #width,height
    $UpdateRegisterFormButtonEditListOfProjects.Text = "..."
    $UpdateRegisterFormButtonEditListOfProjects.Add_Click({Manage-CustomLists -ListType RegisterProjects})
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormButtonEditListOfProjects)
    #Надпись к списку для указания разработчика
    $UpdateRegisterFormDeveloperLabel = New-Object System.Windows.Forms.Label
    $UpdateRegisterFormDeveloperLabel.Location = New-Object System.Drawing.Point(10,106) #x,y
    $UpdateRegisterFormDeveloperLabel.Width = 75
    $UpdateRegisterFormDeveloperLabel.Text = "Разработчик:"
    $UpdateRegisterFormDeveloperLabel.TextAlign = "TopLeft"
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormDeveloperLabel)
    #Список содержащий доступных разработчиков
    $UpdateRegisterFormComboboxDeveloperName = New-Object System.Windows.Forms.ComboBox
    $UpdateRegisterFormComboboxDeveloperName.Location = New-Object System.Drawing.Point(85,103) #-3x,y
    $UpdateRegisterFormComboboxDeveloperName.DropDownStyle = "DropDownList"
    $UpdateRegisterFormComboboxDeveloperName.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Разработчики.xml") {Populate-List -List $UpdateRegisterFormComboboxDeveloperName -PathToXml "$PSScriptRoot\Разработчики.xml"}
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormComboboxDeveloperName)
    #Кнопка Редактирования списка разработчиков
    $UpdateRegisterFormButtonEditListOfDevelopers = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormButtonEditListOfDevelopers.Location = New-Object System.Drawing.Point(290,102) #-4x,y
    $UpdateRegisterFormButtonEditListOfDevelopers.Size = New-Object System.Drawing.Point(22,23) #width,height
    $UpdateRegisterFormButtonEditListOfDevelopers.Text = "..."
    $UpdateRegisterFormButtonEditListOfDevelopers.Add_Click({Manage-CustomLists -ListType Developers})
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormButtonEditListOfDevelopers)
    #Кнопка Начать
    $UpdateRegisterFormApplyButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormApplyButton.Location = New-Object System.Drawing.Point(10,154) #x,y
    $UpdateRegisterFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormApplyButton.Text = "Начать"
    $UpdateRegisterFormApplyButton.Enabled = $true
    $UpdateRegisterFormApplyButton.Add_Click({
        $TextInMessage = "Не указаны следующие параметры:`r`n"
        $ErrorPresent = $false
        if ($script:SelectedRegister -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к файлу учета."}
        if ($script:SelectedFolderWithFilesBeingPublished -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан путь к папке с спецификациями, в которых упоминаются публикуемые файлы."}
        if ($UpdateRegisterFormComboboxProjectName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран код проект."}
        if ($UpdateRegisterFormComboboxDeveloperName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран разработчик."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно выполнить операцию" -Type OK
        } else {
            if ((Show-MessageBox -Message "Перед началом операции убедитесь в том, что у вас нет открытых Word и Excel документов.`r`nВо время работы скрипт закроет все Word и Excel документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                Populate-Register
                $UpdateRegisterForm.Close()
            }
        }
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormApplyButton)
    #Кнопка закрыть
    $UpdateRegisterFormCancelButton = New-Object System.Windows.Forms.Button
    $UpdateRegisterFormCancelButton.Location = New-Object System.Drawing.Point(100,154) #x,y
    $UpdateRegisterFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $UpdateRegisterFormCancelButton.Text = "Закрыть"
    $UpdateRegisterFormCancelButton.Add_Click({
        $UpdateRegisterForm.Close()
    })
    $UpdateRegisterForm.Controls.Add($UpdateRegisterFormCancelButton)
    $UpdateRegisterForm.ShowDialog()
}

Function Save-MetaData ([ValidateSet("Departments", "Projects")]$DataType, $SelectedItem)
{
    if ($DataType -eq "Projects") {
        if (Test-Path -Path "$PSScriptRoot\Проекты.xml") {
            $XmlToSaveMetaData = New-Object System.Xml.XmlDocument
            $XmlToSaveMetaData.Load("$PSScriptRoot\Проекты.xml")
                $XmlToSaveMetaData.SelectNodes("//name") | % {if ($_.InnerXml -eq $SelectedItem) {
                    $ElementAttribute = $XmlToSaveMetaData.CreateAttribute("access_path")
                    $ElementAttribute.Value = "$($EmailSubjectAccessPathInput.Text)"
                    $_.Attributes.Append($ElementAttribute)
                }
            }
            $XmlToSaveMetaData.Save("$PSScriptRoot\Проекты.xml")
        }
    }
    if ($DataType -eq "Departments") {
        if (Test-Path -Path "$PSScriptRoot\Отделы.xml") {
            $XmlToSaveMetaData = New-Object System.Xml.XmlDocument
            $XmlToSaveMetaData.Load("$PSScriptRoot\Отделы.xml")
                $XmlToSaveMetaData.SelectNodes("//name") | % {if ($_.InnerXml -eq $SelectedItem) {
                    $ElementAttribute = $XmlToSaveMetaData.CreateAttribute("conjugation")
                    $ElementAttribute.Value = "$($EmailSubjectComboboxConjugation.SelectedItem)"
                    $_.Attributes.Append($ElementAttribute)
                }
            }
            $XmlToSaveMetaData.Save("$PSScriptRoot\Отделы.xml")
        }
    }
}

Function Build-OutlookMessage ($Department, $Conjugation, $NotificationNumber, $Project, $AccessPath)
{
if (Test-Path -Path "$PSScriptRoot\email.html") {Remove-Item -Path "$PSScriptRoot\email.html"}
Add-Content "$PSScriptRoot\email.html" "<!DOCTYPE html>
<html lang=""ru"">
<head>
<meta charset=""utf-8"">
<title>Электронное сообщение</title>
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
.Version {
    width: 3%;
    text-align: center;
}
.Filename {
    width:47%;
    text-align: left;
}
.Comment {
    width:50%;
    text-align: left;
}
</style>
</head>
<body>
$Department $Conjugation извещение об изменении $NotificationNumber для проекта $Project.
<br>
<br>
Заменен(ы):
<br>" -Encoding UTF8
Add-TableToEmail -List $ListViewReplace -Action "replace"
Add-TableToEmail -List $ListViewRemove -Action "remove"
Add-TableToEmail -List $ListViewAdd -Action "publish"
Add-Content "$PSScriptRoot\email.html" "<br>
Адрес доступа:  $AccessPath
<br>
Замененные, аннулированные программы и ПД, извещение об изменении – в архиве проекта в папке $NotificationNumber.
"
$ol = New-Object -comObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.Subject = "$Department $Conjugation извещение об изменении $NotificationNumber для проекта $Project"
if ($EmailSubjectFormDefaulRecipients.Checked -eq $true) {$mail.To = $script:SendTo}
#$mail.CC = "Копию, можно указать автоматически"
$inspector = $mail.GetInspector
$inspector.Display()
$mailrange = $inspector.WordEditor.Application.Selection
$mailrange.InsertFile("$PSScriptRoot\email.html", "", $false, $false, $false)
Remove-Item -Path "$PSScriptRoot\email.html"
}

Function Add-TableToEmail ($List, $Action)
{
Add-Content "$PSScriptRoot\email.html" "<table style=""width:100%"">
    <tr>
        <th class=""Version"">Изм.</th>
        <th class=""Filename"">Обозначение</th>
        <th class=""Comment"">Примечание</th>
    </tr>
" -Encoding UTF8
    ForEach ($Item in $List.Items) {
        if ($Item.SubItems[2].Text -eq "Документ") {
            Add-Content "$PSScriptRoot\email.html" "    <tr>
        <td class=""Version"">$($Item.SubItems[1].Text)</td>
        <td class=""Filename"">$($Item.Text)</td>
        <td class=""Comment""></td>
    </tr>
" -Encoding UTF8  
        }
        if ($Item.SubItems[2].Text -eq "Программа") {
            Add-Content "$PSScriptRoot\email.html" "    <tr>
        <td class=""Version"">-</td>
        <td class=""Filename"">$($Item.Text)</td>
        <td class=""Comment"">$($Item.SubItems[1].Text)</td>
    </tr>
" -Encoding UTF8  
        }
    }
Add-Content "$PSScriptRoot\email.html" "</table>" -Encoding UTF8  
if ($Action -eq "replace") {
Add-Content "$PSScriptRoot\email.html" "<br>
Аннулирован(ы):
" -Encoding UTF8
}
if ($Action -eq "remove") {
Add-Content "$PSScriptRoot\email.html" "<br>
Выпущен(ы):
" -Encoding UTF8
}
}

Function Custom-Form ()
{
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    #Главное окно
    $ScriptMainWindow = New-Object System.Windows.Forms.Form
    $ScriptMainWindow.ShowIcon = $false
    $ScriptMainWindow.AutoSize = $true
    $ScriptMainWindow.Text = "ИИ"
    $ScriptMainWindow.AutoSizeMode = "GrowAndShrink"
    $ScriptMainWindow.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $ScriptMainWindow.SizeGripStyle = "Hide"
    $ScriptMainWindow.ShowInTaskbar = $true
    $ScriptMainWindow.StartPosition = "CenterScreen"
    $ScriptMainWindow.MinimizeBox = $true
    $ScriptMainWindow.MaximizeBox = $false
    $ScriptMainWindow.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    #Группа элементов Настройка списков
    $ListSettingsGroup = New-Object System.Windows.Forms.GroupBox
    $ListSettingsGroup.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ListSettingsGroup.Size = New-Object System.Drawing.Point(1308,555) #width,height
    $ListSettingsGroup.Text = "Настройка списков"
    $ScriptMainWindow.Controls.Add($ListSettingsGroup)
   
    #Надпись к списку Выпустить
    $ListViewAddLabel = New-Object System.Windows.Forms.Label
    $ListViewAddLabel.Location =  New-Object System.Drawing.Point(10,50) #x,y
    $ListViewAddLabel.Width = 200
    $ListViewAddLabel.Height = 15
    $ListViewAddLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewAddLabel.Text = "Выпустить (0):"
    $ListViewAddLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewAddLabel)
    #Список Выпустить
    $ListViewAdd = New-Object System.Windows.Forms.ListView
    $ListViewAdd.Location = New-Object System.Drawing.Point(10,66) #x, y
    $ListViewAdd.View = "Details"
    $ListViewAdd.FullRowSelect = $true
    $ListViewAdd.MultiSelect = $false
    $ListViewAdd.HideSelection = $false
    $ListViewAdd.Width = 400
    $ListViewAdd.Height = 370
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Обозначение" -Width 267 | Out-Null
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Изм./MD5" -Width 69 | Out-Null
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Тип" -Width 43 | Out-Null
    $ListViewAdd_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewAdd.Columns[0].Width -ne 267) {
            $ListViewAdd.Columns[0].Width = 267
        }
        if ($ListViewAdd.Columns[1].Width -ne 69) {
            $ListViewAdd.Columns[1].Width = 69
        }
        if ($ListViewAdd.Columns[2].Width -ne 43) {
            $ListViewAdd.Columns[2].Width = 43
        }
    }
    $ListViewAdd.Add_ItemSelectionChanged({
    if ($ListViewAdd.SelectedItems[0].Index -ne $null) {
        $SelectedIndexAdd = $ListViewAdd.SelectedItems[0].Index
        Unselect-ItemsInOtherLists -List1 $ListViewReplace -List2 $ListViewRemove 
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $true
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        $ListViewAdd.Items[$SelectedIndexAdd].Selected = $true
        Update-SelectedFileDetails -FileName $ListViewAdd.Items[$SelectedIndexAdd].Text -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute $ListViewAdd.Items[$SelectedIndexAdd].SubItems[1].Text -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType $ListViewAdd.Items[$SelectedIndexAdd].SubItems[2].Text -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    if ($ListViewAdd.SelectedIndices.Count -eq 0 -and $ListViewReplace.SelectedIndices.Count -eq 0 -and $ListViewRemove.SelectedIndices.Count -eq 0) {
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-SelectedFileDetails -FileName "" -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute "" -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType "" -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    })
    $ListViewAdd.add_ColumnWidthChanged($ListViewAdd_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewAdd)
    
    #Button Move-to-left between Add and Replace
    $ButtonMoveToLeftBetweenAddAndReplace = New-Object System.Windows.Forms.Button
    $ButtonMoveToLeftBetweenAddAndReplace.Location = New-Object System.Drawing.Point(420,205) #x,y
    $ButtonMoveToLeftBetweenAddAndReplace.Size = New-Object System.Drawing.Point(24,24) #width,height
    $ButtonMoveToLeftBetweenAddAndReplace.Text = "<"
    $ButtonMoveToLeftBetweenAddAndReplace.Add_Click({
        Move-ItemToAnotherList -MoveFrom $ListViewReplace -MoveTo $ListViewAdd
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
    $ListSettingsGroup.Controls.Add($ButtonMoveToLeftBetweenAddAndReplace)

    #Button Move-to-right between Add and Replace
    $ButtonMoveToRightBetweenAddAndReplace = New-Object System.Windows.Forms.Button
    $ButtonMoveToRightBetweenAddAndReplace.Location = New-Object System.Drawing.Point(420,265) #x,y
    $ButtonMoveToRightBetweenAddAndReplace.Size = New-Object System.Drawing.Point(24,24) #width,height
    $ButtonMoveToRightBetweenAddAndReplace.Text = ">"
    $ButtonMoveToRightBetweenAddAndReplace.Add_Click({
        Move-ItemToAnotherList -MoveFrom $ListViewAdd -MoveTo $ListViewReplace
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
    $ListSettingsGroup.Controls.Add($ButtonMoveToRightBetweenAddAndReplace)
    
    #Надпись к списку Заменить
    $ListViewReplaceLabel = New-Object System.Windows.Forms.Label
    $ListViewReplaceLabel.Location =  New-Object System.Drawing.Point(454,50) #x,y
    $ListViewReplaceLabel.Width = 200
    $ListViewReplaceLabel.Height = 15
    $ListViewReplaceLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewReplaceLabel.Text = "Заменить (0):"
    $ListViewReplaceLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewReplaceLabel)

    #Список Заменить
    $ListViewReplace = New-Object System.Windows.Forms.ListView
    $ListViewReplace.Location = New-Object System.Drawing.Point(454,66) #x, y
    $ListViewReplace.View = "Details"
    $ListViewReplace.FullRowSelect = $true
    $ListViewReplace.MultiSelect = $false
    $ListViewReplace.HideSelection = $false
    $ListViewReplace.Width = 400
    $ListViewReplace.Height = 370
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Обозначение" -Width 267 | Out-Null
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Изм./MD5" -Width 69 | Out-Null
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Тип" -Width 43 | Out-Null
    $ListViewReplace_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewReplace.Columns[0].Width -ne 267) {
            $ListViewReplace.Columns[0].Width = 267
        }
        if ($ListViewReplace.Columns[1].Width -ne 69) {
            $ListViewReplace.Columns[1].Width = 69
        }
        if ($ListViewReplace.Columns[2].Width -ne 43) {
            $ListViewReplace.Columns[2].Width = 43
        }
    }
    $ListViewReplace.Add_ItemSelectionChanged({
    if ($ListViewReplace.SelectedItems[0].Index -ne $null) {
        $SelectedIndexReplace = $ListViewReplace.SelectedItems[0].Index
        Unselect-ItemsInOtherLists -List1 $ListViewAdd -List2 $ListViewRemove
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $true
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $true
        $ListViewReplace.Items[$SelectedIndexReplace].Selected = $true
        Update-SelectedFileDetails -FileName $ListViewReplace.Items[$SelectedIndexReplace].Text -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute $ListViewReplace.Items[$SelectedIndexReplace].SubItems[1].Text -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType $ListViewReplace.Items[$SelectedIndexReplace].SubItems[2].Text -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    if ($ListViewAdd.SelectedIndices.Count -eq 0 -and $ListViewReplace.SelectedIndices.Count -eq 0 -and $ListViewRemove.SelectedIndices.Count -eq 0) {
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-SelectedFileDetails -FileName "" -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute "" -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType "" -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    })
    $ListViewReplace.add_ColumnWidthChanged($ListViewReplace_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewReplace)

    #Button Move-to-right between Replace and Remove
    $ButtonMoveToRightBetweenReplaceAndRemove = New-Object System.Windows.Forms.Button
    $ButtonMoveToRightBetweenReplaceAndRemove.Location = New-Object System.Drawing.Point(864,205) #x,y
    $ButtonMoveToRightBetweenReplaceAndRemove.Size = New-Object System.Drawing.Point(24,24) #width,height
    $ButtonMoveToRightBetweenReplaceAndRemove.Text = ">"
    $ButtonMoveToRightBetweenReplaceAndRemove.Add_Click({
        Move-ItemToAnotherList -MoveFrom $ListViewReplace -MoveTo $ListViewRemove
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
    $ListSettingsGroup.Controls.Add($ButtonMoveToRightBetweenReplaceAndRemove)
    
    #Button Move-to-left between Replace and Remove
    $ButtonMoveToLeftBetweenReplaceAndRemove = New-Object System.Windows.Forms.Button
    $ButtonMoveToLeftBetweenReplaceAndRemove.Location = New-Object System.Drawing.Point(864,265) #x,y
    $ButtonMoveToLeftBetweenReplaceAndRemove.Size = New-Object System.Drawing.Point(24,24) #width,height
    $ButtonMoveToLeftBetweenReplaceAndRemove.Text = "<"
    $ButtonMoveToLeftBetweenReplaceAndRemove.Add_Click({
        Move-ItemToAnotherList -MoveFrom $ListViewRemove -MoveTo $ListViewReplace
        $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
        $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
        $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
    $ListSettingsGroup.Controls.Add($ButtonMoveToLeftBetweenReplaceAndRemove)

    #Надпись к списку Аннулировать
    $ListViewRemoveLabel = New-Object System.Windows.Forms.Label
    $ListViewRemoveLabel.Location =  New-Object System.Drawing.Point(898,50) #x,y
    $ListViewRemoveLabel.Width = 200
    $ListViewRemoveLabel.Height = 15
    $ListViewRemoveLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewRemoveLabel.Text = "Аннулировать (0):"
    $ListViewRemoveLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewRemoveLabel)

    #Список Аннулировать
    $ListViewRemove = New-Object System.Windows.Forms.ListView
    $ListViewRemove.Location = New-Object System.Drawing.Point(898,66) #x, y
    $ListViewRemove.View = "Details"
    $ListViewRemove.FullRowSelect = $true
    $ListViewRemove.MultiSelect = $false
    $ListViewRemove.HideSelection = $false
    $ListViewRemove.Width = 400
    $ListViewRemove.Height = 370
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Обозначение" -Width 267 | Out-Null
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Изм./MD5" -Width 69 | Out-Null
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Тип" -Width 43 | Out-Null
        $ListViewRemove_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewRemove.Columns[0].Width -ne 267) {
            $ListViewRemove.Columns[0].Width = 267
        }
        if ($ListViewRemove.Columns[1].Width -ne 69) {
            $ListViewRemove.Columns[1].Width = 69
        }
        if ($ListViewRemove.Columns[2].Width -ne 43) {
            $ListViewRemove.Columns[2].Width = 43
        }
    }
    $ListViewRemove.Add_ItemSelectionChanged({
    if ($ListViewRemove.SelectedItems[0].Index -ne $null) {
    $SelectedIndexRemove = $ListViewRemove.SelectedItems[0].Index
    Unselect-ItemsInOtherLists -List1 $ListViewAdd -List2 $ListViewReplace
    $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $true
    $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
    $ListViewRemove.Items[$SelectedIndexRemove].Selected = $true
    Update-SelectedFileDetails -FileName $ListViewRemove.Items[$SelectedIndexRemove].Text -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute $ListViewRemove.Items[$SelectedIndexRemove].SubItems[1].Text -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType $ListViewRemove.Items[$SelectedIndexRemove].SubItems[2].Text -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    if ($ListViewAdd.SelectedIndices.Count -eq 0 -and $ListViewReplace.SelectedIndices.Count -eq 0 -and $ListViewRemove.SelectedIndices.Count -eq 0) {
    $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
    $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false
    Update-SelectedFileDetails -FileName "" -FileNameLabel $ListSettingsSelectedItemFileName -FileAttribute "" -FileAttributeLabel $ListSettingsSelectedItemFileAttribute -FileType "" -FileTypeLabel $ListSettingsSelectedItemFileType
    }
    })
    $ListViewRemove.add_ColumnWidthChanged($ListViewRemove_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewRemove)
    #Надпись к Всего файлов в списках
    $ListSettingsGroupTotalEntries = New-Object System.Windows.Forms.Label
    $ListSettingsGroupTotalEntries.Location =  New-Object System.Drawing.Point(10,25) #x,y
    $ListSettingsGroupTotalEntries.Width = 200
    $ListSettingsGroupTotalEntries.Height = 15
    $ListSettingsGroupTotalEntries.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Regular)
    $ListSettingsGroupTotalEntries.Text = "Всего файлов в списках: 0"
    $ListSettingsGroupTotalEntries.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListSettingsGroupTotalEntries)

    #Группа элементов Выбранная запись
    $ListSettingsSelectedItem = New-Object System.Windows.Forms.GroupBox
    $ListSettingsSelectedItem.Location = New-Object System.Drawing.Point(10,445) #x,y
    $ListSettingsSelectedItem.Size = New-Object System.Drawing.Point(513,100) #width,height
    $ListSettingsSelectedItem.Text = "Выбранная запись"
    $ListSettingsGroup.Controls.Add($ListSettingsSelectedItem)
    #Поле Обозначение для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileName = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileName.Location =  New-Object System.Drawing.Point(15,20) #x,y
    $ListSettingsSelectedItemFileName.Width = 490
    $ListSettingsSelectedItemFileName.Height = 15
    $ListSettingsSelectedItemFileName.Text = "Обозначение:"
    $ListSettingsSelectedItemFileName.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileName)
    #Поле Изм./MD5 для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileAttribute = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileAttribute.Location =  New-Object System.Drawing.Point(15,45) #x,y
    $ListSettingsSelectedItemFileAttribute.Width = 490
    $ListSettingsSelectedItemFileAttribute.Height = 15
    $ListSettingsSelectedItemFileAttribute.Text = "Изм./MD5:"
    $ListSettingsSelectedItemFileAttribute.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileAttribute)
    #Поле Тип файла для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileType = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileType.Location =  New-Object System.Drawing.Point(15,70) #x,y
    $ListSettingsSelectedItemFileType.Width = 490
    $ListSettingsSelectedItemFileType.Height = 15
    $ListSettingsSelectedItemFileType.Text = "Тип файла:"
    $ListSettingsSelectedItemFileType.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileType)

    #Группа элементов Действия с записью
    $ListSettingsItemActions = New-Object System.Windows.Forms.GroupBox
    $ListSettingsItemActions.Location = New-Object System.Drawing.Point(567,445) #x,y
    $ListSettingsItemActions.Size = New-Object System.Drawing.Point(287,100) #width,height
    $ListSettingsItemActions.Text = "Действия с записью"
    $ListSettingsGroup.Controls.Add($ListSettingsItemActions)
    #Добавить запись
    $ButtonAddItem = New-Object System.Windows.Forms.Button
    $ButtonAddItem.Location = New-Object System.Drawing.Point(10,17) #x,y
    $ButtonAddItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonAddItem.Text = "Добавить..."
    $ButtonAddItem.Add_Click({Add-ItemToList})
    $ListSettingsItemActions.Controls.Add($ButtonAddItem)
    #Редактировать запись
    $ButtonEditItemOnList = New-Object System.Windows.Forms.Button
    $ButtonEditItemOnList.Location = New-Object System.Drawing.Point(10,43) #x,y
    $ButtonEditItemOnList.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonEditItemOnList.Text = "Редактировать..."
    $ButtonEditItemOnList.Add_Click({
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {Edit-ItemOnList -ListObject $ListViewAdd}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {Edit-ItemOnList -ListObject $ListViewReplace}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {Edit-ItemOnList -ListObject $ListViewRemove}
    })
    $ListSettingsItemActions.Controls.Add($ButtonEditItemOnList)
    #Удалить запись
    $ButtonDeleteItem = New-Object System.Windows.Forms.Button
    $ButtonDeleteItem.Location = New-Object System.Drawing.Point(10,69) #x,y
    $ButtonDeleteItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonDeleteItem.Text = "Удалить"
    $ButtonDeleteItem.Add_Click({
    if ($ListViewAdd.SelectedIndices.Count -ne 0 -or $ListViewReplace.SelectedIndices.Count -ne 0 -or $ListViewRemove.SelectedIndices.Count -ne 0) {
        if ((Show-MessageBox -Title "Подтвердите действие" -Type YesNo -Message "Вы уверены, что хотите удалить выбранную запись из списка?") -eq "Yes") {
            if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].Remove()}
            if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].Remove()}
            if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].Remove()}
            $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
            $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
            $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
            $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false 
            Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
        }
    }
    })
    $ListSettingsItemActions.Controls.Add($ButtonDeleteItem)
    #Выделить цветом
    $ButtonMarkWithColor = New-Object System.Windows.Forms.Button
    $ButtonMarkWithColor.Location = New-Object System.Drawing.Point(140,17) #x,y
    $ButtonMarkWithColor.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonMarkWithColor.Text = "Выделить цветом"
    $ButtonMarkWithColor.Add_Click({
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
    })
    $ListSettingsItemActions.Controls.Add($ButtonMarkWithColor)
    #Выбрать цвет
    $ButtonSelectColor = New-Object System.Windows.Forms.Button
    $ButtonSelectColor.Location = New-Object System.Drawing.Point(255,17) #x,y
    $ButtonSelectColor.Size = New-Object System.Drawing.Point(22,22) #width,height
    $ColorDialog = New-Object System.Windows.Forms.ColorDialog
    $ButtonSelectColor.BackColor = [System.Drawing.Color]::LightGreen
    $ColorDialog.Color = [System.Drawing.Color]::LightGreen
    $ButtonSelectColor.Add_Click({
    $ColorDialog.ShowDialog()
    $ButtonSelectColor.BackColor = $ColorDialog.Color
    })
    $ListSettingsItemActions.Controls.Add($ButtonSelectColor)
    #Отменить выделение
    $ButtonRemoveColoring = New-Object System.Windows.Forms.Button
    $ButtonRemoveColoring.Location = New-Object System.Drawing.Point(140,43) #x,y
    $ButtonRemoveColoring.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonRemoveColoring.Text = "Отменить выделение"
    $ButtonRemoveColoring.Add_Click({
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].BackColor = [System.Drawing.Color]::White}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].BackColor = [System.Drawing.Color]::White}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].BackColor = [System.Drawing.Color]::White}
    Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
    })
    $ListSettingsItemActions.Controls.Add($ButtonRemoveColoring)
    <#Заполнить списки
    $ButtonPopulateLists = New-Object System.Windows.Forms.Button
    $ButtonPopulateLists.Location = New-Object System.Drawing.Point(140,69) #x,y
    $ButtonPopulateLists.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonPopulateLists.Text = "Заполнить списки"
    $ButtonPopulateLists.Add_Click({
              for ($i = 0; $i -lt 20; $i++) {
                $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("PABKRF-RU-EN-11.66.60.dUSM.00.00")
                $ItemToAdd.SubItems.Add("1")
                $ItemToAdd.SubItems.Add("Документ")
                $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                $ListViewReplace.Items.Add($ItemToAdd)
                }
                              for ($i = 0; $i -lt 20; $i++) {
                $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("PABKRF-RU-EN-22.66.60.dUSM.00.00")
                $ItemToAdd.SubItems.Add("1")
                $ItemToAdd.SubItems.Add("Документ")
                $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                $ListViewAdd.Items.Add($ItemToAdd)
                }
                              for ($i = 0; $i -lt 20; $i++) {
                $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("PABKRF_final_build.asx")
                $ItemToAdd.SubItems.Add("jsy45aso1l49sor89s78g8t6w4l6n8v9q")
                $ItemToAdd.SubItems.Add("Программа")
                $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                $ListViewRemove.Items.Add($ItemToAdd)
                }
    Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ListSettingsItemActions.Controls.Add($ButtonPopulateLists)
    #>
    #Группа элементов Действия со списками
    $ListSettingsListActions = New-Object System.Windows.Forms.GroupBox
    $ListSettingsListActions.Location = New-Object System.Drawing.Point(898,445) #x,y
    $ListSettingsListActions.Size = New-Object System.Drawing.Point(400,100) #width,height
    $ListSettingsListActions.Text = "Действия со списками"
    $ListSettingsGroup.Controls.Add($ListSettingsListActions)
    #Очистить списки
    $ButtonClearLists = New-Object System.Windows.Forms.Button
    $ButtonClearLists.Location = New-Object System.Drawing.Point(10,17) #x,y
    $ButtonClearLists.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonClearLists.Text = "Очистить..."
    $ButtonClearLists.Add_Click({
    Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove
    Clear-Lists
    })
    $ListSettingsListActions.Controls.Add($ButtonClearLists)
    #Отменить выделение
    $ButtonRemoveColoringLists = New-Object System.Windows.Forms.Button
    $ButtonRemoveColoringLists.Location = New-Object System.Drawing.Point(10,43) #x,y
    $ButtonRemoveColoringLists.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonRemoveColoringLists.Text = "Отменить выделение..."
    $ButtonRemoveColoringLists.Add_Click({
    Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove
    Discard-Coloring
    })
    $ListSettingsListActions.Controls.Add($ButtonRemoveColoringLists)
    #Чекбокс Отобразить сетку
    $CheckboxDisplayGrid = New-Object System.Windows.Forms.CheckBox
    $CheckboxDisplayGrid.Width = 150
    $CheckboxDisplayGrid.Text = "Отобразить сетку"
    $CheckboxDisplayGrid.Location = New-Object System.Drawing.Point(10,69) #x,y
    $CheckboxDisplayGrid.Enabled = $true
    $CheckboxDisplayGrid.Add_CheckStateChanged({
        if ($CheckboxDisplayGrid.Checked -eq $true) {
            $ListViewAdd.GridLines = $true; $ListViewReplace.GridLines = $true; $ListViewRemove.GridLines = $true
        } else {
            $ListViewAdd.GridLines = $false; $ListViewReplace.GridLines = $false; $ListViewRemove.GridLines = $false
        }
    })
    $ListSettingsListActions.Controls.Add($CheckboxDisplayGrid)
    #Импортировать из XML
    $ButtonImportFromXml = New-Object System.Windows.Forms.Button
    $ButtonImportFromXml.Location = New-Object System.Drawing.Point(167,17) #x,y
    $ButtonImportFromXml.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonImportFromXml.Text = "Загрузить из XML..."
    $ButtonImportFromXml.Add_Click({
    $SpecifiedFileForImport = Open-File -Filter "XML files (*.xml)| *.xml" -MultipleSelectionFlag $false
    if ($SpecifiedFileForImport -ne $null) {Import-FromXml -SpecifiedFile $SpecifiedFileForImport}
    })
    $ListSettingsListActions.Controls.Add($ButtonImportFromXml)
    #Экспортировать в XML
    $ButtonExportToXml = New-Object System.Windows.Forms.Button
    $ButtonExportToXml.Location = New-Object System.Drawing.Point(167,43) #x,y
    $ButtonExportToXml.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonExportToXml.Text = "Сохранить в XML..."
    $ButtonExportToXml.Add_Click({
    $SpecifiedExportPath = Save-File
    if ($SpecifiedExportPath -ne $null) {Export-ToXmlFile -SpecifiedFile $SpecifiedExportPath}
    })
    $ListSettingsListActions.Controls.Add($ButtonExportToXml)
    #Пакетный импорт файлов
    $ButtonBatchFileImport = New-Object System.Windows.Forms.Button
    $ButtonBatchFileImport.Location = New-Object System.Drawing.Point(167,69) #x,y
    $ButtonBatchFileImport.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonBatchFileImport.Text = "Пакетный импорт..."
    $ButtonBatchFileImport.Add_Click({BulkImportForm})
    $ListSettingsListActions.Controls.Add($ButtonBatchFileImport)

    #Группа элементов Параметры извещения
    $UpdateNotificationParameters = New-Object System.Windows.Forms.GroupBox
    $UpdateNotificationParameters.Location = New-Object System.Drawing.Point(10,575) #x,y
    $UpdateNotificationParameters.Size = New-Object System.Drawing.Point(659,145) #width,height
    $UpdateNotificationParameters.Text = "Параметры извещения"
    $ScriptMainWindow.Controls.Add($UpdateNotificationParameters)
    #Надпись к полю Номер извещения
    $UpdateNotificationNumber = New-Object System.Windows.Forms.Label
    $UpdateNotificationNumber.Location =  New-Object System.Drawing.Point(10,25) #x,y
    $UpdateNotificationNumber.Width = 147
    $UpdateNotificationNumber.Text = "Номер извещения:"
    $UpdateNotificationNumber.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($UpdateNotificationNumber)
    #Поле для ввода Номера извещения
    $UpdateNotificationNumberInput = New-Object System.Windows.Forms.MaskedTextBox
    $UpdateNotificationNumberInput.Location = New-Object System.Drawing.Point(162,23) #x,y
    $UpdateNotificationNumberInput.Width = 62
    $UpdateNotificationNumberInput.Mask = "00-00-0000"
    $UpdateNotificationNumberInput.ForeColor = "Black"
    $UpdateNotificationParameters.Controls.Add($UpdateNotificationNumberInput)
    #Надпись к календарю для указания Даты выпуска
    $CalendarIssueDateLabel = New-Object System.Windows.Forms.Label
    $CalendarIssueDateLabel.Location =  New-Object System.Drawing.Point(10,55) #x,y
    $CalendarIssueDateLabel.Width = 147
    $CalendarIssueDateLabel.Text = "Дата выпуска:"
    $CalendarIssueDateLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($CalendarIssueDateLabel)
    #Календарь для указания Даты выпуска
    $CalendarIssueDateInput = New-Object System.Windows.Forms.DateTimePicker
    $CalendarIssueDateInput.Location = New-Object System.Drawing.Point(162,53) #x,y
    $CalendarIssueDateInput.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $CalendarIssueDateInput.Width = 110
    #$CalendarIssueDateInput.Text = "03.02.1990"
    $UpdateNotificationParameters.Controls.Add($CalendarIssueDateInput)
    #Надпись к календарю для указания Срока внесения изменений
    $CalendarApplyUpdatesUntilLabel = New-Object System.Windows.Forms.Label
    $CalendarApplyUpdatesUntilLabel.Location =  New-Object System.Drawing.Point(10,85) #x,y
    $CalendarApplyUpdatesUntilLabel.Width = 147
    $CalendarApplyUpdatesUntilLabel.Text = "Срок внесения изменений:"
    $CalendarApplyUpdatesUntilLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($CalendarApplyUpdatesUntilLabel)
    #Календарь для указания Срока внесения изменений
    $CalendarApplyUpdatesUntilInput = New-Object System.Windows.Forms.DateTimePicker
    $CalendarApplyUpdatesUntilInput.Location = New-Object System.Drawing.Point(162,82) #x,y
    $CalendarApplyUpdatesUntilInput.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $CalendarApplyUpdatesUntilInput.Width = 110
    $UpdateNotificationParameters.Controls.Add($CalendarApplyUpdatesUntilInput)

    #Надпись к кнопке Остальные поля
    $CalendarApplyUpdatesOtherFieldsLabel = New-Object System.Windows.Forms.Label
    $CalendarApplyUpdatesOtherFieldsLabel.Location =  New-Object System.Drawing.Point(10,115) #x,y
    $CalendarApplyUpdatesOtherFieldsLabel.Width = 147
    $CalendarApplyUpdatesOtherFieldsLabel.Text = "Остальные поля:"
    $CalendarApplyUpdatesOtherFieldsLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($CalendarApplyUpdatesOtherFieldsLabel)

    #Кнопка Остальные поля
    $ButtonOtherFields = New-Object System.Windows.Forms.Button
    $ButtonOtherFields.Location = New-Object System.Drawing.Point(161,111) #x,y
    $ButtonOtherFields.Size = New-Object System.Drawing.Point(112,22) #width,height
    $ButtonOtherFields.Text = "Настроить..."
    $ButtonOtherFields.Add_Click({
    Setup-OtherFields
    })
    $UpdateNotificationParameters.Controls.Add($ButtonOtherFields)


    #Надпись к списку для указания Название отдела
    $ComboboxDepartmentNameLabel = New-Object System.Windows.Forms.Label
    $ComboboxDepartmentNameLabel.Location = New-Object System.Drawing.Point(317,25) #x,y
    $ComboboxDepartmentNameLabel.Width = 100
    $ComboboxDepartmentNameLabel.Text = "Отдел:"
    $ComboboxDepartmentNameLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($ComboboxDepartmentNameLabel)
    #Список содержащий доступные Названия отделов
    $ComboboxDepartmentName = New-Object System.Windows.Forms.ComboBox
    $ComboboxDepartmentName.Location = New-Object System.Drawing.Point(422,23) #x,y
    $ComboboxDepartmentName.DropDownStyle = "DropDownList"
    $ComboboxDepartmentName.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Отделы.xml") {Populate-List -List $ComboboxDepartmentName -PathToXml "$PSScriptRoot\Отделы.xml"}
    $UpdateNotificationParameters.Controls.Add($ComboboxDepartmentName)
    #Кнопка Редактирования списка отделов
    $ButtonEditListOfDepartments = New-Object System.Windows.Forms.Button
    $ButtonEditListOfDepartments.Location = New-Object System.Drawing.Point(627,22) #x,y
    $ButtonEditListOfDepartments.Size = New-Object System.Drawing.Point(22,23) #width,height
    $ButtonEditListOfDepartments.Text = "..."
    $ButtonEditListOfDepartments.Add_Click({Manage-CustomLists -ListType Departments})
    $UpdateNotificationParameters.Controls.Add($ButtonEditListOfDepartments)

    #Надпись к списку для указания Составил
    $ComboboxCreatedByLabel = New-Object System.Windows.Forms.Label
    $ComboboxCreatedByLabel.Location = New-Object System.Drawing.Point(317,55) #x,y
    $ComboboxCreatedByLabel.Width = 100
    $ComboboxCreatedByLabel.Text = "Выпустил:"
    $ComboboxCreatedByLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($ComboboxCreatedByLabel)
    #Список содержащий доступные ФИО
    $ComboboxCreatedBy = New-Object System.Windows.Forms.ComboBox
    $ComboboxCreatedBy.Location = New-Object System.Drawing.Point(422,53) #x,y
    $ComboboxCreatedBy.DropDownStyle = "DropDownList"
    $ComboboxCreatedBy.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Сотрудники.xml") {Populate-List -List $ComboboxCreatedBy -PathToXml "$PSScriptRoot\Сотрудники.xml"}
    $UpdateNotificationParameters.Controls.Add($ComboboxCreatedBy)
    #Кнопка Редактирования списка ФИО
    $ButtonEditListOfNamesCreatedBy = New-Object System.Windows.Forms.Button
    $ButtonEditListOfNamesCreatedBy.Location = New-Object System.Drawing.Point(627,52) #x,y
    $ButtonEditListOfNamesCreatedBy.Size = New-Object System.Drawing.Point(22,23) #width,height
    $ButtonEditListOfNamesCreatedBy.Text = "..."
    $ButtonEditListOfNamesCreatedBy.Add_Click({Manage-CustomLists -ListType Employees})
    $UpdateNotificationParameters.Controls.Add($ButtonEditListOfNamesCreatedBy)

    #Надпись к списку для указания Проверил
    $ComboboxCheckedByLabel = New-Object System.Windows.Forms.Label
    $ComboboxCheckedByLabel.Location = New-Object System.Drawing.Point(317,85) #x,y
    $ComboboxCheckedByLabel.Width = 100
    $ComboboxCheckedByLabel.Text = "Проверил:"
    $ComboboxCheckedByLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($ComboboxCheckedByLabel)
    #Список содержащий доступные ФИО
    $ComboboxCheckedBy = New-Object System.Windows.Forms.ComboBox
    $ComboboxCheckedBy.Location = New-Object System.Drawing.Point(422,82) #x,y
    $ComboboxCheckedBy.DropDownStyle = "DropDownList"
    $ComboboxCheckedBy.Width = 200
    if (Test-Path -Path "$PSScriptRoot\Сотрудники.xml") {Populate-List -List $ComboboxCheckedBy -PathToXml "$PSScriptRoot\Сотрудники.xml"}
    $UpdateNotificationParameters.Controls.Add($ComboboxCheckedBy)
    #Кнопка Редактирования списка ФИО
    $ButtonEditListOfNamesCheckedBy = New-Object System.Windows.Forms.Button
    $ButtonEditListOfNamesCheckedBy.Location = New-Object System.Drawing.Point(627,81) #x,y
    $ButtonEditListOfNamesCheckedBy.Size = New-Object System.Drawing.Point(22,23) #width,height
    $ButtonEditListOfNamesCheckedBy.Text = "..."
    $ButtonEditListOfNamesCheckedBy.Add_Click({Manage-CustomLists -ListType Employees})
    $UpdateNotificationParameters.Controls.Add($ButtonEditListOfNamesCheckedBy)

    #Надпись к списку для указания Кода
    $ComboboxCodeLabel = New-Object System.Windows.Forms.Label
    $ComboboxCodeLabel.Location = New-Object System.Drawing.Point(317,115) #x,y
    $ComboboxCodeLabel.Width = 100
    $ComboboxCodeLabel.Text = "Код:"
    $ComboboxCodeLabel.TextAlign = "TopRight"
    $UpdateNotificationParameters.Controls.Add($ComboboxCodeLabel)
    #Список содержащий доступные коды
    $ListOfСodes = @("1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16")
    $ComboboxCodes = New-Object System.Windows.Forms.ComboBox
    $ComboboxCodes.Location = New-Object System.Drawing.Point(422,112) #x,y
    $ComboboxCodes.DropDownStyle = "DropDownList"
    $ComboboxCodes.Width = 200
    $ListOfСodes | % {$ComboboxCodes.Items.Add($_)}
    $UpdateNotificationParameters.Controls.Add($ComboboxCodes)
    #Кнопка Редактирования списка ФИО
    $ButtonCodesHelp = New-Object System.Windows.Forms.Button
    $ButtonCodesHelp.Location = New-Object System.Drawing.Point(627,111) #x,y
    $ButtonCodesHelp.Size = New-Object System.Drawing.Point(22,23) #width,height
    $ButtonCodesHelp.Text = "?"
    $ButtonCodesHelp.Add_Click({
        Invoke-Item "$PSScriptRoot\Help\description_of_codes.html"
        #$PSScriptRoot\test\index.html
    })
    $UpdateNotificationParameters.Controls.Add($ButtonCodesHelp)

    #Кнопка запустить
    $ButtonRun = New-Object System.Windows.Forms.Button
    $ButtonRun.Location = New-Object System.Drawing.Point(680,580) #x,y
    $ButtonRun.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonRun.Text = "Создать ИИ"
    $ButtonRun.Add_Click({
        $TextInMessage = "Не указаны или некорректно указаны следующие параметры извещения:`r`n"
        $ErrorPresent = $false
        if ($UpdateNotificationNumberInput.Text -eq '  -  -') {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан номер извещения."}
        if ($UpdateNotificationNumberInput.Text -ne '  -  -') {if ($UpdateNotificationNumberInput.Text -notmatch '\d\d-\d\d-\d\d\d\d') {$ErrorPresent = $true; $TextInMessage += "`r`nНомер извещения указан неполностью, либо содержит недопустимые символы."}}
        if ($ComboboxCreatedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран отдел."}
        if ($ComboboxCheckedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Выпустил."}
        if ($ComboboxDepartmentName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Проверил."}
        if ($ComboboxCodes.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран код."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать генерацию ИИ" -Type OK
        } else {
            if (Test-Path -Path "$PSScriptRoot\$($UpdateNotificationNumberInput.Text).docx") {
                if ((Show-MessageBox -Message "Извещение с номером $($UpdateNotificationNumberInput.Text) уже существует в папке.`r`n`r`nНажмите Да, чтобы продолжить (извещение будет перезаписано).`r`nНажмите Нет, чтобы приостановить генерацию извещения." -Title "Извещение уже существует" -Type YesNo) -eq "Yes") {
                    if ((Show-MessageBox -Message "Перед началом генерации ИИ убедитесь в том, что у вас нет открытых Word документов.`r`nВо время работы скрипт закроет все Word документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                        Remove-Item -Path "$PSScriptRoot\$($UpdateNotificationNumberInput.Text).docx"
                        Start-Sleep -Seconds 2
                        $script:CounterForWordItems = 1
                        #Generate-UpdateNotification -NotificationName $UpdateNotificationNumberInput.Text
                        Move-ListsToWordDocument -NotificationName $UpdateNotificationNumberInput.Text
                        Export-ToXmlFile -SpecifiedFile "$PSScriptRoot\$($UpdateNotificationNumberInput.Text).xml"
                    }
                }
            } else {
                if ((Show-MessageBox -Message "Перед началом генерации ИИ убедитесь в том, что у вас нет открытых Word документов.`r`nВо время работы скрипт закроет все Word документы, не сохраняя их, что может привести к потере данных!`r`nПродолжить?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {
                    $script:CounterForWordItems = 1
                    #Generate-UpdateNotification -NotificationName $UpdateNotificationNumberInput.Text
                    Move-ListsToWordDocument -NotificationName $UpdateNotificationNumberInput.Text
                    Export-ToXmlFile -SpecifiedFile "$PSScriptRoot\$($UpdateNotificationNumberInput.Text).xml"
                }
            }
        }
    })
    $ScriptMainWindow.Controls.Add($ButtonRun)
    #Кнопка Провер. уник. обознач.
    $ButtonUniquenessCheckScript = New-Object System.Windows.Forms.Button
    $ButtonUniquenessCheckScript.Location = New-Object System.Drawing.Point(680,606) #x,y
    $ButtonUniquenessCheckScript.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonUniquenessCheckScript.Text = "Провер. уник. обознач."
    $ButtonUniquenessCheckScript.Add_Click({
        UniquenessCheckForm
    })
    $ScriptMainWindow.Controls.Add($ButtonUniquenessCheckScript)
    #Кнопка закрыть
    $ButtonCloseScript = New-Object System.Windows.Forms.Button
    $ButtonCloseScript.Location = New-Object System.Drawing.Point(680,632) #x,y
    $ButtonCloseScript.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonCloseScript.Text = "Закрыть скрипт"
    $ButtonCloseScript.Add_Click({
    if ((Show-MessageBox -Message "Закрыть скрипт?" -Title "Подтвердите действие" -Type YesNo) -eq "Yes") {$ScriptMainWindow.Close()}
    })
    $ScriptMainWindow.Controls.Add($ButtonCloseScript)
    #Кнопка Справка
    $ButtonHelp = New-Object System.Windows.Forms.Button
    $ButtonHelp.Location = New-Object System.Drawing.Point(680,658) #x,y
    $ButtonHelp.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonHelp.Text = "Справка"
    $ButtonHelp.Add_Click({
    Invoke-Item "$PSScriptRoot\Help\index.html"
    })
    $ScriptMainWindow.Controls.Add($ButtonHelp)
    #Кнопка Внести изменения
    $MakeChanges = New-Object System.Windows.Forms.Button
    $MakeChanges.Location = New-Object System.Drawing.Point(835,580) #x,y
    $MakeChanges.Size = New-Object System.Drawing.Point(137,22) #width,height
    $MakeChanges.Text = "Внести изменения..."
    $MakeChanges.Add_Click({
    ApplyChangesForm
    })
    $ScriptMainWindow.Controls.Add($MakeChanges)
    #Кнопка Релиз для клиента...
    $ClientRelease = New-Object System.Windows.Forms.Button
    $ClientRelease.Location = New-Object System.Drawing.Point(835,606) #x,y
    $ClientRelease.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ClientRelease.Text = "Релиз для клиента..."
    $ClientRelease.Add_Click({
    ClientReleaseForm
    })
    $ScriptMainWindow.Controls.Add($ClientRelease)
    #Кнопка Создать письмо...
    $CreateLetter = New-Object System.Windows.Forms.Button
    $CreateLetter.Location = New-Object System.Drawing.Point(835,632) #x,y
    $CreateLetter.Size = New-Object System.Drawing.Point(137,22) #width,height
    $CreateLetter.Text = "Создать письмо..."
    $CreateLetter.Add_Click({
        $TextInMessage = "Не указаны или некорректно указаны следующие параметры извещения:`r`n"
        $ErrorPresent = $false
        if ($UpdateNotificationNumberInput.Text -eq '  -  -') {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан номер извещения."}
        if ($UpdateNotificationNumberInput.Text -ne '  -  -') {if ($UpdateNotificationNumberInput.Text -notmatch '\d\d-\d\d-\d\d\d\d') {$ErrorPresent = $true; $TextInMessage += "`r`nНомер извещения указан неполностью, либо содержит недопустимые символы."}}
        if ($ComboboxCreatedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран отдел."}
        if ($ComboboxCheckedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Выпустил."}
        if ($ComboboxDepartmentName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Проверил."}
        if ($ComboboxCodes.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран код."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать создание электронного письма" -Type OK
        } else {
            CreateLetterForm
        }
    })
    $ScriptMainWindow.Controls.Add($CreateLetter)
    #Кнопка Внести в реестр...
    $UpdadeRegister = New-Object System.Windows.Forms.Button
    $UpdadeRegister.Location = New-Object System.Drawing.Point(835,658) #x,y
    $UpdadeRegister.Size = New-Object System.Drawing.Point(137,22) #width,height
    $UpdadeRegister.Text = "Внести в реестр..."
    $UpdadeRegister.Add_Click({
        $TextInMessage = "Не указаны или некорректно указаны следующие параметры извещения:`r`n"
        $ErrorPresent = $false
        if ($UpdateNotificationNumberInput.Text -eq '  -  -') {$ErrorPresent = $true; $TextInMessage += "`r`nНе указан номер извещения."}
        if ($UpdateNotificationNumberInput.Text -ne '  -  -') {if ($UpdateNotificationNumberInput.Text -notmatch '\d\d-\d\d-\d\d\d\d') {$ErrorPresent = $true; $TextInMessage += "`r`nНомер извещения указан неполностью, либо содержит недопустимые символы."}}
        if ($ComboboxCreatedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран отдел."}
        if ($ComboboxCheckedBy.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Выпустил."}
        if ($ComboboxDepartmentName.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбрано ФИО для поля Проверил."}
        if ($ComboboxCodes.SelectedItem -eq $null) {$ErrorPresent = $true; $TextInMessage += "`r`nНе выбран код."}
        if ($ErrorPresent -eq $true) {
            Show-MessageBox -Message $TextInMessage -Title "Невозможно начать создание электронного письма" -Type OK
        } else {
        UpdateRegisterForm
        }
    })
    $ScriptMainWindow.Controls.Add($UpdadeRegister)
    #Кнопка Прим. заливку в реестре...
    $ApplyColoringInRegister = New-Object System.Windows.Forms.Button
    $ApplyColoringInRegister.Location = New-Object System.Drawing.Point(835,684) #x,y
    $ApplyColoringInRegister.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ApplyColoringInRegister.Text = "Прим. зал. в реестре..."
    $ApplyColoringInRegister.Add_Click({ApplyColoringToRegisterForm})
    $ScriptMainWindow.Controls.Add($ApplyColoringInRegister)
    $ScriptMainWindow.ShowDialog()
}

Custom-Form | Out-Null
