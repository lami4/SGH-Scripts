clear
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
    #$MoveTo.Items.Add($ItemToAdd)
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
    $AddItemFormAttributeValueInput.Text = "Укажите Изм. или MD5..."
    $AddItemFormAttributeValueInput.ForeColor = "Gray"
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
    $AddItemFormAddButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $AddItemFormAddButton.Text = "Добавить запись"
    $AddItemFormAddButton.Add_Click({
        if ($AddItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение и Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($AddItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -ne "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($AddItemFormFileNameInput.Text -ne "Укажите обозначение..." -and $AddItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } else {
                $ItemsOnTheList = @()
                if ($AddItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}}
                if ($AddItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}}
                if ($AddItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}}
                if ($ItemsOnTheList -contains $AddItemFormFileNameInput.Text) {
                Show-MessageBox -Message "Файл с таким обозначением уже содежится в указанном списке." -Title "Невозможно выполнить действие" -Type OK
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
                    $AddItemForm.Close()
                }
            }
        }
    })
    $AddItemForm.Controls.Add($AddItemFormAddButton)
    
    #Кнопка закрыть
    $AddItemFormCancelButton = New-Object System.Windows.Forms.Button
    $AddItemFormCancelButton.Location = New-Object System.Drawing.Point(90,180) #x,y
    $AddItemFormCancelButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $AddItemFormCancelButton.Text = "Закрыть"
    $AddItemFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $AddItemFormCancelButton.Add_Click({
        $AddItemForm.Close()
    })
    $AddItemForm.Controls.Add($AddItemFormCancelButton)
    $AddItemForm.ActiveControl = $AddItemFormFileTypeLabel
    $AddItemForm.ShowDialog()
}

Function Edit-ItemOnList ()
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
    $EditItemFormFileNameInput.Text = "Укажите обозначение..."
    $EditItemFormFileNameInput.ForeColor = "Gray"
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
    $DataTypes | % {$EditItemFormFileTypeCombobox.Items.add($_)}
    $EditItemFormFileTypeCombobox.SelectedIndex = 0
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
    $EditItemFormAttributeValueInput.Text = "Укажите Изм. или MD5..."
    $EditItemFormAttributeValueInput.ForeColor = "Gray"
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
    
    #Кнопка добавить
    $EditItemFormAddButton = New-Object System.Windows.Forms.Button
    $EditItemFormAddButton.Location = New-Object System.Drawing.Point(10,109) #x,y
    $EditItemFormAddButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $EditItemFormAddButton.Text = "Добавить запись"
    $EditItemFormAddButton.Add_Click({
        if ($EditItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение и Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($EditItemFormFileNameInput.Text -eq "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -ne "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите обозначение файла." -Title "Значения указаны не во всех полях" -Type OK
        } elseif ($EditItemFormFileNameInput.Text -ne "Укажите обозначение..." -and $EditItemFormAttributeValueInput.Text -eq "Укажите Изм. или MD5...") {
            Show-MessageBox -Message "Пожалуйста, укажите Изм./MD5 для файла." -Title "Значения указаны не во всех полях" -Type OK
        } else {
                $ItemsOnTheList = @()
                if ($EditItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items | % {$ItemsOnTheList += $_.Text}}
                if ($EditItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items | % {$ItemsOnTheList += $_.Text}}
                if ($EditItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items | % {$ItemsOnTheList += $_.Text}}
                if ($ItemsOnTheList -contains $EditItemFormFileNameInput.Text) {
                Show-MessageBox -Message "Файл с таким обозначением уже содежится в указанном списке." -Title "Невозможно выполнить действие" -Type OK
            } else {
                if ($EditItemFormFileTypeCombobox.SelectedItem -eq "Документ" -and $EditItemFormAttributeValueInput.Text.Length -gt 5) {
                    if ((Show-MessageBox -Message "Указанный Изм. содержит больше пяти символов. Возможно вы ошибочно указали MD5 или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указан подозрительный Изм." -Type YesNo) -eq "Yes") {
                        $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                        $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                        if ($EditItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                        if ($EditItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                        if ($EditItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                        $EditItemForm.Close()
                    }
                } elseif ($EditItemFormFileTypeCombobox.SelectedItem -eq "Программа" -and $EditItemFormAttributeValueInput.Text.Length -ne 32) {
                    if ((Show-MessageBox -Message "Указанная сумма MD5 некорректна. Возможно вы непольностью указали ее, ошибочно указали Изм. или выбрали неверный тип файла.`r`nВсе равно продолжить?" -Title "Для файла указана подозрительная MD5 сумма" -Type YesNo) -eq "Yes") {
                        $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                        $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                        $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                        if ($EditItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                        if ($EditItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                        if ($EditItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                        Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                        $EditItemForm.Close()
                    }
                } else {
                    $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($EditItemFormFileNameInput.Text)")
                    $ItemToAdd.SubItems.Add("$($EditItemFormAttributeValueInput.Text)")
                    $ItemToAdd.SubItems.Add("$($EditItemFormFileTypeCombobox.SelectedItem)")
                    $ItemToAdd.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                    if ($EditItemFormRadioButtonAdd.Checked -eq $true) {$ListViewAdd.Items.Insert(0, $ItemToAdd)}
                    if ($EditItemFormRadioButtonReplace.Checked -eq $true) {$ListViewReplace.Items.Insert(0, $ItemToAdd)}
                    if ($EditItemFormRadioButtonRemove.Checked -eq $true) {$ListViewRemove.Items.Insert(0, $ItemToAdd)}
                    Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
                    $EditItemForm.Close()
                }
            }
        }
    })
    $EditItemForm.Controls.Add($EditItemFormAddButton)
    
    #Кнопка закрыть
    $EditItemFormCancelButton = New-Object System.Windows.Forms.Button
    $EditItemFormCancelButton.Location = New-Object System.Drawing.Point(90,109) #x,y
    $EditItemFormCancelButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $EditItemFormCancelButton.Text = "Закрыть"
    $EditItemFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $EditItemFormCancelButton.Add_Click({
    $EditItemForm.Close()
    })
    $EditItemForm.Controls.Add($EditItemFormCancelButton)
    $EditItemForm.ActiveControl = $EditItemFormFileTypeLabel
    $EditItemForm.ShowDialog()
}

Function Custom-Form ()
{
    Add-Type -AssemblyName System.Windows.Forms
    #FORM
    $ScriptMainWindow = New-Object System.Windows.Forms.Form
    $ScriptMainWindow.ShowIcon = $false
    $ScriptMainWindow.AutoSize = $true
    $ScriptMainWindow.Text = "Генерация ИИ"
    $ScriptMainWindow.AutoSizeMode = "GrowAndShrink"
    $ScriptMainWindow.WindowState = "Normal"
    $ScriptMainWindow.SizeGripStyle = "Hide"
    $ScriptMainWindow.ShowInTaskbar = $true
    $ScriptMainWindow.StartPosition = "CenterScreen"
    $ScriptMainWindow.MinimizeBox = $true
    $ScriptMainWindow.MaximizeBox = $false
    $ScriptMainWindow.Padding = New-Object System.Windows.Forms.Padding(0,0,10,0)
    
    #Groupbox 'Настройка списков'
    $ListSettingsGroup = New-Object System.Windows.Forms.GroupBox
    $ListSettingsGroup.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ListSettingsGroup.Size = New-Object System.Drawing.Point(1308,800) #width,height
    $ListSettingsGroup.Text = "Настройка списков"
    $ScriptMainWindow.Controls.Add($ListSettingsGroup)
    
    #Надпись к списку Выпустить
    $ListViewAddLabel = New-Object System.Windows.Forms.Label
    $ListViewAddLabel.Location =  New-Object System.Drawing.Point(10,25) #x,y
    $ListViewAddLabel.Width = 200
    $ListViewAddLabel.Height = 15
    $ListViewAddLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewAddLabel.Text = "Выпустить (0):"
    $ListViewAddLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewAddLabel)

    
    #Список Выпустить
    $ListViewAdd = New-Object System.Windows.Forms.ListView
    $ListViewAdd.Location = New-Object System.Drawing.Point(10,42) #x, y
    $ListViewAdd.View = "Details"
    $ListViewAdd.FullRowSelect = $true
    $ListViewAdd.MultiSelect = $false
    $ListViewAdd.HideSelection = $false
    $ListViewAdd.Width = 400
    $ListViewAdd.Height = 370
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Обозначение" -Width 267
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Изм./MD5" -Width 69
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Тип" -Width 43
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
    $ButtonMoveToLeftBetweenAddAndReplace.Location = New-Object System.Drawing.Point(420,50) #x,y
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
    $ButtonMoveToRightBetweenAddAndReplace.Location = New-Object System.Drawing.Point(420,90) #x,y
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
    $ListViewReplaceLabel.Location =  New-Object System.Drawing.Point(454,25) #x,y
    $ListViewReplaceLabel.Width = 200
    $ListViewReplaceLabel.Height = 15
    $ListViewReplaceLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewReplaceLabel.Text = "Заменить (0):"
    $ListViewReplaceLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewReplaceLabel)

    #Список Заменить
    $ListViewReplace = New-Object System.Windows.Forms.ListView
    $ListViewReplace.Location = New-Object System.Drawing.Point(454,42) #x, y
    $ListViewReplace.View = "Details"
    $ListViewReplace.FullRowSelect = $true
    $ListViewReplace.MultiSelect = $false
    $ListViewReplace.HideSelection = $false
    $ListViewReplace.Width = 400
    $ListViewReplace.Height = 370
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Обозначение" -Width 267
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Изм./MD5" -Width 69
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Тип" -Width 43
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
    $ButtonMoveToRightBetweenReplaceAndRemove.Location = New-Object System.Drawing.Point(864,50) #x,y
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
    $ButtonMoveToLeftBetweenReplaceAndRemove.Location = New-Object System.Drawing.Point(864,90) #x,y
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
    $ListViewRemoveLabel.Location =  New-Object System.Drawing.Point(898,25) #x,y
    $ListViewRemoveLabel.Width = 200
    $ListViewRemoveLabel.Height = 15
    $ListViewRemoveLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
    $ListViewRemoveLabel.Text = "Аннулировать (0):"
    $ListViewRemoveLabel.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListViewRemoveLabel)

    #Список Аннулировать
    $ListViewRemove = New-Object System.Windows.Forms.ListView
    $ListViewRemove.Location = New-Object System.Drawing.Point(898,42) #x, y
    $ListViewRemove.View = "Details"
    $ListViewRemove.FullRowSelect = $true
    $ListViewRemove.MultiSelect = $false
    $ListViewRemove.HideSelection = $false
    $ListViewRemove.Width = 400
    $ListViewRemove.Height = 370
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Обозначение" -Width 267
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Изм./MD5" -Width 69
    Add-HeaderToViewList -ListView $ListViewRemove -HeaderText "Тип" -Width 43
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
    $ListSettingsGroupTotalEntries.Location =  New-Object System.Drawing.Point(10,420) #x,y
    $ListSettingsGroupTotalEntries.Width = 200
    $ListSettingsGroupTotalEntries.Height = 15
    $ListSettingsGroupTotalEntries.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Regular)
    $ListSettingsGroupTotalEntries.Text = "Всего файлов в списках: 0"
    $ListSettingsGroupTotalEntries.TextAlign = "TopLeft"
    $ListSettingsGroup.Controls.Add($ListSettingsGroupTotalEntries)

    #Группа элементов Выбранный файл
    $ListSettingsSelectedItem = New-Object System.Windows.Forms.GroupBox
    $ListSettingsSelectedItem.Location = New-Object System.Drawing.Point(10,440) #x,y
    $ListSettingsSelectedItem.Size = New-Object System.Drawing.Point(400,100) #width,height
    $ListSettingsSelectedItem.Text = "Выбранный файл"
    $ListSettingsGroup.Controls.Add($ListSettingsSelectedItem)
    #Поле Обозначение для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileName = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileName.Location =  New-Object System.Drawing.Point(15,20) #x,y
    $ListSettingsSelectedItemFileName.Width = 380
    $ListSettingsSelectedItemFileName.Height = 15
    $ListSettingsSelectedItemFileName.Text = "Обозначение:"
    $ListSettingsSelectedItemFileName.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileName)
    #Поле Изм./MD5 для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileAttribute = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileAttribute.Location =  New-Object System.Drawing.Point(15,45) #x,y
    $ListSettingsSelectedItemFileAttribute.Width = 380
    $ListSettingsSelectedItemFileAttribute.Height = 15
    $ListSettingsSelectedItemFileAttribute.Text = "Изм./MD5:"
    $ListSettingsSelectedItemFileAttribute.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileAttribute)
    #Поле Тип файла для группы элементов Выбранный файл
    $ListSettingsSelectedItemFileType = New-Object System.Windows.Forms.Label
    $ListSettingsSelectedItemFileType.Location =  New-Object System.Drawing.Point(15,70) #x,y
    $ListSettingsSelectedItemFileType.Width = 380
    $ListSettingsSelectedItemFileType.Height = 15
    $ListSettingsSelectedItemFileType.Text = "Тип файла:"
    $ListSettingsSelectedItemFileType.TextAlign = "TopLeft"
    $ListSettingsSelectedItem.Controls.Add($ListSettingsSelectedItemFileType)
    
    #ADD ITEMS TO 'LIST SETTINGS' GROUP
    $ButtonAddItem = New-Object System.Windows.Forms.Button
    $ButtonAddItem.Location = New-Object System.Drawing.Point(10,770) #x,y
    $ButtonAddItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonAddItem.Text = "Добавить запись"
    $ButtonAddItem.Add_Click({Add-ItemToList})
    $ListSettingsGroup.Controls.Add($ButtonAddItem)

    #Button 'Delete'
    $ButtonDeleteItem = New-Object System.Windows.Forms.Button
    $ButtonDeleteItem.Location = New-Object System.Drawing.Point(130,770) #x,y
    $ButtonDeleteItem.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonDeleteItem.Text = "Удалить запись"
    $ButtonDeleteItem.Add_Click({
    Write-Host $ListViewAdd.SelectedItems.Count
    Write-Host $ListViewReplace.SelectedItems.Count
    Write-Host $ListViewRemove.SelectedItems.Count
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].Remove()}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].Remove()}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].Remove()}
    $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
    $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false 
    Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
    })
    $ListSettingsGroup.Controls.Add($ButtonDeleteItem)

    #Выбрать цвет
    $ButtonSelectColor = New-Object System.Windows.Forms.Button
    $ButtonSelectColor.Location = New-Object System.Drawing.Point(370,770) #x,y
    $ButtonSelectColor.Size = New-Object System.Drawing.Point(22,22) #width,height
    $ColorDialog = New-Object System.Windows.Forms.ColorDialog
    $ButtonSelectColor.BackColor = [System.Drawing.Color]::LightGreen
    $ColorDialog.Color = [System.Drawing.Color]::LightGreen
    $ButtonSelectColor.Add_Click({
    $ColorDialog.ShowDialog()
    $ButtonSelectColor.BackColor = $ColorDialog.Color
    })
    $ListSettingsGroup.Controls.Add($ButtonSelectColor)

    #Выделить цветом
    $ButtonMarkWithColor = New-Object System.Windows.Forms.Button
    $ButtonMarkWithColor.Location = New-Object System.Drawing.Point(490,770) #x,y
    $ButtonMarkWithColor.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonMarkWithColor.Text = "Выделить цветом"
    $ButtonMarkWithColor.Add_Click({
    Write-Host $ColorDialog.Color
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].BackColor = $ColorDialog.Color}
    Unselect-ItemsInOtherLists -List3 $ListViewAdd -List1 $ListViewReplace -List2 $ListViewRemove 
    })
    $ListSettingsGroup.Controls.Add($ButtonMarkWithColor)

    #Заполнить списки
    $ButtonPopulateLists = New-Object System.Windows.Forms.Button
    $ButtonPopulateLists.Location = New-Object System.Drawing.Point(250,770) #x,y
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
    $ListSettingsGroup.Controls.Add($ButtonPopulateLists)

    #Редактировать запись
    $ButtonEditItemOnList = New-Object System.Windows.Forms.Button
    $ButtonEditItemOnList.Location = New-Object System.Drawing.Point(610,770) #x,y
    $ButtonEditItemOnList.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ButtonEditItemOnList.Text = "Изменить запись"
    $ButtonEditItemOnList.Add_Click({
    Edit-ItemOnList
    })
    $ListSettingsGroup.Controls.Add($ButtonEditItemOnList)

    
    <#
    #Button 'Edit'
    $CreatePropertiesPageButtonEditItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonEditItem.Location = New-Object System.Drawing.Point(98,230) #x,y
    $CreatePropertiesPageButtonEditItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonEditItem.Text = "Edit item"
    $CreatePropertiesPageButtonEditItem.Add_Click({
    if ($CreatePropertyListView.SelectedIndices.Count -gt 0) {AddApply-ItemToList -FormName "Edit item" -ButtonName "Apply" -Type Apply} else {write-host "FUCK YOU!"} 
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonEditItem)
    #Button 'Clear'
    $CreatePropertiesPageButtonClearList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonClearList.Location = New-Object System.Drawing.Point(274,230) #x,y
    $CreatePropertiesPageButtonClearList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonClearList.Text = "Clear list"
    $CreatePropertiesPageButtonClearList.Add_Click({
    $CreatePropertyListView.Items.Clear()
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonClearList)
    #Button 'Export'
    $CreatePropertiesPageButtonExportList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonExportList.Location = New-Object System.Drawing.Point(362,230) #x,y
    $CreatePropertiesPageButtonExportList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonExportList.Text = "Export list"
    $CreatePropertiesPageButtonExportList.Add_Click({
    #$CreatePropertyListView.Items[0].SubItems | % {Write-Host $_.Text}
    Write-Host $CreatePropertyListView.Items.Count
    $CreatePropertyListView.Items | % {$_.SubItems | % {Write-Host $_.Text}}
    $PathToExportedList = Save-File
            if ($PathToExportedList -ne $null) {
            if (Test-Path -Path $PathToExportedList) {Remove-Item -Path $PathToExportedList -Force}
            $CreatePropertyListView.Items | % {$ExportedString = ""; $_.SubItems | % {$ExportedString += "$($_.Text);"}; $ExportedString = $ExportedString.TrimEnd(";"); Add-Content -Path $PathToExportedList -Value $ExportedString}
        }
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonExportList)
    #Button 'Import'
    $CreatePropertiesPageButtonImportList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonImportList.Location = New-Object System.Drawing.Point(450,230) #x,y
    $CreatePropertiesPageButtonImportList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonImportList.Text = "Import list"
    $CreatePropertiesPageButtonImportList.Add_Click({})
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonImportList)
    $CreatePropertiesPage.Controls.Add($CreatePropertyTest)#>
    $ScriptMainWindow.ShowDialog()
}
Custom-Form
