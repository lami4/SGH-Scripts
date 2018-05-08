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
    $DataTypes | % {$EditItemFormFileTypeCombobox.Items.add($_)}
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
                $ItemsOnTheList = @()
                $ListObject.Items | % {$ItemsOnTheList += $_.Text}
                if ($ItemsOnTheList -contains $EditItemFormFileNameInput.Text -and $EditItemFormFileNameInput.Text -ne $ListObject.Items[$ListObject.SelectedIndices[0]].Text) {
                Show-MessageBox -Message "Файл с таким обозначением уже содежится в списке." -Title "Невозможно выполнить действие" -Type OK
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
    $ManageCustomListsCloseButton.Enabled = $BooleanRest
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

Function Generate-XmlList ($List, [ValidateSet("Departments", "Employees")]$ListType)
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
    $XmlList.AppendChild($RootElement) | Out-Null
    Foreach ($ListItem in $List) {
        $ElementName = $XmlList.CreateNode("element","name",$null)
        $ElementName.InnerText = $ListItem
        if ($ListType -eq "Departments") {$XmlList.SelectSingleNode("/departments").AppendChild($ElementName)}
        if ($ListType -eq "Employees") {$XmlList.SelectSingleNode("/employees").AppendChild($ElementName)}
    }
    if ($ListType -eq "Departments") {$XmlList.Save("$PSScriptRoot\Отделы.xml")}
    if ($ListType -eq "Employees") {$XmlList.Save("$PSScriptRoot\Сотрудники.xml")}
}

Function Manage-CustomLists ($PathToLost, [ValidateSet("Departments", "Employees")]$ListType)
{
    $ManageCustomListsForm = New-Object System.Windows.Forms.Form
    $ManageCustomListsForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ManageCustomListsForm.ShowIcon = $false
    $ManageCustomListsForm.AutoSize = $true
    if ($ListType -eq "Departments") {$ManageCustomListsForm.Text = "Редактировать список отделов"}
    if ($ListType -eq "Employees") {$ManageCustomListsForm.Text = "Редактировать список сотрудников"}
    $ManageCustomListsForm.AutoSizeMode = "GrowAndShrink"
    $ManageCustomListsForm.WindowState = "Normal"
    $ManageCustomListsForm.SizeGripStyle = "Hide"
    $ManageCustomListsForm.ShowInTaskbar = $true
    $ManageCustomListsForm.StartPosition = "CenterScreen"
    $ManageCustomListsForm.MinimizeBox = $false
    $ManageCustomListsForm.MaximizeBox = $false
    #$ManageCustomListsForm.Size = New-Object System.Drawing.Point(550,300) #width,height 
    #Надпись к списку, который содержит список отделов/сотрудников компании
    $GetPropertyLabelBlacklistListBox = New-Object System.Windows.Forms.Label
    $GetPropertyLabelBlacklistListBox.Location =  New-Object System.Drawing.Point(10,10) #x,y
    $GetPropertyLabelBlacklistListBox.Width = 250
    $GetPropertyLabelBlacklistListBox.Height = 13
    if ($ListType -eq "Departments") {$GetPropertyLabelBlacklistListBox.Text = "Список отделов компании:"}
    if ($ListType -eq "Employees") {$GetPropertyLabelBlacklistListBox.Text = "Список сотрудников компании:"}
    $ManageCustomListsForm.Controls.Add($GetPropertyLabelBlacklistListBox)
    #Список отделов/сотрудников компании
    #$DefaultBlackList = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(10,25) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,260) #width,height
    #$DefaultBlackList | % {$GetPropertyListBoxBlackList.Items.Add($_)} | Out-Null
    if ($ListType -eq "Departments") {if (Test-Path "$PSScriptRoot\Отделы.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Отделы.xml"}}
    if ($ListType -eq "Employees") {if (Test-Path "$PSScriptRoot\Сотрудники.xml") {Populate-List -List $GetPropertyListBoxBlackList -PathToXml "$PSScriptRoot\Сотрудники.xml"}}
    $GetPropertyListBoxBlackList.Add_SelectedIndexChanged({
        if ($GetPropertyListBoxBlackList.SelectedIndex -ne -1) {
            #Write-Host "$($GetPropertyListBoxBlackList.SelectedIndex)"
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
    
    $GetPropertyInputboxAddItem.ForeColor = "Gray"
    $GetPropertyInputboxAddItem.Add_GotFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "Укажите название отдела..." -or $GetPropertyInputboxAddItem.Text -eq "Укажите ФИО...") {
            $GetPropertyInputboxAddItem.Text = ""
            $GetPropertyInputboxAddItem.ForeColor = "Black"
        }
        })
    $GetPropertyInputboxAddItem.Add_LostFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "") {
            if ($ListType -eq "Departments") {$GetPropertyInputboxAddItem.Text = "Укажите название отдела..."}
            if ($ListType -eq "Employees") {$GetPropertyInputboxAddItem.Text = "Укажите ФИО..."}
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
            } else {
                $SelectedIndex = $GetPropertyListBoxBlackList.SelectedIndex
                $GetPropertyListBoxBlackList.Items.Insert($GetPropertyListBoxBlackList.SelectedIndex, ($GetPropertyInputboxEditItem.Text).Trim(' '))
                $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
                $GetPropertyListBoxBlackList.SelectedIndex = $SelectedIndex
                $GetPropertyInputboxEditItem.Text = ($GetPropertyInputboxEditItem.Text).Trim(' ')
                Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
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
        $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
        $GetPropertyInputboxEditItem.Text = "Выберите запись из списка..."
        })
    $ManageCustomListsForm.Controls.Add($GetPropertyButtonDeleteItem)
    #Надпись для кнопки удалить
    $GetPropertyLabelButtonDelete = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonDelete.Location =  New-Object System.Drawing.Point(350,107) #x,y
    $GetPropertyLabelButtonDelete.Size =  New-Object System.Drawing.Point(180,15) #width,height
    if ($ListType -eq "Departments") {$GetPropertyLabelButtonDelete.Text = "Удалить отдел из списка"}
    if ($ListType -eq "Employees") {$GetPropertyLabelButtonDelete.Text = "Удалить сотрудника из списка"}
    
    $ManageCustomListsForm.Controls.Add($GetPropertyLabelButtonDelete)
    #Кнопка очистить
    $GetPropertyButtonClear = New-Object System.Windows.Forms.Button
    $GetPropertyButtonClear.Location = New-Object System.Drawing.Point(235,132) #x,y
    $GetPropertyButtonClear.Size = New-Object System.Drawing.Point(110,22) #width,height
    $GetPropertyButtonClear.Text = "Очистить список"
    $GetPropertyButtonClear.Add_Click({
        $ClickResult = Show-MessageBox -Title "Подтвердите действие" -Type YesNo -Message "Вы уверены, что хотите удалить все записи из списка?"
        if ($ClickResult -eq "Yes") {$GetPropertyListBoxBlackList.Items.Clear()}
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
    $ManageCustomListsSaveButton.Text = "Сохранить"
    $ManageCustomListsSaveButton.Add_Click({
    if ($ListType -eq "Departments") {
    Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Departments
        $ComboboxDepartmentName.Items.Clear()
        Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
            $ComboboxDepartmentName.Items.Add($ItemInList)
        }
    }
    if ($ListType -eq "Employees") {
    Generate-XmlList -List $GetPropertyListBoxBlackList.Items -ListType Employees
        $ComboboxCheckedBy.Items.Clear()
        $ComboboxCreatedBy.Items.Clear()
        Foreach ($ItemInList in $GetPropertyListBoxBlackList.Items) {
            $ComboboxCheckedBy.Items.Add($ItemInList)
            $ComboboxCreatedBy.Items.Add($ItemInList)
        }
    }
    $ManageCustomListsForm.Close()
    })
    $ManageCustomListsForm.Controls.Add($ManageCustomListsSaveButton)
    #Кнопка Закрыть
    $ManageCustomListsCloseButton = New-Object System.Windows.Forms.Button
    $ManageCustomListsCloseButton.Location = New-Object System.Drawing.Point(350,255) #x,y
    $ManageCustomListsCloseButton.Size = New-Object System.Drawing.Point(110,22) #width,height
    $ManageCustomListsCloseButton.Text = "Закрыть"
    $ManageCustomListsCloseButton.Add_Click({$ManageCustomListsForm.Close()})
    $ManageCustomListsForm.Controls.Add($ManageCustomListsCloseButton)
    $ManageCustomListsForm.ShowDialog()
}

Function Apply-FormattingInListTable($TableObject, $WordApp)
{
    #Сделать границы таблицы видимыми
    $TableObject.Borders.Enable = $true
    #Ширина таблицы (19,4 см)
    $TableObject.Columns.Item(1).Width = $WordApp.CentimetersToPoints(18)
    #Отсутуп от левого края в ячейках
    $TableObject.LeftPadding = $WordApp.CentimetersToPoints(0.1)
    #Отсутуп от правого края в ячейках
    $TableObject.RightPadding = $WordApp.CentimetersToPoints(0.1)
    #Установить вертикальное выравнивание по центру для всех ячеек
    $TableObject.Cell(1, 1).VerticalAlignment = 1
    #Настройка шрифта в таблице
    $TableObject.Cell(1, 1).Range.Font.Name = "Arial"
    $TableObject.Cell(1, 1).Range.Font.Size = 9
    #Интервал после (0 пт) для каждой ячейки таблицы
    $TableObject.Range.ParagraphFormat.SpaceAfter = 0
    #Междустрочный интервал (одинарный) для каждой ячейки таблицы
    $TableObject.Range.ParagraphFormat.LineSpacingRule = 0
    #Установить высоту на минимум
    $TableObject.Cell(15, 1).HeightRule = 1
    $TableObject.Rows.Height = $word.CentimetersToPoints(0.5)
    #Разбить таблицу на 4 столбца и установить их ширину
    $TableObject.Cell(1, 1).Split(1, 4)
    $TableObject.Cell(1, 1).Column.Width = $WordApp.CentimetersToPoints(1.5)
    $TableObject.Cell(1, 2).Column.Width = $WordApp.CentimetersToPoints(1.5)
    $TableObject.Cell(1, 3).Column.Width = $WordApp.CentimetersToPoints(7.5)
    $TableObject.Cell(1, 4).Column.Width = $WordApp.CentimetersToPoints(7.5)
    #Добавить надписи в таблицу
    $TableObject.Cell(1, 1).Range.Text = "Поз."
    $TableObject.Cell(1, 2).Range.Text = "Изм."
    $TableObject.Cell(1, 3).Range.Text = "Обозначение"
    $TableObject.Cell(1, 4).Range.Text = "Примечание"
    #Установить выравниевание по центру для заголовка
    $TableObject.Rows.Item(1).Range.ParagraphFormat.Alignment = 1
    #Выровнять таблицу по центру в документу
    $TableObject.Rows.Alignment = 1
    #Добавить строку для входных данных
    $TableObject.Rows.Add()
    #Запретить переносить ячейку на селдующую страницу
    $TableObject.Rows.Item(2).AllowBreakAcrossPages = $false
    #Отформатировать строку для входных данных
    $TableObject.Cell(2, 3).Range.ParagraphFormat.Alignment = 0
    $TableObject.Cell(2, 2).LeftPadding = $word.CentimetersToPoints(0.1)
    $TableObject.Cell(2, 2).RightPadding = $word.CentimetersToPoints(0.1)
}

Function Generate-UpdateNotification ($NotificationName)
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
Start-Sleep -Seconds 5
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
$table.Cell(1, 2).Range.Text = "БТД"
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

for ($i = 0; $i -lt 29; $i++) {
$table.Cell(15, 1).Delete()
}


#Отключить фиксированную высоту для ячейки, которая содержит перечень фдокументов и программ
$table.Cell(15, 1).HeightRule = 1
#Вставить таблицы со списками для аннлирования, замены и публикации
$document.Tables.Item(1).Cell(15, 1).TopPadding = $word.CentimetersToPoints(0.2)
$document.Tables.Item(1).Cell(15, 1).BottomPadding = $word.CentimetersToPoints(0.2)
$document.Tables.Item(1).Cell(15, 1).Range = [char]10 + [char]10 + "Заменить:" + [char]10 + [char]10 + [char]10 + [char]10 + "Аннулировать:" + [char]10 + [char]10 + [char]10 + [char]10 + "Выпустить:" + [char]10 + [char]10 + [char]10  + [char]10
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(4).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(9).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.Paragraphs.Item(14).Range.Select()
$document.Tables.Add($word.Selection.Range, 1, 1)
$document.Tables.Item(1).Cell(15, 1).Range.ParagraphFormat.Alignment = 1
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(1) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(2) -WordApp $word
Apply-FormattingInListTable -TableObject $document.Tables.Item(1).Tables.Item(3) -WordApp $word
Start-Sleep -Seconds 2
$document.SaveAs([ref]"$PSScriptRoot\$NotificationName.docx")
Start-Sleep -Seconds 2
$document.Close()
Start-Sleep -Seconds 2
$word.Quit()
Start-Sleep -Seconds 5
}

Function Add-Data ($TableObject, $List)
{
    $RowCounter = 2
    Foreach ($Item in $List.Items) {
        $TableObject.Cell($RowCounter, 1).Range.Text = $script:CounterForWordItems
        $script:CounterForWordItems += 1
        if ($Item.Subitems[2].Text -eq "Программа") {
            $TableObject.Cell($RowCounter, 2).Range.Text = "-"
            $TableObject.Cell($RowCounter, 3).Range.Text = $Item.Text
            $TableObject.Cell($RowCounter, 4).Range.Text = $Item.Subitems[1].Text
        } else {
            $TableObject.Cell($RowCounter, 2).Range.Text = $Item.Subitems[1].Text
            $TableObject.Cell($RowCounter, 3).Range.Text = $Item.Text
        }
        $TableObject.Rows.Add()
        $RowCounter += 1
    }
}

Function Move-ListsToWordDocument ($NotificationName)
{
$script:counter = 1
#Создать экземпляр приложения MS Word
$WordToPopulate = New-Object -ComObject Word.Application
#Создать документ MS Word
$DocumentToPopulate = $WordToPopulate.Documents.Open("$PSScriptRoot\$NotificationName.docx")
#Сделать вызванное приложение невидемым
$WordToPopulate.Visible = $false
Start-Sleep -Seconds 5
Write-Host "Заполняю таблицу Выпустить..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(1).Tables.Item(1) -List $ListViewAdd
Write-Host "Заполняю таблицу Заменитить..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(1).Tables.Item(2) -List $ListViewReplace
Write-Host "Заполняю таблицу Аннулировать..."
Add-Data -TableObject $DocumentToPopulate.Tables.Item(1).Tables.Item(3) -List $ListViewRemove
Write-Host "Обновляю поля в Word документе..."
#Вставить поля
$HeaderTablePopulate = $DocumentToPopulate.Sections.Item(1).Headers.Item(1).Range.Tables.Item(1)
$TablePopulate = $DocumentToPopulate.Tables.Item(1)
$TablePopulate.Cell(5, 4).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 26)
$TablePopulate.Cell(5, 4).Range.ParagraphFormat.Alignment = 1

$HeaderTablePopulate.Cell(2, 4).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 33)

$TablePopulate.Cell(5, 3).Range.Select()
$DocumentToPopulate.Application.Selection.Collapse(1)
$myField = $DocumentToPopulate.Fields.Add($DocumentToPopulate.Application.Selection.Range, 33)
$TablePopulate.Cell(5, 3).Range.ParagraphFormat.Alignment = 1
Start-Sleep -Seconds 2
$DocumentToPopulate.Save()
Start-Sleep -Seconds 2
$DocumentToPopulate.Close()
Start-Sleep -Seconds 2
$WordToPopulate.Quit()
Start-Sleep -Seconds 2
Kill -Name WINWORD -ErrorAction SilentlyContinue
Write-Host "Извешение сгенерировано. Скрипт закончил работу."
}

Function Custom-Form ()
{
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    #Главное окно
    $ScriptMainWindow = New-Object System.Windows.Forms.Form
    $ScriptMainWindow.ShowIcon = $false
    $ScriptMainWindow.AutoSize = $true
    $ScriptMainWindow.Text = "Генерация ИИ"
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

    #Группа элементов Выбранный файл
    $ListSettingsSelectedItem = New-Object System.Windows.Forms.GroupBox
    $ListSettingsSelectedItem.Location = New-Object System.Drawing.Point(10,445) #x,y
    $ListSettingsSelectedItem.Size = New-Object System.Drawing.Point(513,100) #width,height
    $ListSettingsSelectedItem.Text = "Выбранный файл"
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
    if ($ListViewAdd.SelectedIndices.Count -gt 0) {$ListViewAdd.Items[$ListViewAdd.SelectedIndices[0]].Remove()}
    if ($ListViewReplace.SelectedIndices.Count -gt 0) {$ListViewReplace.Items[$ListViewReplace.SelectedIndices[0]].Remove()}
    if ($ListViewRemove.SelectedIndices.Count -gt 0) {$ListViewRemove.Items[$ListViewRemove.SelectedIndices[0]].Remove()}
    $ButtonMoveToLeftBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToRightBetweenAddAndReplace.Enabled = $false
    $ButtonMoveToLeftBetweenReplaceAndRemove.Enabled = $false
    $ButtonMoveToRightBetweenReplaceAndRemove.Enabled = $false 
    Update-ListCounters -AddListCounter $ListViewAddLabel -AddList $ListViewAdd -ReplaceListCounter $ListViewReplaceLabel -ReplaceList $ListViewReplace -RemoveListCounter $ListViewRemoveLabel -RemoveList $ListViewRemove -TotalEntriesCounter $ListSettingsGroupTotalEntries
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
    #Заполнить списки
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
    $ButtonClearLists.Add_Click({Clear-Lists})
    $ListSettingsListActions.Controls.Add($ButtonClearLists)
    #Отменить выделение
    $ButtonRemoveColoringLists = New-Object System.Windows.Forms.Button
    $ButtonRemoveColoringLists.Location = New-Object System.Drawing.Point(10,43) #x,y
    $ButtonRemoveColoringLists.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonRemoveColoringLists.Text = "Отменить выделение..."
    $ButtonRemoveColoringLists.Add_Click({Discard-Coloring})
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
    $ButtonImportFromXml.Text = "Импорт из XML..."
    $ButtonImportFromXml.Add_Click({})
    $ListSettingsListActions.Controls.Add($ButtonImportFromXml)
    #Экспортировать в XML
    $ButtonExportToXml = New-Object System.Windows.Forms.Button
    $ButtonExportToXml.Location = New-Object System.Drawing.Point(167,43) #x,y
    $ButtonExportToXml.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonExportToXml.Text = "Экспорт в XML..."
    $ButtonExportToXml.Add_Click({})
    $ListSettingsListActions.Controls.Add($ButtonExportToXml)
    #Пакетный импорт файлов
    $ButtonBatchFileImport = New-Object System.Windows.Forms.Button
    $ButtonBatchFileImport.Location = New-Object System.Drawing.Point(167,69) #x,y
    $ButtonBatchFileImport.Size = New-Object System.Drawing.Point(137,22) #width,height
    $ButtonBatchFileImport.Text = "Пакетный импорт..."
    $ButtonBatchFileImport.Add_Click({})
    $ListSettingsListActions.Controls.Add($ButtonBatchFileImport)

    #Группа элементов Параметры извещения
    $UpdateNotificationParameters = New-Object System.Windows.Forms.GroupBox
    $UpdateNotificationParameters.Location = New-Object System.Drawing.Point(10,575) #x,y
    $UpdateNotificationParameters.Size = New-Object System.Drawing.Point(659,115) #width,height
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
    #$UpdateNotificationNumberInput.Text = "11-22-3333"
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
    #$CalendarApplyUpdatesUntilInput.Text = "03.02.1990"
    $UpdateNotificationParameters.Controls.Add($CalendarApplyUpdatesUntilInput)
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
    $ListOfNames = @("С. Селюто","В. Горемыкин","И. Чижиков")
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
    if ($ErrorPresent -eq $true) {
        Show-MessageBox -Message $TextInMessage -Title "Невозможно начать генерацию ИИ" -Type OK
    } else {
        $script:CounterForWordItems = 1
        Generate-UpdateNotification -NotificationName $UpdateNotificationNumberInput.Text
        Move-ListsToWordDocument -NotificationName $UpdateNotificationNumberInput.Text
    }
    })
    
    $ScriptMainWindow.Controls.Add($ButtonRun)

    
    <#
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

Custom-Form | Out-Null
