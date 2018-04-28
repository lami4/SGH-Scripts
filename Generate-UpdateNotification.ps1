clear
Function Custom-Form
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
    
    #Groupbox 'Настройка списков'
    $ListSettingsGroup = New-Object System.Windows.Forms.GroupBox
    $ListSettingsGroup.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ListSettingsGroup.Size = New-Object System.Drawing.Point(1500,267) #width,height
    $ListSettingsGroup.Text = "Настройка списков"
    $ScriptMainWindow.Controls.Add($ListSettingsGroup)
    
    #Table Add
    $ListViewAdd = New-Object System.Windows.Forms.ListView
    $ListViewAdd.Location = New-Object System.Drawing.Point(10,20) #x, y
    $ListViewAdd.View = "Details"
    $ListViewAdd.FullRowSelect = $true
    $ListViewAdd.MultiSelect = $false
    $ListViewAdd.HideSelection = $false
    $ListViewAdd.Width = 379
    $ListViewAdd.Height = 200
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Обозначение" -Width 125
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Изм./MD5" -Width 125
    Add-HeaderToViewList -ListView $ListViewAdd -HeaderText "Тип" -Width 125
    $ListViewAdd_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewAdd.Columns[0].Width -ne 125) {
            $ListViewAdd.Columns[0].Width = 125
        }
        if ($ListViewAdd.Columns[1].Width -ne 125) {
            $ListViewAdd.Columns[1].Width = 125
        }
        if ($ListViewAdd.Columns[2].Width -ne 125) {
            $ListViewAdd.Columns[2].Width = 125
        }
    }
    $ListViewAdd.add_ColumnWidthChanged($ListViewAdd_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewAdd)
    
    #Table Replace
    $ListViewReplace = New-Object System.Windows.Forms.ListView
    $ListViewReplace.Location = New-Object System.Drawing.Point(420,20) #x, y
    $ListViewReplace.View = "Details"
    $ListViewReplace.FullRowSelect = $true
    $ListViewReplace.MultiSelect = $false
    $ListViewReplace.HideSelection = $false
    $ListViewReplace.Width = 379
    $ListViewReplace.Height = 200
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Обозначение" -Width 125
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Изм./MD5" -Width 125
    Add-HeaderToViewList -ListView $ListViewReplace -HeaderText "Тип" -Width 125
    $ListViewReplace_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewReplace.Columns[0].Width -ne 125) {
            $ListViewReplace.Columns[0].Width = 125
        }
        if ($ListViewReplace.Columns[1].Width -ne 125) {
            $ListViewReplace.Columns[1].Width = 125
        }
        if ($ListViewReplace.Columns[2].Width -ne 125) {
            $ListViewReplace.Columns[2].Width = 125
        }
    }
    $ListViewReplace.add_ColumnWidthChanged($ListViewReplace_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewReplace)
    
    #Table Cancel
    $ListViewCancel = New-Object System.Windows.Forms.ListView
    $ListViewCancel.Location = New-Object System.Drawing.Point(830,20) #x, y
    $ListViewCancel.View = "Details"
    $ListViewCancel.FullRowSelect = $true
    $ListViewCancel.MultiSelect = $false
    $ListViewCancel.HideSelection = $false
    $ListViewCancel.Width = 379
    $ListViewCancel.Height = 200
    Add-HeaderToViewList -ListView $ListViewCancel -HeaderText "Обозначение" -Width 125
    Add-HeaderToViewList -ListView $ListViewCancel -HeaderText "Изм./MD5" -Width 125
    Add-HeaderToViewList -ListView $ListViewCancel -HeaderText "Тип" -Width 125
        $ListViewCancel_ColumnWidthChanged = [System.Windows.Forms.ColumnWidthChangedEventHandler]{
        if ($ListViewCancel.Columns[0].Width -ne 125) {
            $ListViewCancel.Columns[0].Width = 125
        }
        if ($ListViewCancel.Columns[1].Width -ne 125) {
            $ListViewCancel.Columns[1].Width = 125
        }
        if ($ListViewCancel.Columns[2].Width -ne 125) {
            $ListViewCancel.Columns[2].Width = 125
        }
    }
    $ListViewCancel.add_ColumnWidthChanged($ListViewCancel_ColumnWidthChanged)
    $ListSettingsGroup.Controls.Add($ListViewCancel)
    #ADD ITEMS TO 'LIST SETTINGS' GROUP
    
    
    <##Button 'Add'
    $CreatePropertiesPageButtonAddItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonAddItem.Location = New-Object System.Drawing.Point(10,230) #x,y
    $CreatePropertiesPageButtonAddItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonAddItem.Text = "Add item"
    $CreatePropertiesPageButtonAddItem.Add_Click({
    AddApply-ItemToList -FormName "Add item" -ButtonName "Add" -Type Add
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonAddItem)
    #Button 'Edit'
    $CreatePropertiesPageButtonEditItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonEditItem.Location = New-Object System.Drawing.Point(98,230) #x,y
    $CreatePropertiesPageButtonEditItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonEditItem.Text = "Edit item"
    $CreatePropertiesPageButtonEditItem.Add_Click({
    if ($CreatePropertyListView.SelectedIndices.Count -gt 0) {AddApply-ItemToList -FormName "Edit item" -ButtonName "Apply" -Type Apply} else {write-host "FUCK YOU!"} 
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonEditItem)
    #Button 'Delete'
    $CreatePropertiesPageButtonDeleteItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonDeleteItem.Location = New-Object System.Drawing.Point(186,230) #x,y
    $CreatePropertiesPageButtonDeleteItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonDeleteItem.Text = "Delete item"
    $CreatePropertiesPageButtonDeleteItem.Add_Click({
    if ($CreatePropertyListView.SelectedIndices.Count -gt 0) {$CreatePropertyListView.Items[$CreatePropertyListView.SelectedIndices[0]].Remove()} else {write-host "FUCK YOU!"} 
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonDeleteItem)
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

Function Add-HeaderToViewList ($ListView, $HeaderText, $Width) {
$ColumnHeader = New-Object System.Windows.Forms.ColumnHeader
$ColumnHeader.Text = $HeaderText
$ColumnHeader.Width = $Width
$ListView.Columns.Add($ColumnHeader)
}
