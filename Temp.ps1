    #CREATE PROPERTIES PAGE
    $CreatePropertiesPage = New-Object System.Windows.Forms.TabPage
    $CreatePropertiesPage.Text = "Create Custom Properties”
    $TabControl.Controls.Add($CreatePropertiesPage)
    #CREATE PROPERTIES PAGE ELEMENTS
    #Groupbox 'List Settings'
    $CreatePropertiesPageListSettings = New-Object System.Windows.Forms.GroupBox
    $CreatePropertiesPageListSettings.Location = New-Object System.Drawing.Point(10,10) #x,y
    $CreatePropertiesPageListSettings.Size = New-Object System.Drawing.Point(530,267) #width,height
    $CreatePropertiesPageListSettings.Text = "List Settings"
    $CreatePropertiesPage.Controls.Add($CreatePropertiesPageListSettings)
    #Table
    $CreatePropertyListView = New-Object System.Windows.Forms.ListView
    $CreatePropertyListView.Location = New-Object System.Drawing.Point(10,20) #x, y
    $CreatePropertyListView.View = "Details"
    $CreatePropertyListView.FullRowSelect = $true
    $CreatePropertyListView.Width = 510
    $CreatePropertyListView.Height = 200
    Add-HeaderToViewList -ListView $CreatePropertyListView -HeaderText "Property name" -Width 200
    Add-HeaderToViewList -ListView $CreatePropertyListView -HeaderText "Property value" -Width 200
    Add-HeaderToViewList -ListView $CreatePropertyListView -HeaderText "Data type" -Width 105
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertyListView)
    #Button 'Add'
    $CreatePropertiesPageButtonAddItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonAddItem.Location = New-Object System.Drawing.Point(10,230) #x,y
    $CreatePropertiesPageButtonAddItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonAddItem.Text = "Add item"
    $CreatePropertiesPageButtonAddItem.Add_Click({
    $AddItemDialogResult = Add-ItemToList
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonAddItem)
    #Button 'Edit'
    $CreatePropertiesPageButtonEditItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonEditItem.Location = New-Object System.Drawing.Point(98,230) #x,y
    $CreatePropertiesPageButtonEditItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonEditItem.Text = "Edit item"
    $CreatePropertiesPageButtonEditItem.Add_Click({})
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonEditItem)
    #Button 'Delete'
    $CreatePropertiesPageButtonDeleteItem = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonDeleteItem.Location = New-Object System.Drawing.Point(186,230) #x,y
    $CreatePropertiesPageButtonDeleteItem.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonDeleteItem.Text = "Delete item"
    $CreatePropertiesPageButtonDeleteItem.Add_Click({})
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonDeleteItem)
    #Button 'Clear'
    $CreatePropertiesPageButtonClearList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonClearList.Location = New-Object System.Drawing.Point(274,230) #x,y
    $CreatePropertiesPageButtonClearList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonClearList.Text = "Clear list"
    $CreatePropertiesPageButtonClearList.Add_Click({})
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonClearList)
    #Button 'Export'
    $CreatePropertiesPageButtonExportList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonExportList.Location = New-Object System.Drawing.Point(362,230) #x,y
    $CreatePropertiesPageButtonExportList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonExportList.Text = "Export list"
    $CreatePropertiesPageButtonExportList.Add_Click({
    $CreatePropertyListView.Items[0].SubItems | % {Write-Host $_}
    })
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonExportList)
    #Button 'Import'
    $CreatePropertiesPageButtonImportList = New-Object System.Windows.Forms.Button
    $CreatePropertiesPageButtonImportList.Location = New-Object System.Drawing.Point(450,230) #x,y
    $CreatePropertiesPageButtonImportList.Size = New-Object System.Drawing.Point(70,22) #width,height
    $CreatePropertiesPageButtonImportList.Text = "Import list"
    $CreatePropertiesPageButtonImportList.Add_Click({})
    $CreatePropertiesPageListSettings.Controls.Add($CreatePropertiesPageButtonImportList)
    $CreatePropertiesPage.Controls.Add($CreatePropertyTest)
    
Function Add-ItemToList () {
    $ItemForm = New-Object System.Windows.Forms.Form
    $ItemForm.ShowIcon = $false
    $ItemForm.AutoSize = $true
    $ItemForm.Text = "Add item"
    $ItemForm.AutoSizeMode = "GrowAndShrink"
    $ItemForm.WindowState = "Normal"
    $ItemForm.SizeGripStyle = "Hide"
    $ItemForm.ShowInTaskbar = $true
    $ItemForm.StartPosition = "CenterScreen"
    $ItemForm.MinimizeBox = $false
    $ItemForm.MaximizeBox = $false
    #Label for 'Property name' input field
    $ItemFormPropertyNameLabel = New-Object System.Windows.Forms.Label
    $ItemFormPropertyNameLabel.Location =  New-Object System.Drawing.Point(10,15) #x,y
    $ItemFormPropertyNameLabel.Width = 81
    $ItemFormPropertyNameLabel.Text = "Property name:"
    $ItemFormPropertyNameLabel.TextAlign = "TopRight"
    $ItemForm.Controls.Add($ItemFormPropertyNameLabel)
    #'Property name' input field
    $ItemFormPropertyNameInput = New-Object System.Windows.Forms.TextBox 
    $ItemFormPropertyNameInput.Location = New-Object System.Drawing.Size(95,13) #x,y
    $ItemFormPropertyNameInput.Width = 190
    $ItemFormPropertyNameInput.Text = "Type in property name..."
    $ItemFormPropertyNameInput.ForeColor = "Gray"
    $ItemFormPropertyNameInput.Add_GotFocus({
        if ($ItemFormPropertyNameInput.Text -eq "Type in property name...") {
            $ItemFormPropertyNameInput.Text = ""
            $ItemFormPropertyNameInput.ForeColor = "Black"
        }
        })
    $ItemFormPropertyNameInput.Add_LostFocus({
        if ($ItemFormPropertyNameInput.Text -eq "") {
            $ItemFormPropertyNameInput.Text = "Type in property name..."
            $ItemFormPropertyNameInput.ForeColor = "Gray"
        }
        })
    $ItemForm.Controls.Add($ItemFormPropertyNameInput)
    #Label for 'Property value' input field
    $ItemFormPropertyValueLabel = New-Object System.Windows.Forms.Label
    $ItemFormPropertyValueLabel.Location =  New-Object System.Drawing.Point(10,45) #x,y
    $ItemFormPropertyValueLabel.Width = 81
    $ItemFormPropertyValueLabel.Text = "Property value:"
    $ItemFormPropertyValueLabel.TextAlign = "TopRight"
    $ItemForm.Controls.Add($ItemFormPropertyValueLabel)
    #'Property value' input field
    $ItemFormPropertyValueInput = New-Object System.Windows.Forms.TextBox 
    $ItemFormPropertyValueInput.Location = New-Object System.Drawing.Size(95,43) #x,y
    $ItemFormPropertyValueInput.Width = 190
    $ItemFormPropertyValueInput.Text = "Type in property value..."
    $ItemFormPropertyValueInput.ForeColor = "Gray"
    $ItemFormPropertyValueInput.Add_GotFocus({
        if ($ItemFormPropertyValueInput.Text -eq "Type in property value...") {
            $ItemFormPropertyValueInput.Text = ""
            $ItemFormPropertyValueInput.ForeColor = "Black"
        }
        })
    $ItemFormPropertyValueInput.Add_LostFocus({
        if ($ItemFormPropertyValueInput.Text -eq "") {
            $ItemFormPropertyValueInput.Text = "Type in property value..."
            $ItemFormPropertyValueInput.ForeColor = "Gray"
        }
        })
    $ItemForm.Controls.Add($ItemFormPropertyValueInput)
    #Label for 'Property value' input field
    $ItemFormPropertyTypeLabel = New-Object System.Windows.Forms.Label
    $ItemFormPropertyTypeLabel.Location =  New-Object System.Drawing.Point(10,75) #x,y
    $ItemFormPropertyTypeLabel.Width = 81
    $ItemFormPropertyTypeLabel.Text = "Data type:"
    $ItemFormPropertyTypeLabel.TextAlign = "TopRight"
    $ItemForm.Controls.Add($ItemFormPropertyTypeLabel)
    #Combobox 'Data type'
    $DataTypes = @("Text","Yes or no","Integer", "Number")
    $ItemFormPropertyTypeCombobox = New-Object System.Windows.Forms.ComboBox
    $ItemFormPropertyTypeCombobox.Location = New-Object System.Drawing.Point(95,73) #x,y
    $ItemFormPropertyTypeCombobox.DropDownStyle = "DropDownList"
    $DataTypes | % {$ItemFormPropertyTypeCombobox.Items.add($_)}
    $ItemFormPropertyTypeCombobox.SelectedIndex = 0
    $ItemForm.Controls.Add($ItemFormPropertyTypeCombobox)
    #Buttom 'Add'
    $ItemFormAddButton = New-Object System.Windows.Forms.Button
    $ItemFormAddButton.Location = New-Object System.Drawing.Point(10,115) #x,y
    $ItemFormAddButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $ItemFormAddButton.Text = "Add"
    $ItemFormAddButton.Add_Click({
        $ItemToAdd = New-Object System.Windows.Forms.ListViewItem("$($ItemFormPropertyNameInput.Text)")
        $ItemToAdd.SubItems.Add("$($ItemFormPropertyValueInput.Text)")
        $ItemToAdd.SubItems.Add("$($ItemFormPropertyTypeCombobox.SelectedItem)")
        $CreatePropertyListView.Items.Add($ItemToAdd)
        Write-Host $ItemFormPropertyTypeCombobox.SelectedItem
        $ItemForm.Close()
    })
    $ItemForm.Controls.Add($ItemFormAddButton)
    #Buttom 'Cancel'
    $ItemFormCancelButton = New-Object System.Windows.Forms.Button
    $ItemFormCancelButton.Location = New-Object System.Drawing.Point(215,115) #x,y
    $ItemFormCancelButton.Size = New-Object System.Drawing.Point(70,22) #width,height
    $ItemFormCancelButton.Text = "Cancel"
    $ItemFormCancelButton.Margin = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ItemFormCancelButton.Add_Click({
        $ItemForm.Close()
    })
    $ItemForm.Controls.Add($ItemFormCancelButton)
    $ItemForm.ActiveControl = $ItemFormPropertyTypeLabel
    $ItemForm.ShowDialog()
}

Function Add-HeaderToViewList ($ListView, $HeaderText, $Width) {
$ColumnHeader = New-Object System.Windows.Forms.ColumnHeader
$ColumnHeader.Text = $HeaderText
$ColumnHeader.Width = $Width
$ListView.Columns.Add($ColumnHeader)
}
