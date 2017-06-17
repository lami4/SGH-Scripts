Function Custom-Form {
    Add-Type -AssemblyName System.Windows.Forms
    #FORM
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon = $false
    $Form.AutoSize = $true
    $Form.Text = "Script Settings"
    $Form.AutoSizeMode = "GrowAndShrink"
    $Form.WindowState = "Normal"
    $Form.SizeGripStyle = "Hide"
    $Form.ShowInTaskbar = $true
    $Form.StartPosition = "CenterScreen"
    $Form.MinimizeBox = $false
    $Form.MaximizeBox = $false
    #TAB CONTROL
    $TabControl = New-object System.Windows.Forms.TabControl
    $TabControl.Location = New-Object System.Drawing.Size(5,5)
    $TabControl.Size = New-Object System.Drawing.Size(550,550) #width,height
    $Form.Controls.Add($TabControl)
    #GET PROPERTIES PAGE
    $GetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $GetPropertiesPage.Text = "Get Properties”
    $TabControl.Controls.Add($GetPropertiesPage)
    #GET PROPERTIES PAGE ELEMENTS
    #Button 'Browse...'
    $GetPropertyButtonBrowse = New-Object System.Windows.Forms.Button
    $GetPropertyButtonBrowse.Location = New-Object System.Drawing.Point(25,25)
    $GetPropertyButtonBrowse.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonBrowse.Text = "Browse..."
    $GetPropertyButtonBrowse.Add_Click({})
    $GetPropertiesPage.Controls.Add($GetPropertyButtonBrowse)
    #Label for 'Browse...' button
    $GetPropertyLabelButtonBrowse = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonBrowse.Location =  New-Object System.Drawing.Point(110,32)
    $GetPropertyLabelButtonBrowse.Size =  New-Object System.Drawing.Point(400,40)
    $GetPropertyLabelButtonBrowse.Text = "Specify folder with documents whose properties will be extracted"
    $GetPropertiesPage.Controls.Add($GetPropertyLabelButtonBrowse)
    #Checkbox 'Get Built-In Properties'
    $GetPropertyCheckboxGetBuiltInProperties = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxGetBuiltInProperties.Location = New-Object System.Drawing.Point(25,65)
    $GetPropertyCheckboxGetBuiltInProperties.Width = 300
    $GetPropertyCheckboxGetBuiltInProperties.Text = "Get Built-In Properties"
    $GetPropertyCheckboxGetBuiltInProperties.Add_CheckStateChanged({})
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxGetBuiltInProperties)
    #Checkbox 'Get Custom Properties'
    $GetPropertyCheckboxGetCustomProperties = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxGetCustomProperties.Location = New-Object System.Drawing.Point(25,90)
    $GetPropertyCheckboxGetCustomProperties.Width = 300
    $GetPropertyCheckboxGetCustomProperties.Text = "Get Custom Properties"
    $GetPropertyCheckboxGetCustomProperties.Add_CheckStateChanged({})
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxGetCustomProperties)
    #Checkbox 'Ignore Properties That Have No Value'
    $GetPropertyCheckboxIgnorePropertiesWithNoValue = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Location = New-Object System.Drawing.Point(25,115)
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Width = 300
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Text = "Ignore Properties That Have No Value"
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Add_CheckStateChanged({})
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxIgnorePropertiesWithNoValue)
    #Checkbox 'Ignore Properties That Have No Value'
    $GetPropertyCheckboxUseBlacklist = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxUseBlacklist.Location = New-Object System.Drawing.Point(25,140)
    $GetPropertyCheckboxUseBlacklist.Width = 300
    $GetPropertyCheckboxUseBlacklist.Text = "Use Blacklist"
    $GetPropertyCheckboxUseBlacklist.Add_CheckStateChanged({})
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxUseBlacklist)
    #Groupbox 'Blacklist settings'
    $GetPropertyGroupboxBlacklistSettings = New-Object System.Windows.Forms.GroupBox
    $GetPropertyGroupboxBlacklistSettings.Location = New-Object System.Drawing.Point(25,170) #x,y
    $GetPropertyGroupboxBlacklistSettings.Size = New-Object System.Drawing.Point(500,295) #width,height
    $GetPropertyGroupboxBlacklistSettings.Text = "Blacklist Settings"
    $GetPropertyGroupboxBlacklistSettings.Enabled = $true
    $GetPropertiesPage.Controls.Add($GetPropertyGroupboxBlacklistSettings)
    #Listbox 'Black list'
    $DefaultBlackList = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(15,25) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,260) #width,height
    $DefaultBlackList | % {$GetPropertyListBoxBlackList.Items.Add($_)}
    $GetPropertyListBoxBlackList.Add_SelectedIndexChanged({
        if ($GetPropertyListBoxBlackList.SelectedIndex -ne -1) {
        Write-Host "$($GetPropertyListBoxBlackList.SelectedIndex)"
        $GetPropertyInputboxEditItem.Text = $GetPropertyListBoxBlackList.SelectedItem
        } else {
        $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it"
        }
    })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyListBoxBlackList)
    #Button 'Add'
    $GetPropertyButtonAddItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonAddItem.Location = New-Object System.Drawing.Point(235,24) #x,y
    $GetPropertyButtonAddItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonAddItem.Text = "Add"
    $GetPropertyButtonAddItem.Add_Click({
        $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonAddItem)
    #Inputbox 'Add item to the list'
    $GetPropertyInputboxAddItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxAddItem.Location = New-Object System.Drawing.Size(295,25) #x,y
    $GetPropertyInputboxAddItem.Width = 190
    $GetPropertyInputboxAddItem.Text = "Type in property name to add it"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxAddItem)
    #Button 'Edit'
    $GetPropertyButtonEditItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonEditItem.Location = New-Object System.Drawing.Point(235,52) #x,y
    $GetPropertyButtonEditItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonEditItem.Text = "Edit"
    $GetPropertyButtonEditItem.Add_Click({})
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonEditItem)
    #Inputbox 'Edit the selected item'
    $GetPropertyInputboxEditItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxEditItem.Location = New-Object System.Drawing.Size(295,53) #x,y
    $GetPropertyInputboxEditItem.Width = 190
    $GetPropertyInputboxEditItem.Enabled = $false
    $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxEditItem)
    #Button 'Delete'
    $GetPropertyButtonDeleteItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonDeleteItem.Location = New-Object System.Drawing.Point(235,80) #x,y
    $GetPropertyButtonDeleteItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonDeleteItem.Text = "Delete"
    $GetPropertyButtonDeleteItem.Add_Click({
        $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
    })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonDeleteItem)
    #Button 'Export'
    $GetPropertyButtonExportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExportList.Location = New-Object System.Drawing.Point(235,108) #x,y
    $GetPropertyButtonExportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonExportList.Text = "Export"
    $GetPropertyButtonExportList.Add_Click({})
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonExportList)
    #Button 'Import'
    $GetPropertyButtonImportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonImportList.Location = New-Object System.Drawing.Point(235,136) #x,y
    $GetPropertyButtonImportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonImportList.Text = "Import"
    $GetPropertyButtonImportList.Add_Click({})
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonImportList)
    #SET PROPERTIES PAGE
    $SetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $SetPropertiesPage.Text = "Set Properties”
    $TabControl.Controls.Add($SetPropertiesPage)

    $Form.ShowDialog()
}

Custom-Form
