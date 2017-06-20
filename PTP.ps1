Function Custom-Form 
{
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
    $TabControl.Size = New-Object System.Drawing.Size(560,545) #width,height
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
    $GetPropertyButtonBrowse.Add_Click({
        $PathToSelectedFolder = Select-Folder -Description "Select folder with files whose properties will be extracted"
        if ($PathToSelectedFolder -ne $null) {
            if ($PathToSelectedFolder.Length -gt 85) {
                #$ToolTip.RemoveAll()
                $GetPropertyLabelButtonBrowse.Text = "Specified directory's name is too long to display it here. Hover to see the full path."
                #Tooltip for label of 'Browse...' button
                #$ToolTip = New-Object System.Windows.Forms.ToolTip
                $ToolTip.SetToolTip($GetPropertyLabelButtonBrowse, "$PathToSelectedFolder")
            } else {
                #$ToolTip.RemoveAll()
                $GetPropertyLabelButtonBrowse.Text = "Specified directory: '$(Split-Path -Path $PathToSelectedFolder -Leaf)'. Hover to see the full path."
                $ToolTip.SetToolTip($GetPropertyLabelButtonBrowse, "$PathToSelectedFolder")
            }
        }
        })
    $GetPropertiesPage.Controls.Add($GetPropertyButtonBrowse)
    #Label for 'Browse...' button
    $GetPropertyLabelButtonBrowse = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonBrowse.Location =  New-Object System.Drawing.Point(110,32)
    $GetPropertyLabelButtonBrowse.Width = 400
    $GetPropertyLabelButtonBrowse.Text = "Specify folder with documents whose properties will be extracted"
    $GetPropertyLabelButtonBrowse.AutoSize = $true
    $GetPropertyLabelButtonBrowse.MaximumSize = New-Object System.Drawing.Point(430,38)
    $GetPropertiesPage.Controls.Add($GetPropertyLabelButtonBrowse)
    #Tooltip for label of 'Browse...' button
    $ToolTip = New-Object System.Windows.Forms.ToolTip
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
    #Checkbox 'Use Blacklist'
    $GetPropertyCheckboxUseBlacklist = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxUseBlacklist.Location = New-Object System.Drawing.Point(25,140)
    $GetPropertyCheckboxUseBlacklist.Width = 300
    $GetPropertyCheckboxUseBlacklist.Text = "Use Blacklist"
    $GetPropertyCheckboxUseBlacklist.Add_CheckStateChanged({
    if ($GetPropertyCheckboxUseBlacklist.Checked -eq $true) {$GetPropertyGroupboxBlacklistSettings.Enabled = $true} else {$GetPropertyGroupboxBlacklistSettings.Enabled = $false}
    })
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxUseBlacklist)
    #Groupbox 'Blacklist settings'
    $GetPropertyGroupboxBlacklistSettings = New-Object System.Windows.Forms.GroupBox
    $GetPropertyGroupboxBlacklistSettings.Location = New-Object System.Drawing.Point(25,170) #x,y
    $GetPropertyGroupboxBlacklistSettings.Size = New-Object System.Drawing.Point(500,300) #width,height
    $GetPropertyGroupboxBlacklistSettings.Text = "Blacklist Settings"
    $GetPropertyGroupboxBlacklistSettings.Enabled = $false
    $GetPropertiesPage.Controls.Add($GetPropertyGroupboxBlacklistSettings) 
    #Label for 'Black list' listbox
    $GetPropertyLabelBlacklistListBox = New-Object System.Windows.Forms.Label
    $GetPropertyLabelBlacklistListBox.Location =  New-Object System.Drawing.Point(13,19) #x,y
    $GetPropertyLabelBlacklistListBox.Width = 250
    $GetPropertyLabelBlacklistListBox.Height = 13
    $GetPropertyLabelBlacklistListBox.Text = "Properties blacklisted by default:"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelBlacklistListBox)
    #Listbox 'Black list'
    $DefaultBlackList = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(15,35) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,260) #width,height
    $DefaultBlackList | % {$GetPropertyListBoxBlackList.Items.Add($_)}
    $GetPropertyListBoxBlackList.Add_SelectedIndexChanged({
        if ($GetPropertyListBoxBlackList.SelectedIndex -ne -1) {
            Write-Host "$($GetPropertyListBoxBlackList.SelectedIndex)"
            $GetPropertyInputboxEditItem.Text = $GetPropertyListBoxBlackList.SelectedItem
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyListBoxBlackList)
    #Button 'Add'
    $GetPropertyButtonAddItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonAddItem.Location = New-Object System.Drawing.Point(235,34) #x,y
    $GetPropertyButtonAddItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonAddItem.Text = "Add"
    $GetPropertyButtonAddItem.Add_Click({
        if ($GetPropertyInputboxAddItem.Text -ne "Type in property name to add it...") {
            $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
            $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
            $GetPropertyInputboxAddItem.ForeColor = "Gray"
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonAddItem)
    #Inputbox 'Add item to the list'
    $GetPropertyInputboxAddItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxAddItem.Location = New-Object System.Drawing.Size(295,35) #x,y
    $GetPropertyInputboxAddItem.Width = 190
    $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
    $GetPropertyInputboxAddItem.ForeColor = "Gray"
    $GetPropertyInputboxAddItem.Add_GotFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "Type in property name to add it...") {
            $GetPropertyInputboxAddItem.Text = ""
            $GetPropertyInputboxAddItem.ForeColor = "Black"
        }
        })
    $GetPropertyInputboxAddItem.Add_LostFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "") {
            $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
            $GetPropertyInputboxAddItem.ForeColor = "Gray"
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxAddItem)
    #Button 'Edit'
    $GetPropertyButtonEditItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonEditItem.Location = New-Object System.Drawing.Point(235,62) #x,y
    $GetPropertyButtonEditItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonEditItem.Text = "Edit"
    $GetPropertyButtonEditItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -ne "Select item on the blacklist to edit it...") {
            Disable-AllExceptEditing -BooleanRest $false -BooleanEditing $true
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonEditItem)
    #Inputbox 'Edit the selected item'
    $GetPropertyInputboxEditItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxEditItem.Location = New-Object System.Drawing.Size(295,63) #x,y
    $GetPropertyInputboxEditItem.Width = 190
    $GetPropertyInputboxEditItem.Enabled = $false
    $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it..."
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxEditItem)
    #Button 'Apply'
    $GetPropertyButtonApplyItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonApplyItem.Location = New-Object System.Drawing.Point(295,85) #x,y
    $GetPropertyButtonApplyItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonApplyItem.Text = "Apply"
    $GetPropertyButtonApplyItem.Enabled = $false
    $GetPropertyButtonApplyItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -eq $GetPropertyListBoxBlackList.SelectedItem) {
            Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
        } else {
            $SelectedIndex = $GetPropertyListBoxBlackList.SelectedIndex
            $GetPropertyListBoxBlackList.Items.Insert($GetPropertyListBoxBlackList.SelectedIndex, ($GetPropertyInputboxEditItem.Text).Trim(' '))
            $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
            $GetPropertyListBoxBlackList.SelectedIndex = $SelectedIndex
            Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonApplyItem)
    #Button 'Cancel'
    $GetPropertyButtonCancelItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonCancelItem.Location = New-Object System.Drawing.Point(347,85) #x,y
    $GetPropertyButtonCancelItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonCancelItem.Text = "Cancel"
    $GetPropertyButtonCancelItem.Enabled = $false
    $GetPropertyButtonCancelItem.Add_Click({
        Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false    
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonCancelItem)
    #Button 'Delete'
    $GetPropertyButtonDeleteItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonDeleteItem.Location = New-Object System.Drawing.Point(235,113) #x,y
    $GetPropertyButtonDeleteItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonDeleteItem.Text = "Delete"
    $GetPropertyButtonDeleteItem.Add_Click({
        $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
        $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it..."
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonDeleteItem)
    #Label for 'Delete' button
    $GetPropertyLabelButtonDelete = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonDelete.Location =  New-Object System.Drawing.Point(295,116) #x,y
    $GetPropertyLabelButtonDelete.Size =  New-Object System.Drawing.Point(203,15) #width,height
    $GetPropertyLabelButtonDelete.Text = "Delete unwanted item from the blacklist"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonDelete)
    #Button 'Export'
    $GetPropertyButtonExportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExportList.Location = New-Object System.Drawing.Point(235,141) #x,y
    $GetPropertyButtonExportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonExportList.Text = "Export"
    $GetPropertyButtonExportList.Add_Click({
        $PathToExportedFile = Save-File
        if ($PathToExportedFile -ne $null) {
            If (Test-Path -Path $PathToExportedFile) {Remove-Item -Path $PathToExportedFile -Force}
            $GetPropertyListBoxBlackList.Items | % {Add-Content -Path $PathToExportedFile -Value $_}
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonExportList)
    #Label for 'Export' button
    $GetPropertyLabelButtonExport = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonExport.Location =  New-Object System.Drawing.Point(295,144) #x,y
    $GetPropertyLabelButtonExport.Size =  New-Object System.Drawing.Point(200,15) #width,height
    $GetPropertyLabelButtonExport.Text = "Export your blacklist for later use"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonExport)
    #Button 'Import'
    $GetPropertyButtonImportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonImportList.Location = New-Object System.Drawing.Point(235,169) #x,y
    $GetPropertyButtonImportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonImportList.Text = "Import"
    $GetPropertyButtonImportList.Add_Click({
        $PathToImportedFile = Open-File
        Write-Host $PathToImportedFile
            if ($PathToImportedFile -ne $null) {
                $TxtBlackListContent = @(Get-Content -Path $PathToImportedFile)
                $GetPropertyListBoxBlackList.Items.Clear()
                $TxtBlackListContent | % {$GetPropertyListBoxBlackList.Items.Add($_)}
            }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonImportList)
    #Label for 'Import' button
    $GetPropertyLabelButtonImport = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonImport.Location =  New-Object System.Drawing.Point(295,172) #x,y
    $GetPropertyLabelButtonImport.Size =  New-Object System.Drawing.Point(200,15) #width,height
    $GetPropertyLabelButtonImport.Text = "Import your blacklist"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonImport)
    #Checkbox 'Get Only Blacklisted Properties'
    $GetPropertyCheckboxTurnIntoWhite = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxTurnIntoWhite.Location = New-Object System.Drawing.Point(235,200) #x,y
    $GetPropertyCheckboxTurnIntoWhite.Width = 200
    $GetPropertyCheckboxTurnIntoWhite.Text = "Get Only Blacklisted Properties"
    $GetPropertyCheckboxTurnIntoWhite.Add_CheckStateChanged({})
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyCheckboxTurnIntoWhite)
    #Button 'Extract Properties'
    $GetPropertyButtonExtract = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExtract.Location = New-Object System.Drawing.Point(25,480) #x,y
    $GetPropertyButtonExtract.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonExtract.Text = "Run"
    $GetPropertyButtonExtract.Add_Click({})
    $GetPropertiesPage.Controls.Add($GetPropertyButtonExtract)
    #Button 'Exit'
    $GetPropertyButtonExit = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExit.Location = New-Object System.Drawing.Point(115,480) #x,y
    $GetPropertyButtonExit.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonExit.Text = "Exit"
    $GetPropertyButtonExit.Add_Click({})
    $GetPropertiesPage.Controls.Add($GetPropertyButtonExit)
    #SET PROPERTIES PAGE
    $SetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $SetPropertiesPage.Text = "Set Properties”
    $TabControl.Controls.Add($SetPropertiesPage)
    $Form.ShowDialog()
}

Function Disable-AllExceptEditing ($BooleanRest, $BooleanEditing) 
{
        $GetPropertyButtonBrowse.Enabled = $BooleanRest
        $GetPropertyLabelButtonBrowse.Enabled = $BooleanRest
        $GetPropertyCheckboxGetBuiltInProperties.Enabled = $BooleanRest
        $GetPropertyCheckboxGetCustomProperties.Enabled = $BooleanRest
        $GetPropertyCheckboxIgnorePropertiesWithNoValue.Enabled = $BooleanRest
        $GetPropertyCheckboxUseBlacklist.Enabled = $BooleanRest
        $GetPropertyListBoxBlackList.Enabled = $BooleanRest
        $GetPropertyButtonAddItem.Enabled = $BooleanRest
        $GetPropertyInputboxAddItem.Enabled = $BooleanRest
        $GetPropertyButtonDeleteItem.Enabled = $BooleanRest
        $GetPropertyLabelButtonDelete.Enabled = $BooleanRest
        $GetPropertyButtonExportList.Enabled = $BooleanRest
        $GetPropertyLabelButtonExport.Enabled = $BooleanRest
        $GetPropertyButtonImportList.Enabled = $BooleanRest
        $GetPropertyLabelButtonImport.Enabled = $BooleanRest
        $GetPropertyInputboxEditItem.Enabled = $BooleanEditing
        $GetPropertyButtonApplyItem.Enabled = $BooleanEditing
        $GetPropertyButtonCancelItem.Enabled = $BooleanEditing
        $GetPropertyButtonEditItem.Enabled = $BooleanRest
        $GetPropertyCheckboxTurnIntoWhite.Enabled = $BooleanRest
        $GetPropertyButtonExtract.Enabled = $BooleanRest
        $GetPropertyButtonExit.Enabled = $BooleanRest
}

Function Save-File
{ 
    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "Text file (*.txt)| *.txt"
    $DialogResult = $SaveFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SaveFileDialog.FileName} else {return $null}
}

Function Open-File
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Text file (*.txt)| *.txt"
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $OpenFileDialog.FileName} else {return $null}
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

Custom-Form
