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
    $TabControl.Size = New-Object System.Drawing.Size(500,550)
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
    #Listbox 'Black list'
    $DefaultBlackList = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(25,140) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,250) #width,height
    $DefaultBlackList | % {$GetPropertyListBoxBlackList.Items.Add($_)}
    $GetPropertiesPage.Controls.Add($GetPropertyListBoxBlackList)
    #Inputbox 'Add item to the list'
    $GetPropertyInputboxAddItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxAddItem.Location = New-Object System.Drawing.Size(250,140) #x,y
    $GetPropertyInputboxAddItem.Size = New-Object System.Drawing.Size(180,30) #width,height
    $GetPropertiesPage.Controls.Add($GetPropertyInputboxAddItem)
    #Button 'Add'
    $GetPropertyButtonAddItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonAddItem.Location = New-Object System.Drawing.Point(435,140) #x,y
    $GetPropertyButtonAddItem.Size = New-Object System.Drawing.Point(50,30) #width,height
    $GetPropertyButtonAddItem.Text = "Add"
    #$GetPropertyButtonAddItem.BackColor = "Red"
    $GetPropertyButtonAddItem.TabStop = $false 
    $GetPropertyButtonAddItem.FlatStyle = "Flat"
    $GetPropertyButtonAddItem.FlatAppearance.BorderSize = 0
    $GetPropertyButtonAddItem.Add_Click({
        $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
        })
    $GetPropertiesPage.Controls.Add($GetPropertyButtonAddItem)
    #SET PROPERTIES PAGE
    $SetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $SetPropertiesPage.Text = "Set Properties”
    $TabControl.Controls.Add($SetPropertiesPage)

    $Form.ShowDialog()
}

Custom-Form
