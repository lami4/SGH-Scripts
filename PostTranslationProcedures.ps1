#arrays
$files = @()

#Functions
$script:TranslateBuiltInProperties = $false
$script:TranslateCustomProperties = $false
$script:UpdateFieldsInDocumentBody = $false
$script:UpdateFieldsInFootersAndHeaders = $false
$script:UpdateTOC = $false
$script:UnhideHiddenText = $false
$script:RunATFmacro = $false
$script:RunLWIDBmacro = $false
$script:RunAFISmacro = $false
$script:SelectedFile = ""
$script:SelectedFolder = ""

Function Custom-Form {
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.Form
$dialog.ShowIcon = $false
$dialog.AutoSize = $true
$dialog.Text = "Script Settings"
$dialog.AutoSizeMode = "GrowAndShrink"
$dialog.WindowState = "Normal"
$dialog.SizeGripStyle = "Hide"
$dialog.ShowInTaskbar = $true
$dialog.StartPosition = "CenterScreen"
$dialog.MinimizeBox = $false
$dialog.MaximizeBox = $false
#Labels
#BrowseFile label
$labelBrowseFile = New-Object System.Windows.Forms.Label
$labelBrowseFile.Text = "Specify Excel file that contains translated properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 110
$SystemDrawingPoint.Y = 73
$labelBrowseFile.Location = $SystemDrawingPoint
$labelBrowseFile.Width = 300
$labelBrowseFile.Enabled = $false
#BrowseFolder label
$labelBrowseFolder = New-Object System.Windows.Forms.Label
$labelBrowseFolder.Text = "Specify folder that contains documents to be processed"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 110
$SystemDrawingPoint.Y = 32
$labelBrowseFolder.Location = $SystemDrawingPoint
$labelBrowseFolder.Width = 300
$labelBrowseFolder.Height = 40
$labelBrowseFolder.Enabled = $true
#Buttons
#BrowseFile
$buttonBrowseFile = New-Object System.Windows.Forms.Button
$buttonBrowseFile.Height = 30
$buttonBrowseFile.Width = 80
$buttonBrowseFile.Text = "Browse..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 65
$buttonBrowseFile.Location = $SystemDrawingPoint
$buttonBrowseFile.Enabled = $false
$buttonBrowseFile.Add_Click({
                        Select-File
                        if ($script:SelectedFile -ne "") {$labelBrowseFile.Text = "Selected file: $([System.IO.Path]::GetFileName($script:SelectedFile))"}
                        if ($script:SelectedFile -ne "" -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true}
})
#BrowseFolder
$buttonBrowseFolder = New-Object System.Windows.Forms.Button
$buttonBrowseFolder.Height = 30
$buttonBrowseFolder.Width = 80
$buttonBrowseFolder.Text = "Browse..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$buttonBrowseFolder.Location = $SystemDrawingPoint
$buttonBrowseFolder.Enabled = $true
$buttonBrowseFolder.Add_Click({
                        Select-Folder -description "Specify folder that contains documents to be processed"
                           if ($script:SelectedFolder -ne "") {
                           [string]$ThreeDirectories = "..."
                           $ThreeDirectories += "\$(Split-Path (Split-Path (Split-Path "$script:SelectedFolder" -Parent) -Parent) -Leaf)"
                           $ThreeDirectories += "\$(Split-Path (Split-Path "$script:SelectedFolder" -Parent) -Leaf)"
                           $ThreeDirectories += "\$((Get-Item "$script:SelectedFolder").Name)"
                           $labelBrowseFolder.Text = "Selected path: $ThreeDirectories"
                           }
                           if ($script:SelectedFolder -ne "" -and $checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true}
                           if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -eq "") {$buttonRunScript.Enabled = $false}
})
#Run Script
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Height = 30
$buttonRunScript.Width = 80
$buttonRunScript.Text = "Run Script"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 335
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click({
                            if ($checkboxTranslateBuiltInProperties.Checked) {$script:TranslateBuiltInProperties = $true};
                            if ($checkboxTranslateCustomProperties.Checked) {$script:TranslateCustomProperties = $true};
                            if ($checkboxUpdateFieldsBody.Checked) {$script:UpdateFieldsInDocumentBody = $true};
                            if ($checkboxUpdateFieldsFooterHeader.Checked) {$script:UpdateFieldsInFootersAndHeaders = $true};
                            if ($checkboxUpdateTOC.Checked) {$script:UpdateTOC = $true};
                            if ($checkboxUnhideHiddenText.Checked) {$script:UnhideHiddenText = $true};
                            if ($checkboxRunATFmacro.Checked) {$script:RunATFmacro = $true};
                            if ($checkboxRunLWIDBmacro.Checked) {$script:RunLWIDBmacro = $true};
                            if ($checkboxRunAFISmacro.Checked) {$script:RunAFISmacro = $true};
                            $dialog.DialogResult = "OK";
                            $dialog.Close()})
$buttonRunScript.Enabled = $false
#Exit
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Height = 30
$buttonExit.Width = 80
$buttonExit.Text = "Exit"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 115
$SystemDrawingPoint.Y = 335
$buttonExit.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonExit.Margin = $SystemWindowsFormsMargin
$buttonExit.Add_Click({
$dialog.Close();
$dialog.DialogResult = "Cancel"
})
#Checkboxes
#Translate Built-In Properties
$checkboxTranslateBuiltInProperties = New-Object System.Windows.Forms.CheckBox
$checkboxTranslateBuiltInProperties.Width = 300
$checkboxTranslateBuiltInProperties.Text = "Replace Document Built-In Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 100
$checkboxTranslateBuiltInProperties.Location = $SystemDrawingPoint
$checkboxTranslateBuiltInProperties.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                                                           if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -eq "") {$buttonBrowseFile.Enabled = $true;$labelBrowseFile.Enabled = $true;$buttonRunScript.Enabled = $false} else {$buttonBrowseFile.Enabled = $false;$labelBrowseFile.Enabled = $false}
                                                           if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -ne "") {$buttonBrowseFile.Enabled = $true;$labelBrowseFile.Enabled = $true;$buttonRunScript.Enabled = $false}
                                                           if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -ne "" -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true}
                                                          })
#Translate Custom Properties
$checkboxTranslateCustomProperties = New-Object System.Windows.Forms.CheckBox
$checkboxTranslateCustomProperties.Width = 300
$checkboxTranslateCustomProperties.Text = "Replace Document Custom Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 125
$checkboxTranslateCustomProperties.Location = $SystemDrawingPoint
$checkboxTranslateCustomProperties.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}
                                                          if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -eq "") {$buttonBrowseFile.Enabled = $true;$labelBrowseFile.Enabled = $true;$buttonRunScript.Enabled = $false} else {$buttonBrowseFile.Enabled = $false;$labelBrowseFile.Enabled = $false}
                                                          if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -ne "") {$buttonBrowseFile.Enabled = $true;$labelBrowseFile.Enabled = $true;$buttonRunScript.Enabled = $false}
                                                          if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -and $script:SelectedFile -ne "" -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true}
                                                         })
#Update Fields in Document Body
$checkboxUpdateFieldsBody = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsBody.Width = 300
$checkboxUpdateFieldsBody.Text = "Update Fields in Document Body"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 150
$checkboxUpdateFieldsBody.Location = $SystemDrawingPoint
$checkboxUpdateFieldsBody.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Fields in Document Footers and Headers
$checkboxUpdateFieldsFooterHeader = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsFooterHeader.Width = 300
$checkboxUpdateFieldsFooterHeader.Text = "Update Fields in Document Footers and Headers"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 175
$checkboxUpdateFieldsFooterHeader.Location = $SystemDrawingPoint
$checkboxUpdateFieldsFooterHeader.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Table of Content in Documents
$checkboxUpdateTOC = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateTOC.Width = 300
$checkboxUpdateTOC.Text = "Update Table of Content in Document"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 200
$checkboxUpdateTOC.Location = $SystemDrawingPoint
$checkboxUpdateTOC.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Unhide Hidden Text
$checkboxUnhideHiddenText = New-Object System.Windows.Forms.CheckBox
$checkboxUnhideHiddenText.Width = 300
$checkboxUnhideHiddenText.Text = "Unhide Hidden Text in Document"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 225
$checkboxUnhideHiddenText.Location = $SystemDrawingPoint
$checkboxUnhideHiddenText.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "ApplyTitleFormattingInTranslatedDocument" macro
$checkboxRunATFmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunATFmacro.Width = 375
$checkboxRunATFmacro.Text = "Run 'ApplyTitleFormattingInTranslatedDocument' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 250
$checkboxRunATFmacro.Location = $SystemDrawingPoint
$checkboxRunATFmacro.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "LocateWatermarksInTranslatedDocumentBody" macro
$checkboxRunLWIDBmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunLWIDBmacro.Width = 375
$checkboxRunLWIDBmacro.Text = "Run 'LocateWatermarksInTranslatedDocument' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 275
$checkboxRunLWIDBmacro.Location = $SystemDrawingPoint
$checkboxRunLWIDBmacro.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "ApplyFormattingInSpecification" macro
$checkboxRunAFISmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunAFISmacro.Width = 375
$checkboxRunAFISmacro.Text = "Run 'ApplyFormattingInSpecification' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 300
$checkboxRunAFISmacro.Location = $SystemDrawingPoint
$checkboxRunAFISmacro.Add_CheckStateChanged({if ($checkboxRunAFISmacro.Checked -or $checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked -and $script:SelectedFolder -ne "") {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#GET PROPERTIES UI
#Run Script
$buttonRunScriptForSP = New-Object System.Windows.Forms.Button
$buttonRunScriptForSP.Height = 30
$buttonRunScriptForSP.Width = 80
$buttonRunScriptForSP.Text = "Get props"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 360
$buttonRunScriptForSP.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScriptForSP.Margin = $SystemWindowsFormsMargin
$buttonRunScriptForSP.Add_Click({})
$buttonRunScriptForSP.Enabled = $false
#Exit
$buttonExitForSP = New-Object System.Windows.Forms.Button
$buttonExitForSP.Height = 30
$buttonExitForSP.Width = 80
$buttonExitForSP.Text = "Exit"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 115
$SystemDrawingPoint.Y = 360
$buttonExitForSP.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonExitForSP.Margin = $SystemWindowsFormsMargin
$buttonExitForSP.Add_Click({})
#BrowseFolderButtonForSP
$buttonBrowseFolderForSP = New-Object System.Windows.Forms.Button
$buttonBrowseFolderForSP.Height = 30
$buttonBrowseFolderForSP.Width = 80
$buttonBrowseFolderForSP.Text = "Browse..."
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$buttonBrowseFolderForSP.Location = $SystemDrawingPoint
$buttonBrowseFolderForSP.Enabled = $true
$buttonBrowseFolderForSP.Add_Click({})
#BrowseFolderLabelForSP
$labelBrowseFolderForSP = New-Object System.Windows.Forms.Label
$labelBrowseFolderForSP.Text = "Specify folder with documents whose properties will be extracted"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 110
$SystemDrawingPoint.Y = 32
$labelBrowseFolderForSP.Location = $SystemDrawingPoint
$labelBrowseFolderForSP.Width = 400
$labelBrowseFolderForSP.Height = 40
$labelBrowseFolderForSP.Enabled = $true
#CheckboxGetBuiltInProperties
$checkboxGetBuiltInProperties = New-Object System.Windows.Forms.CheckBox
$checkboxGetBuiltInProperties.Width = 300
$checkboxGetBuiltInProperties.Text = "Get Built-In Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 65
$checkboxGetBuiltInProperties.Location = $SystemDrawingPoint
$checkboxGetBuiltInProperties.Add_CheckStateChanged({})
#CheckboxGetCustomInProperties
$checkboxGetCustomProperties = New-Object System.Windows.Forms.CheckBox
$checkboxGetCustomProperties.Width = 300
$checkboxGetCustomProperties.Text = "Get Custom Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 90
$checkboxGetCustomProperties.Location = $SystemDrawingPoint
$checkboxGetCustomProperties.Add_CheckStateChanged({})
#CheckboxGetPropertiesWithNoValue
$checkboxGetPropertiesWithNoValue = New-Object System.Windows.Forms.CheckBox
$checkboxGetPropertiesWithNoValue.Width = 300
$checkboxGetPropertiesWithNoValue.Text = "Ignore Properties That Have No Value"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 115
$checkboxGetPropertiesWithNoValue.Location = $SystemDrawingPoint
$checkboxGetPropertiesWithNoValue.Add_CheckStateChanged({})
#ListBox
$blacklist = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
$GetPropertiesListBox = New-Object System.Windows.Forms.ListBox
$GetPropertiesListBox.Width = 190
$GetPropertiesListBox.Height = 200
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 140
$GetPropertiesListBox.Location = $SystemDrawingPoint
$blacklist | % {$GetPropertiesListBox.Items.Add($_)}
$GetPropertiesListBox.HorizontalScrollbar = $true
#inputbox
$GetPropertiesTextBox = New-Object System.Windows.Forms.TextBox 
$GetPropertiesTextBox.Location = New-Object System.Drawing.Size(230,140) 
$GetPropertiesTextBox.Size = New-Object System.Drawing.Size(200,30)
#AddPropertyToTheListButton
$buttonAddPropertyToTheListButton = New-Object System.Windows.Forms.Button
$buttonAddPropertyToTheListButton.Height = 22
$buttonAddPropertyToTheListButton.Width = 60
$buttonAddPropertyToTheListButton.Text = "Add"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 430
$SystemDrawingPoint.Y = 139
$buttonAddPropertyToTheListButton.Location = $SystemDrawingPoint
$buttonAddPropertyToTheListButton.Enabled = $true
$buttonAddPropertyToTheListButton.Add_Click({
$GetPropertiesListBox.Items.Insert(0, $GetPropertiesTextBox.Text)
})
#tabs
#Tab Control 
$TabControl = New-object System.Windows.Forms.TabControl
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 5
$SystemDrawingPoint.Y = 5
$tabControl.Location = $SystemDrawingPoint
$SystemDrawingSize = New-Object System.Drawing.Size
$SystemDrawingSize.Height = 450
$SystemDrawingSize.Width = 550
$tabControl.Size = $SystemDrawingSize
$TabControl.AutoSize = $true
#GetProperties page
$GetPropertiesPage = New-Object System.Windows.Forms.TabPage
$GetPropertiesPage.Name = "Get Properties"
$GetPropertiesPage.Text = "Get Properties”
#SetProperties page
$SetPropertiesPage = New-Object System.Windows.Forms.TabPage
$SetPropertiesPage.Name = "Set Properties"
$SetPropertiesPage.Text = "Set Properties”

$dialog.Controls.Add($tabControl)
$tabControl.Controls.Add($GetPropertiesPage)
$tabControl.Controls.Add($SetPropertiesPage)
$GetPropertiesPage.Controls.Add($buttonAddPropertyToTheListButton)
$GetPropertiesPage.Controls.Add($GetPropertiesTextBox)
$GetPropertiesPage.Controls.Add($GetPropertiesListBox)
$GetPropertiesPage.Controls.Add($buttonBrowseFolderForSP)
$GetPropertiesPage.Controls.Add($labelBrowseFolderForSP)
$GetPropertiesPage.Controls.Add($checkboxGetCustomProperties)
$GetPropertiesPage.Controls.Add($checkboxGetBuiltInProperties)
$GetPropertiesPage.Controls.Add($checkboxGetPropertiesWithNoValue)
$GetPropertiesPage.Controls.Add($buttonRunScriptForSP)
$GetPropertiesPage.Controls.Add($buttonExitForSP)
$SetPropertiesPage.Controls.Add($checkboxTranslateBuiltInProperties)
$SetPropertiesPage.Controls.Add($checkboxTranslateCustomProperties)
$SetPropertiesPage.Controls.Add($checkboxUpdateFieldsBody)
$SetPropertiesPage.Controls.Add($checkboxUpdateFieldsFooterHeader)
$SetPropertiesPage.Controls.Add($checkboxUpdateTOC)
$SetPropertiesPage.Controls.Add($checkboxUnhideHiddenText)
$SetPropertiesPage.Controls.Add($checkboxRunATFmacro)
$SetPropertiesPage.Controls.Add($checkboxRunLWIDBmacro)
$SetPropertiesPage.Controls.Add($checkboxRunAFISmacro)
$SetPropertiesPage.Controls.Add($buttonRunScript)
$SetPropertiesPage.Controls.Add($labelBrowseFolder)
$SetPropertiesPage.Controls.Add($labelBrowseFile)
$SetPropertiesPage.Controls.Add($buttonBrowseFolder)
$SetPropertiesPage.Controls.Add($buttonBrowseFile)
$SetPropertiesPage.Controls.Add($buttonExit)
$dialog.ShowDialog()
}

Function Input-YesOrNo ($Question, $BoxTitle) {
$a = New-Object -ComObject wscript.shell
$intAnswer = $a.popup($Question,0,$BoxTitle,4)
If ($intAnswer -eq 6) {
  $script:yesNoUserInput = 1
}
}

Function Select-File {
Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = "$PSScriptRoot"
$f.Filter = "MS Excel Files (*.xls*)|*.xls*|All Files (*.*)|*.*"
$show = $f.ShowDialog()
If ($show -eq "OK") {$script:SelectedFile = $f.FileName}
}

Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {$script:SelectedFolder = $objForm.SelectedPath}
}

Function Set-Properties ($PropertyName, $PropertyValue, $DocumentProperties, $Binding) {
$pn = [System.__ComObject].InvokeMember(“item”,$Binding::GetProperty,$null,$DocumentProperties,$PropertyName)
[System.__ComObject].InvokeMember(“value”,$Binding::SetProperty,$null,$pn,$PropertyValue)
}

$result = Custom-Form
if ($result -ne "OK") {Exit}
Write-Host "$script:TranslateBuiltInProperties"
Write-Host "$script:TranslateCustomProperties"
Write-Host "$script:UpdateFieldsInDocumentBody"
Write-Host "$script:UpdateFieldsInFootersAndHeaders"
Write-Host "$script:UpdateTOC"
Write-Host "$script:UnhideHiddenText"
Write-Host "$script:RunATFmacro"
Write-Host "$script:RunLWIDBmacro"
Write-host "$script:RunAFISmacro"
if ($script:TranslateBuiltInProperties -eq $true -or $script:TranslateCustomProperties -eq $true) {
#Translation of properties
#word
$application = New-Object -ComObject word.application
$application.Visible = $false
#excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Open($script:SelectedFile)
$worksheet = $workbook.Worksheets.Item(1)
$xldown = -4121
$lastNonemptyCellInColumn = $worksheet.Range("E:E").End($xldown).Row
for ($i = 2; $i -le $lastNonemptyCellInColumn; $i++) {
[string]$valueInCell = $worksheet.Cells.Item($i, "E").Value()
$files += $valueInCell
}
Write-Host "Getting ready to replace properties in documents..."
Write-Host "Properties will be replaced in" $files.Count "files."
for ($i = 0; $i -lt $files.Count; $i++) {
    $currentFileName = $files[$i]
    Write-Host "Replacing properties in $currentFileName..."
    #checks if the document exists
    $existence = Test-Path -Path "$script:SelectedFolder\$currentFileName"
    if ($existence -eq $true) {
    $document = $application.documents.open("$script:SelectedFolder\$currentFileName")
    #Unprotects the document
    try {$document.Unprotect(); Write-Host "Document was protected, but got unprotected by the script."} catch {Write-Host "Document is not protected. Unprotection is not required."}
    $builtInProperties = $document.BuiltInDocumentProperties
    $customProperties = $document.CustomDocumentProperties
    $binding = “System.Reflection.BindingFlags” -as [type]
    $range = $worksheet.Range("C:C")
    $target = $range.Find($files[$i])
        if ($target -eq $null) {
        Write-Host "No properties to translate for" $files[$i]
        } else {
        $firstHit = $target
        Do
        {
    #Write-Host "Value found ("$target.AddressLocal()")"
    $currentAddress = $target.AddressLocal($false, $false) -replace "C", ""
    $propertyName = $worksheet.Cells.Item($currentAddress, "A").Value()
    #Write-Host "Name:" $propertyName
    $propertyValue = $worksheet.Cells.Item($currentAddress, "B").Value()
    #Write-Host "Value:" $propertyValue
    $propertyType = $worksheet.Cells.Item($currentAddress, "D").Value()
    Write-Host "Type:" $propertyType
        if ($propertyType -eq "B") {
            if ($script:TranslateBuiltInProperties -eq $true) {
            #set new translated values for BuiltInProperties
            Set-Properties -PropertyName $propertyName -PropertyValue $propertyValue -DocumentProperties $builtInProperties -Binding $binding
            }
        } else {
            if ($script:TranslateCustomProperties -eq $true) {
            #set new translated values for CustomProperties
            Set-Properties -PropertyName $propertyName -PropertyValue $propertyValue -DocumentProperties $customProperties -Binding $binding
            }
        }
        $target = $range.FindNext($target)
        }
        While ($target.AddressLocal() -ne $firstHit.AddressLocal())
        }
#Write-Host "End of document"
$document.Close()
} else {Write-Host "Document not found"}
}
Write-Host "Closing Excel and Word applications..."
Start-Sleep -Seconds 3
$workbook.Close()
Start-Sleep -Seconds 3
$excel.Quit()
Start-Sleep -Seconds 3
$application.Quit()
Start-Sleep -Seconds 3
}

#Update fields, unhide text, apply formatting and locate watermarks
if ($script:UpdateFieldsInDocumentBody -eq $true -or $script:UpdateFieldsInFootersAndHeaders -eq $true -or $script:UpdateTOC -eq $true -or $script:UnhideHiddenText -eq $true) {
Write-Host "Getting ready to unhide hidden text/update fields and/or TOCs in documents (depending on what you've checked)..."
$application = New-Object -ComObject word.application
$application.Visible = $false
Get-ChildItem -Path "$script:SelectedFolder\*.*" -Include "*.doc*", "*.dot*" | % {
Write-Host "Processing $($_.Name)..."
$document = $application.documents.open($_.FullName)
#Unprotects the document
try {$document.Unprotect(); Write-Host "Document was protected, but got unprotected by the script."} catch {Write-Host "Document is not protected. Unprotection is not required."}
Start-Sleep -Seconds 5
#updates fields in footers and headers
if ($script:UpdateFieldsInFootersAndHeaders -eq $true) {
    $sectionCount = $document.Sections.Count
    for ($t = 1; $t -le $sectionCount; $t++) {
        $rangeHeader = $document.Sections.Item($t).Headers.Item(1).Range
        $rangeHeader.Fields.Update() | Out-Null
        $rangeFooter = $document.Sections.Item($t).Footers.Item(1).Range
        $rangeFooter.Fields.Update() | Out-Null
}
}
#updates TOC
if ($script:UpdateTOC -eq $true) {
    $tocCount = $document.TablesOfContents.Count
    if ($tocCount -ge 1) {
    $document.TablesOfContents.Item(1).Update()
    $document.TablesOfContents.Item(1).UpdatePageNumbers()
    }
}
#unhides hidden text
if ($script:UnhideHiddenText -eq $true) { 
    $wholestory = $document.Range()
    $wholestory.Font.Hidden = $false
}
#updates fields in the document body
if ($script:UpdateFieldsInDocumentBody -eq $true) {
    $wholestory = $document.Range()
    $page = $wholestory.Information(4)  
    if ($_.BaseName -match "SPC") {
    try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(5, 10).Range.Text = $page} catch {Write-Host "Cannot update Total Number of Pages as the document is using an old title page" -ForegroundColor Red}
    } else {
    try {$document.Tables.Item(1).Cell(10, 12).Range.Text = $page} catch {Write-Host "Cannot update Total Number of Pages as the document is using an old title page" -ForegroundColor Red}
    }
    $document.Fields.Update()
}
$document.Close()
}
Start-Sleep -Seconds 3
$application.Quit()
}

#runs scripts in documents
if ($script:RunATFmacro -eq $true -or $script:RunLWIDBmacro -eq $true -or $script:RunAFISmacro -eq $true) {
Write-Host "Getting ready to run macroses..."
$application = New-Object -ComObject Word.Application
$application.Visible = $false
Get-ChildItem -Path "$script:SelectedFolder\*.*" -Include "*.doc*", "*.dot*" | % {
$document = $application.documents.open($_.FullName)
if ($script:RunATFmacro -eq $true -and $_.Name -notmatch "SPC") {Write-Host "Working on $($_.Name):";Write-Host "Running ApplyTitleFormattingInTranslatedDocument...";$application.Run("ApplyTitleFormattingInTranslatedDocument")}
if ($script:RunLWIDBmacro -eq $true -and $_.Name -notmatch "SPC") {Write-Host "Working on $($_.Name):";Write-Host "Running LocateWatermarksInTranslatedDocument...";$application.Run("LocateWatermarksInTranslatedDocument")}
if ($script:RunAFISmacro -eq $true -and $_.Name -match "SPC") {Write-Host "Working on $($_.Name):";Write-Host "Running ApplyFormattingInSpecification...";$application.Run("ApplyFormattingInSpecification")}
$document.Close()
}
$application.Quit()
}
