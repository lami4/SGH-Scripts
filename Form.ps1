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
#Buttons
#Run Script
$buttonRunScript = New-Object System.Windows.Forms.Button
$buttonRunScript.Height = 30
$buttonRunScript.Width = 80
$buttonRunScript.Text = "Run Script"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 250
$buttonRunScript.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonRunScript.Margin = $SystemWindowsFormsMargin
$buttonRunScript.Add_Click($onClickRunScript)
$buttonRunScript.Enabled = $false
#Exit
$buttonExit = New-Object System.Windows.Forms.Button
$buttonExit.Height = 30
$buttonExit.Width = 80
$buttonExit.Text = "Exit"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 115
$SystemDrawingPoint.Y = 250
$buttonExit.Location = $SystemDrawingPoint
$SystemWindowsFormsMargin = New-Object System.Windows.Forms.Padding
$SystemWindowsFormsMargin.Bottom = 25
$buttonExit.Margin = $SystemWindowsFormsMargin
$buttonExit.Add_Click($onClickRunScript)
#Checkboxes
#Translate Built-In Properties
$checkboxTranslateBuiltInProperties = New-Object System.Windows.Forms.CheckBox
$checkboxTranslateBuiltInProperties.Width = 350
$checkboxTranslateBuiltInProperties.Text = "Translate Document Built-In Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 25
$checkboxTranslateBuiltInProperties.Location = $SystemDrawingPoint
$checkboxTranslateBuiltInProperties.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Translate Custom Properties
$checkboxTranslateCustomProperties = New-Object System.Windows.Forms.CheckBox
$checkboxTranslateCustomProperties.Width = 350
$checkboxTranslateCustomProperties.Text = "Translate Document Custom Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 50
$checkboxTranslateCustomProperties.Location = $SystemDrawingPoint
$checkboxTranslateCustomProperties.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Fields in Document Body
$checkboxUpdateFieldsBody = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsBody.Width = 350
$checkboxUpdateFieldsBody.Text = "Update Fields in Document Body"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 75
$checkboxUpdateFieldsBody.Location = $SystemDrawingPoint
$checkboxUpdateFieldsBody.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Fields in Document Footers and Headers
$checkboxUpdateFieldsFooterHeader = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsFooterHeader.Width = 350
$checkboxUpdateFieldsFooterHeader.Text = "Update Fields in Document Footers and Headers"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 100
$checkboxUpdateFieldsFooterHeader.Location = $SystemDrawingPoint
$checkboxUpdateFieldsFooterHeader.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Table of Content in Documents
$checkboxUpdateTOC = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateTOC.Width = 350
$checkboxUpdateTOC.Text = "Update Table of Content in Documents"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 125
$checkboxUpdateTOC.Location = $SystemDrawingPoint
$checkboxUpdateTOC.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Unhide Hidden Text
$checkboxUnhideHiddenText = New-Object System.Windows.Forms.CheckBox
$checkboxUnhideHiddenText.Width = 350
$checkboxUnhideHiddenText.Text = "Unhide Hidden Text in Documents"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 150
$checkboxUnhideHiddenText.Location = $SystemDrawingPoint
$checkboxUnhideHiddenText.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "ApplyTitleFormatting" macro
$checkboxRunATFmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunATFmacro.Width = 350
$checkboxRunATFmacro.Text = "Run 'ApplyTitleFormatting' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 175
$checkboxRunATFmacro.Location = $SystemDrawingPoint
$checkboxRunATFmacro.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "LocateWatermarksInDocumentBody" macro
$checkboxRunLWIDBmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunLWIDBmacro.Width = 350
$checkboxRunLWIDBmacro.Text = "Run 'LocateWatermarksInDocumentBody' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 200
$checkboxRunLWIDBmacro.Location = $SystemDrawingPoint
$checkboxRunLWIDBmacro.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Event handler
$onClickRunScript = {if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked) {Write-Host "Script started"} else {Write-Host "Script cannot be started"}}
$dialog.Controls.Add($checkboxTranslateBuiltInProperties)
$dialog.Controls.Add($checkboxTranslateCustomProperties)
$dialog.Controls.Add($checkboxUpdateFieldsBody)
$dialog.Controls.Add($checkboxUpdateFieldsFooterHeader)
$dialog.Controls.Add($checkboxUpdateTOC)
$dialog.Controls.Add($checkboxUnhideHiddenText)
$dialog.Controls.Add($checkboxRunATFmacro)
$dialog.Controls.Add($checkboxRunLWIDBmacro)
$dialog.Controls.Add($buttonRunScript)
$dialog.Controls.Add($buttonExit)
$dialog.ShowDialog()
