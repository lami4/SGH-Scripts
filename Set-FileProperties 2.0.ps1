#arrays
$files = @()

#Functions
$script:SetPathToFileWithTranslatedProperties = $false
$script:TranslateBuiltInProperties = $false
$script:TranslateCustomProperties = $false
$script:UpdateFieldsInDocumentBody = $false
$script:UpdateFieldsInFootersAndHeaders = $false
$script:UpdateTOC = $false
$script:UnhideHiddenText = $false
$script:RunATFmacro = $false
$script:RunLWIDBmacro = $false

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
$buttonRunScript.Add_Click({
                            if ($checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked) {$script:SetPathToFileWithTranslatedProperties = $true};
                            if ($checkboxTranslateBuiltInProperties.Checked) {$script:TranslateBuiltInProperties = $true};
                            if ($checkboxTranslateCustomProperties.Checked) {$script:TranslateCustomProperties = $true};
                            if ($checkboxUpdateFieldsBody.Checked) {$script:UpdateFieldsInDocumentBody = $true};
                            if ($checkboxUpdateFieldsFooterHeader.Checked) {$script:UpdateFieldsInFootersAndHeaders = $true};
                            if ($checkboxUpdateTOC.Checked) {$script:UpdateTOC = $true};
                            if ($checkboxUnhideHiddenText.Checked) {$script:UnhideHiddenText = $true};
                            if ($checkboxRunATFmacro.Checked) {$script:RunATFmacro = $true};
                            if ($checkboxRunLWIDBmacro.Checked) {$script:RunLWIDBmacro = $true};
                            $dialog.DialogResult = "OK"
                            $dialog.Close()})
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
$SystemDrawingPoint.Y = 25
$checkboxTranslateBuiltInProperties.Location = $SystemDrawingPoint
$checkboxTranslateBuiltInProperties.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Translate Custom Properties
$checkboxTranslateCustomProperties = New-Object System.Windows.Forms.CheckBox
$checkboxTranslateCustomProperties.Width = 300
$checkboxTranslateCustomProperties.Text = "Replace Document Custom Properties"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 50
$checkboxTranslateCustomProperties.Location = $SystemDrawingPoint
$checkboxTranslateCustomProperties.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Fields in Document Body
$checkboxUpdateFieldsBody = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsBody.Width = 300
$checkboxUpdateFieldsBody.Text = "Update Fields in Document Body"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 75
$checkboxUpdateFieldsBody.Location = $SystemDrawingPoint
$checkboxUpdateFieldsBody.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Fields in Document Footers and Headers
$checkboxUpdateFieldsFooterHeader = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateFieldsFooterHeader.Width = 300
$checkboxUpdateFieldsFooterHeader.Text = "Update Fields in Document Footers and Headers"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 100
$checkboxUpdateFieldsFooterHeader.Location = $SystemDrawingPoint
$checkboxUpdateFieldsFooterHeader.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Update Table of Content in Documents
$checkboxUpdateTOC = New-Object System.Windows.Forms.CheckBox
$checkboxUpdateTOC.Width = 300
$checkboxUpdateTOC.Text = "Update Table of Content in Documents"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 125
$checkboxUpdateTOC.Location = $SystemDrawingPoint
$checkboxUpdateTOC.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Unhide Hidden Text
$checkboxUnhideHiddenText = New-Object System.Windows.Forms.CheckBox
$checkboxUnhideHiddenText.Width = 300
$checkboxUnhideHiddenText.Text = "Unhide Hidden Text in Documents"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 150
$checkboxUnhideHiddenText.Location = $SystemDrawingPoint
$checkboxUnhideHiddenText.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "ApplyTitleFormatting" macro
$checkboxRunATFmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunATFmacro.Width = 300
$checkboxRunATFmacro.Text = "Run 'ApplyTitleFormatting' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 175
$checkboxRunATFmacro.Location = $SystemDrawingPoint
$checkboxRunATFmacro.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
#Run "LocateWatermarksInDocumentBody" macro
$checkboxRunLWIDBmacro = New-Object System.Windows.Forms.CheckBox
$checkboxRunLWIDBmacro.Width = 300
$checkboxRunLWIDBmacro.Text = "Run 'LocateWatermarksInDocumentBody' Macro"
$SystemDrawingPoint = New-Object System.Drawing.Point
$SystemDrawingPoint.X = 25
$SystemDrawingPoint.Y = 200
$checkboxRunLWIDBmacro.Location = $SystemDrawingPoint
$checkboxRunLWIDBmacro.Add_CheckStateChanged({if ($checkboxRunLWIDBmacro.Checked -or $checkboxTranslateBuiltInProperties.Checked -or $checkboxTranslateCustomProperties.Checked -or $checkboxUpdateFieldsBody.Checked -or $checkboxUpdateFieldsFooterHeader.Checked -or $checkboxUpdateTOC.Checked -or $checkboxUnhideHiddenText.Checked -or $checkboxRunATFmacro.Checked) {$buttonRunScript.Enabled = $true} else {$buttonRunScript.Enabled = $false}})
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
$f.ShowHelp = $false
$f.Multiselect = $true
$show = $f.ShowDialog()
If ($show -eq "OK") {if ($f.Multiselect) { $f.FileNames } else { $f.FileName }} else {Exit}
}

Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
        If ($Show -eq "OK") {
        Return $objForm.SelectedPath
        } Else {
        Exit
        }
}

Function Set-Properties ($PropertyName, $PropertyValue, $DocumentProperties, $Binding) {
$pn = [System.__ComObject].InvokeMember(“item”,$Binding::GetProperty,$null,$DocumentProperties,$PropertyName)
[System.__ComObject].InvokeMember(“value”,$Binding::SetProperty,$null,$pn,$PropertyValue)
}

$result = Custom-Form
if ($result -ne "OK") {Exit}
<#Write-Host "$script:SetPathToFileWithTranslatedProperties"
Write-Host "$script:TranslateBuiltInProperties"
Write-Host "$script:TranslateCustomProperties"
Write-Host "$script:UpdateFieldsInDocumentBody"
Write-Host "$script:UpdateFieldsInFootersAndHeaders"
Write-Host "$script:UpdateTOC"
Write-Host "$script:UnhideHiddenText"
Write-Host "$script:RunATFmacro"
Write-Host "$script:RunLWIDBmacro"#>
$selectedFolder = Select-Folder -description "Select folder with files to be processed"
if ($SetPathToFileWithTranslatedProperties -eq $true) {
#Translation of properties
$selectedFile = Select-File
#word
$application = New-Object -ComObject word.application
$application.Visible = $false
#excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Open($selectedFile)
$worksheet = $workbook.Worksheets.Item(1)
$xldown = -4121
$lastNonemptyCellInColumn = $worksheet.Range("D:D").End($xldown).Row
for ($i = 2; $i -le $lastNonemptyCellInColumn; $i++) {
[string]$valueInCell = $worksheet.Cells.Item($i, "D").Value()
$files += $valueInCell
}
Write-Host "Properties will be replaced in" $files.Count "files."
for ($i = 0; $i -lt $files.Count; $i++) {
    $currentFileName = $files[$i]
    Write-Host "Replacing properties in $currentFileName..."
    #checks if the document exists
    $existence = Test-Path -Path "$selectedFolder\$currentFileName"
    if ($existence -eq $true) {
    $document = $application.documents.open("$selectedFolder\$currentFileName")
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
    $propertyType = $worksheet.Cells.Item($currentAddress, "E").Value()
    #Write-Host "Type:" $propertyType
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
Write-Host "Getting ready to perform other selected operations..."
$application = New-Object -ComObject word.application
$application.Visible = $false
Get-ChildItem -Path "$selectedFolder\*.*" -Include "*.doc*", "*.dot*" | % {
Write-Host "Processing" $_.Name
$document = $application.documents.open($_.FullName)
#updates fields in the document body
if ($script:UpdateFieldsInDocumentBody -eq $true) {
$document.Fields.Update() | Out-Null
}
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
$document.Close()
}
}
Start-Sleep -Seconds 3
$application.Quit()
