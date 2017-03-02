$files = @()
$script:yesNoUserInput = 0

#Functions
Function Input-YesOrNo ($Question, $BoxTitle) {
$a = New-Object -ComObject wscript.shell
$intAnswer = $a.popup($Question,0,$BoxTitle,4)
If ($intAnswer -eq 6) {
  $script:yesNoUserInput += 1
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

Input-YesOrNo -Question "Do you want to update fields and TOC in documents?" -BoxTitle "Update document fields"
$selectedFolder = Select-Folder -description "Select folder with files whose properties are to be changed"
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
$lastNonemptyCellInColumn = $worksheet.Range("D2").End($xldown).Row
Write-Host $lastNonemptyCellInColumn
for ($i = 2; $i -le $lastNonemptyCellInColumn; $i++) {
[string]$valueInCell = $worksheet.Cells.Item($i, "D").Value()
$files += $valueInCell
}
Write-Host $files.Count
Write-Host $files
for ($i = 0; $i -lt $files.Count; $i++) {
    $currentFileName = $files[$i]
    Write-Host "Processing $currentFileName..."
    $document = $application.documents.open("$selectedFolder\$currentFileName")
    $builtInProperties = $document.BuiltInDocumentProperties
    $customProperties = $document.CustomDocumentProperties
    $binding = “System.Reflection.BindingFlags” -as [type]
    $range = $worksheet.Range("C:C")
    $target = $range.Find($files[$i])
        if ($target -eq $null) {
        Write-Host "No properties to set for" $files[$i]
        } else {
        $firstHit = $target
        Do
        {
    Write-Host "Value found ("$target.AddressLocal()")"
    $currentAddress = $target.AddressLocal($false, $false) -replace "C", ""
    $propertyName = $worksheet.Cells.Item($currentAddress, "A").Value()
    Write-Host "Name:" $propertyName
    $propertyValue = $worksheet.Cells.Item($currentAddress, "B").Value()
    Write-Host "Value:" $propertyValue
    $propertyType = $worksheet.Cells.Item($currentAddress, "E").Value()
    Write-Host "Type:" $propertyType
        if ($propertyType -eq "B") {
        #set new translated values for BuiltInProperties
        Set-Properties -PropertyName $propertyName -PropertyValue $propertyValue -DocumentProperties $builtInProperties -Binding $binding
        } else {
        #set new translated values for CustomProperties
        Set-Properties -PropertyName $propertyName -PropertyValue $propertyValue -DocumentProperties $customProperties -Binding $binding
        }
        $target = $range.FindNext($target)
        }
        While ($target.AddressLocal() -ne $firstHit.AddressLocal())
        }
if ($script:yesNoUserInput -eq 1) {
#updates fields in the document body
$document.Fields.Update()
#updates fields in footers and headers
$sectionCount = $document.Sections.Count
for ($t = 1; $t -le $sectionCount; $t++) {
$rangeHeader = $document.Sections.Item($t).Headers.Item(1).Range
$rangeHeader.Fields.Update()
$rangeFooter = $document.Sections.Item($t).Footers.Item(1).Range
$rangeFooter.Fields.Update()
}
#updates TOC
$tocCount = $document.TablesOfContents.Count
if ($tocCount -ge 1) {
$document.TablesOfContents.Item(1).Update()
$document.TablesOfContents.Item(1).UpdatePageNumbers()
}
}
Write-Host "End of document"
$document.Close()
}
Write-Host $files
$workbook.Close()
$excel.Quit()
$application.Quit()
