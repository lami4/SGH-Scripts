#functions
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

$selectedPath = Select-Folder
$row = 2
$listOfFiles = 2
$blacklist= @("Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
#excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
#word
$application = New-Object -ComObject word.application
$application.Visible = $false
$worksheet.Range("B:B").NumberFormat = "@"
$worksheet.Cells.Item(1, 1) = "Property name"
$worksheet.Cells.Item(1, 2) = "Property value"
$worksheet.Cells.Item(1, 3) = "Property holder"
$worksheet.Cells.Item(1, 4) = "Processed files"
$worksheet.Cells.Item(1, 5) = "Property type"
Get-ChildItem -Path "$selectedPath\*.*" -Include "*.doc*", "*.dot*" | % {
Write-Host "Taking properties from" $_.Name
$document = $application.documents.open($_.FullName)
$properties = $document.BuiltInDocumentProperties
$binding = “System.Reflection.BindingFlags” -as [type]
foreach ($property in $properties) {
$pn = [System.__ComObject].InvokeMember(“name”,$binding::GetProperty,$null,$property,$null)
    trap [system.exception]
    {
    continue
    }
[string]$propertyValue = [System.__ComObject].InvokeMember(“value”,$binding::GetProperty,$null,$property,$null)
[string]$propertyName = $pn
    if ($propertyValue.Length -gt 0 -and $blacklist -notcontains $propertyName) 
    {
    Write-Host "$propertyName`: $propertyValue"
    $worksheet.Cells.Item($row, 1) = "$propertyName"
    $worksheet.Cells.Item($row, 2) = "$propertyValue"
    $worksheet.Cells.Item($row, 3) = $document.Name
    $worksheet.Cells.Item($row, 5) = "B"
    $row += 1
    }
}
$customProperties = $document.CustomDocumentProperties
foreach($customProperty in $customProperties)
{
$pn = [System.__ComObject].InvokeMember(“name”,$binding::GetProperty,$null,$customProperty,$null)
    trap [system.exception]
    {
    continue
    }
[string]$propertyValue = [System.__ComObject].InvokeMember(“value”,$binding::GetProperty,$null,$customProperty,$null)
[string]$propertyName = $pn
    if ($propertyValue.Length -gt 0 -and $blacklist -notcontains $propertyName) 
    {
    Write-Host "$propertyName`: $propertyValue"
    $worksheet.Cells.Item($row, 1) = "$propertyName"
    $worksheet.Cells.Item($row, 2) = "$propertyValue"
    $worksheet.Cells.Item($row, 3) = $document.Name
    $worksheet.Cells.Item($row, 5) = "C"
    $row += 1
    }
}
$worksheet.Cells.Item($listOfFiles, 4) = $_.Name
$listOfFiles += 1
Write-Host "------End of document-----"
$document.Close()
}
$worksheet.Columns.Item("A").ColumnWidth = 20
$worksheet.Columns.Item("B").ColumnWidth = 60
$worksheet.Columns.Item("C").ColumnWidth = 40
$worksheet.Columns.Item("D").ColumnWidth = 40
$worksheet.Columns.Item("E").ColumnWidth = 20
$worksheet.Cells.Item(1, 1).Font.Bold = $true
$worksheet.Cells.Item(1, 2).Font.Bold = $true
$worksheet.Cells.Item(1, 3).Font.Bold = $true
$worksheet.Cells.Item(1, 4).Font.Bold = $true
$worksheet.Cells.Item(1, 5).Font.Bold = $true
$worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
$worksheet.Cells.Item(1, 2).HorizontalAlignment = -4108
$worksheet.Cells.Item(1, 3).HorizontalAlignment = -4108
$worksheet.Cells.Item(1, 4).HorizontalAlignment = -4108
$worksheet.Cells.Item(1, 5).HorizontalAlignment = -4108
$application.Quit()
$workbook.SaveAs("$PSScriptRoot\Properties.xlsx")
$workbook.Close()
$excel.Quit()
