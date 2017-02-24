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
Get-ChildItem -Path "$selectedPath\*.*" -Include "*.doc*" | % {
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
    $row += 1
    }
}
$worksheet.Cells.Item($listOfFiles, 4) = $_.Name
$listOfFiles += 1
Write-Host "------End of document-----"
$document.Close()
}
$worksheet.Columns.Item("A").ColumnWidth = 45
$worksheet.Columns.Item("B").ColumnWidth = 45
$worksheet.Columns.Item("C").ColumnWidth = 45
$application.quit()
$workbook.SaveAs("$PSScriptRoot\Properties.xlsx")
$workbook.Close()
$excel.Quit()
<#
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Open("C:\Users\Светлана\Desktop\Properties.xlsx")
$worksheet = $workbook.Worksheets.Item(1)
$xlup = -4121
$kek = $worksheet.Range("B2").End($xlup).row
Write-Host $kek
$workbook.Close()
$excel.Quit()
#>
