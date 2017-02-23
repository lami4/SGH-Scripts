$row = 1
$array= @("Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
#excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
#word
$application = New-Object -ComObject word.application
$application.Visible = $false
$document = $application.documents.open("C:\Users\Светлана\Desktop\Документ Microsoft Word (2).docx")
$properties = $document.BuiltInDocumentProperties
$binding = “System.Reflection.BindingFlags” -as [type]
$worksheet.Range("B:B").NumberFormat = "@"
foreach ($property in $properties) {
$pn = [System.__ComObject].InvokeMember(“name”,$binding::GetProperty,$null,$property,$null)
    trap [system.exception]
    {
    continue
    }
[string]$propertyValue = [System.__ComObject].InvokeMember(“value”,$binding::GetProperty,$null,$property,$null)
[string]$propertyName = $pn
    if ($propertyValue.Length -gt 0 -and $array -notcontains $propertyName) 
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
    if ($propertyValue.Length -gt 0 -and $array -notcontains $propertyName) 
    {
    Write-Host "$propertyName`: $propertyValue"
    $worksheet.Cells.Item($row, 1) = "$propertyName"
    $worksheet.Cells.Item($row, 2) = "$propertyValue"
    $worksheet.Cells.Item($row, 3) = $document.Name
    $row += 1
    }
}
$worksheet.Columns.Item("A").ColumnWidth = 45
$worksheet.Columns.Item("B").ColumnWidth = 45
$worksheet.Columns.Item("C").ColumnWidth = 45
$document.Close()
$application.quit()
$workbook.SaveAs("$PSScriptRoot\Properties.xlsx")
$excel.Quit()
