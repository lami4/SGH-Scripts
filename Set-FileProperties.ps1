clear
$files = @()
Function Select-File {
Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = "C:\Users\Светлана\Desktop"
$f.Filter = "MS Excel Files (*.xls*)|*.xls*|All Files (*.*)|*.*"
$f.ShowHelp = $false
$f.Multiselect = $true
$show = $f.ShowDialog()
If ($show -eq "OK") {if ($f.Multiselect) { $f.FileNames } else { $f.FileName }} else {Exit}
}
$selectedFile = Select-File
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
for ($i = 0; $i -le $files.Count; $i++) {
    $range = $worksheet.Range("C2:C$lastNonemptyCellInColumn")
    $target = $range.Find($files[$i])
        if ($target -eq $null) {
        Write-Host "No match found"
        } else {
        $firstHit = $target
        Do
        {
        $changeAddress = $target.AddressLocal($false, $false) -replace "E", ""
        [int]$valueInCell = $sheet.Cells.Item($changeAddress, "J").Value()
        if ($valueInCell.Count -eq 0) {$changes += 0} else {$changes += $valueInCell}
        $target = $range.FindNext($target)
        }
    While ($target -ne $NULL -and $target.AddressLocal() -ne $firstHit.AddressLocal())
    $greatestValue = $changes | Measure-Object -Maximum
    Write-Host $changes
    Write-Host $greatestValue.Maximum
    }
}
Write-Host $files
$workbook.Close()
$excel.Quit()
