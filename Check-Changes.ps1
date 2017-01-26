$changes = @()
$excel = new-object -com excel.application
$excel.Visible = $true
$book = $excel.Workbooks.Open("Z:\OTD.Translate\Учет программ и ПД.xls")
$sheet = $book.WorkSheets.Item("Лист1")
$range = $sheet.Range("E:E")
$target = $range.Find("PABKRF-GL-RU-00.00.00.dRNT.01.00")
if ($target -eq $null) {
    Write-Host "No match found"
    } else {
    $firstHit = $target
    Do
    {
        $changeAddress = $target.AddressLocal($false, $false) -replace "E", ""
        [int]$valueInCell = $sheet.Cells.Item($changeAddress, "J").Value()
        if ($valueInCell.Count -eq 0) {$changes += 0} else {$changes += $valueInCell}
        $target = $Range.FindNext($target)
    }
    While ($target -ne $NULL -and $target.AddressLocal() -ne $firstHit.AddressLocal())
    $greatestValue = $changes | Measure-Object -Maximum
    Write-Host $changes
    Write-Host $greatestValue.Maximum
}
