clear
#opens Excel application and goes to Sheet 1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.WorkBooks.Open("Z:\OTD.Translate\Переводчики\PABKRF\sandbox\test2\1.xlsx")
$worksheet = $workbook.Worksheets.Item(1)
#creates workbook and goes to sheet 1 to keep the parsed strings
$WorkBookForParsedData = $excel.WorkBooks.Add()
$WorkSheetForParsedData = $WorkBookForParsedData.Worksheets.Item(1)
#last populated cell in column
$LastPopulatedtCellInColumn = $worksheet.Cells.Item($worksheet.Rows.Count, 3).End(-4162).Row
#Row counter
$ExecutionTime = Measure-Command {
$RowNumber = 1
for ($i = 2; $i -le $LastPopulatedtCellInColumn; $i++) {
    $ValueInCell = $worksheet.Cells.Item($i, 3).Value()
    if ($ValueInCell -ne "" -and $ValueInCell.Length -gt 0) {
    Write-Host $ValueInCell
    if ($ValueInCell -match [char]10) {
        Write-Host "Cell $i contains a line breaker"
        $ParsableValue = $ValueInCell -replace $([char]10 + [char]10), [char]10
        $ParsedCell = $ParsableValue -split [char]10
        Write-Host $ParsedCell.Length
        for ($t = 0; $t -lt $ParsedCell.length; $t++) {
            if ($ParsedCell[$t] -ne "") {
                Write-Host $ParsedCell[$t]
                #creates ID of the parsed string
                $IdForParsedString = "$i#$t"
                Write-Host $IdForParsedString
                #adds the parsed string to the Excel file that keeps pasrsed text
                $WorkSheetForParsedData.Cells.Item($RowNumber, 1) = $ParsedCell[$t]
                #adds the parsed string's UID to the Excel file that keeps parsed text
                $WorkSheetForParsedData.Cells.Item($RowNumber, 2) = $IdForParsedString
                #adds the ID of the parsed string to a string template
                [regex]$pattern = "$([Regex]::Escape($ParsedCell[$t]))"
                $ValueInCell = $pattern.Replace("$ValueInCell", "$IdForParsedString", 1)
                $RowNumber += 1
            }
        }
        $worksheet.Cells.Item($i, 3) = $ValueInCell
    } else {
        Write-Host "Cell $i does not contain a line breaker"
        #creates ID of the string
        $IdForString = "$i#1"
        #adds the string to the Excel file that keeps pasrsed text
        $WorkSheetForParsedData.Cells.Item($RowNumber, 1) = $ValueInCell
        #adds the string's UID to the Excel file that keeps parsed text
        $WorkSheetForParsedData.Cells.Item($RowNumber, 2) = $IdForString
        #adds ID to the source file
        $worksheet.Cells.Item($i, 3) = $IdForString
        $RowNumber += 1
    }
    }   
}
}
#closes the workbook and Excel application
$WorkBookForParsedData.SaveAs("Z:\OTD.Translate\Переводчики\PABKRF\sandbox\test2\2.xlsx") | Out-Null
$WorkBookForParsedData.Close($false) | Out-Null
$workbook.Close($true) | Out-Null
$excel.Quit() | Out-Null
Write-Host $ExecutionTime
