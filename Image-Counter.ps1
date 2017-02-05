#filter parameters
#in MS Word, width 112 = ~4 cm, height 20 = ~70 mm.
$width = 112
$height = 20

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

#script code
$path = Select-Folder -description "Укажите папку с файлами, в которых нужно подсичтать общее количество картинок."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.WorkBooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Cells.Item(1, 1) = "File name"
$worksheet.Cells.Item(1, 2) = "Images"
$row = 2
Get-ChildItem -Path "$path/*.*" -Include "*.doc*" | % {
    Write-Host $_.Name "is being processed"
    $counter = 0
    $document = $word.Documents.Open($_.FullName)
    $iShapes = $document.InlineShapes
        foreach ($iShape in $iShapes) {
            If ($iShape.Width -ge 112 -and $iShape.Height -ge 20) {
            $counter += 1
            }
        }
    $Shapes = $document.Shapes
        foreach ($Shape in $Shapes) {
            If ($Shape.Width -ge 112 -and $Shape.Height -ge 20) {
            $counter += 1
            }
        }
    $document.Close()
    $counter -= 1
        if ($counter -ge 1) {
        $worksheet.Cells.Item($row, 1) = $_.Name
        $worksheet.Cells.Item($row, 2) = $counter
        $row += 1
        }
    Write-Host $counter
}
$worksheet.Columns.Item("A").ColumnWidth = 45
$worksheet.Columns.Item("B").ColumnWidth = 10
$worksheet.Range("A1:B1").HorizontalAlignment = -4108
$worksheet.Columns.Item("B").HorizontalAlignment = -4108
$worksheet.Range("A1:B1").Font.Bold = $true
$worksheet.Cells.Item($row, 1) = "TOTAL:"
$worksheet.Cells.Item($row, 1).HorizontalAlignment = -4152
$worksheet.Cells.Item($row, 1).Font.Bold = $true
$formula = $row - 1
$worksheet.Cells.Item($row, 2).Formula = "=СУММ(B2:B$formula)"
$word.Quit()
