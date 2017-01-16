clear
#Functions
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

Function Check-DocumentNameInSpecification ($selectedFolder, $currentSpecification) {
    $fileExtension = [IO.Path]::GetExtension($currentSpecification)
    #фильтр для excel файла
    if ($fileExtension -eq ".xlsx" -or $fileExtension -eq ".xls") {
    Write-Host "Excel файл. Требуется ручная проверка"
    } else {
        #Write-Host "$selectedFOlder\$currentSpecification"
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open("$selectedFolder\$currentSpecification")
        #написать проверку на количество таблиц в спецификации
        [int]$tableCount = $document.Tables.Count
        if ($tableCount -eq 0) {
        Write-Host "$currentSpecification не содержит таблиц."
        } elseif ($tableCount -gt 1) {
        Write-Host "$currentSpecification содержит несколько таблиц."
        } else {
        [int]$rowCount = try {$document.Tables.Item(1).Rows.Count + 1} catch {""}
        Write-Host "$currentSpecification : $rowCount"
            for ($i = 1; $i -lt $rowCount; $i++) {
            [string]$valueInDocumentNameCell = ((($document.Tables.Item(1).Cell($i,4).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                #добавить подсчет совпадений и вывод полученного значения в статистику (Ссылается на <количество документов>)
                if ($valueInDocumentNameCell -match '\b(.{13})\d\d\.\d\d\.\d\d\.(.{4})\.\d\d\.\d\d([^\s]*)') {
                    if ($valueInDocumentNameCell -notmatch '\b([A-Z]{6})-([A-Z]{2})-([A-Z]{2})-\d\d\.\d\d\.\d\d\.([a-z]{1})([A-Z]{3})\.\d\d\.\d\d([^\s]*)') {
                    Write-Host "Обозначение содержит русские буквы или недопустимые символы."
                    } else {
                    Write-Host "Обозначение соответствует маске"
                    }
                } else {
                [string]$valueInMd5Cell = ((($document.Tables.Item(1).Cell($i,7).Range.Text).Trim([char]0x0007)) -replace '\s+', ' ').Trim(' ')
                    if ($valueInMd5Cell -match '([m,M]\s*[d,D]\s*5)') {
                        if ($valueInMd5Cell -notmatch '([m,M]\s*[d,D]\s*5)\s*:') {
                        Write-Host "Ячейка с MD5 оформлена некорректно. Отсутствует разделитель."
                        } else {
                        Write-Host "Ячейка с MD5 суммой формлена корректно"
                    }
                    #сделать еще проверку - если есть нет мд5, но есть маска самой суммы неправильно оформлена ячейка
                    #добавить подсчет файлов и вывод полученного значения в статистику (Ссылается на <количество документов>)
                    } else {
                        if ($valueInMd5Cell -match '[a-zA-Z0-9]{32}') {
                        Write-Host "Ячейка с MD5 оформлена некорректно. Отсутствует ключ md5"
                        }
                    }
                }
            }
        }
        #checks values in cells
        [string]$valueInDocVersionCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 1).Range.Text} catch {""}
        [string]$valueInNotificationNoCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 3).Range.Text} catch {""}
        [string]$valueInDocNameCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(1, 6).Range.Text} catch {""}
        [string]$valueInDocTitleCell = try {$document.Sections.Item(1).Footers.Item(2).Range.Tables.Item(1).Cell(4, 5).Range.Text} catch {""}
            if ($valueInDocVersionCell.Length -eq 0 -or $valueInNotificationNoCell.Length -eq 0 -or $valueInDocNameCell.Length -eq 0 -or $valueInDocTitleCell.Length -eq 0) {
            Write-Host "Невозможно получить значения из штампа. Штамп либо отсутствует, либо неверно заверстан."
            } else {
            Write-Host "Значения из штампа получены."
        }
        $document.Close()
        $word.Quit()
    Write-Host "-------End of document-------"   
    }
}


$pathToFolder = Select-Folder -description "Выберите папку, в которой нужно соответствие требованиям."
Measure-Command {
Get-ChildItem -Path $pathToFolder -Exclude "*.pdf" | Where-Object  {$_.BaseName -match 'SPC'} | % {
Check-DocumentNameInSpecification -selectedFolder $pathToFolder -currentSpecification $_.Name
}
}
