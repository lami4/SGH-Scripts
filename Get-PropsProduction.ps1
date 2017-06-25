clear
$SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка"

Function Run-MSApplication ($AppName, $AppExtensions, $SelectedPath, $Text) {
   #Gets each file's extension in the folder specified by user.
   $ExtensionsInSelectedFolder = @(Get-ChildItem -Path $SelectedPath | % {$_.Extension})
   #If a file's extension from the folder specified by user matches an extension in the $AppExtensions array, opens the required application.
   foreach ($Extension in $ExtensionsInSelectedFolder) {
       if ($AppExtensions -contains $Extension) {
           Write-Host "$Text Application started"
           return New-Object -ComObject $AppName
           break
       }
   }
}

Function Extract-FileProperties ($BindingFlags, $CollectionOfProperties) {
    foreach ($Property in $CollectionOfProperties) {
        [array]$PropertyNames += [System.__ComObject].InvokeMember(“name”,$BindingFlags::GetProperty,$null,$Property,$null)
        trap [system.exception] {continue}
        [array]$PropertyValues += [System.__ComObject].InvokeMember(“value”,$BindingFlags::GetProperty,$null,$Property,$null)
    }
    return $PropertyNames, $PropertyValues
}


Function Output-CollectedPropertiesToExcelTable ($OutputWorksheet, $CollectedPropertiesData, $PropertyHolder, $PropertyType, $PropertyHolderExtension) {
    Write-Host "Processing $PropertyHolder..."
    for ($i = 0; $i -lt $CollectedPropertiesData[0].Length; $i++) {
        $OutputWorksheet.Cells.Item($script:RowOutputExcel, 1) = $CollectedPropertiesData[0][$i]
        $OutputWorksheet.Cells.Item($script:RowOutputExcel, 2) = $CollectedPropertiesData[1][$i]
        $OutputWorksheet.Cells.Item($script:RowOutputExcel, 3) = $PropertyHolder
        $OutputWorksheet.Cells.Item($script:RowOutputExcel, 4) = $PropertyType
        $OutputWorksheet.Cells.Item($script:RowOutputExcel, 5) = $PropertyHolderExtension
        $script:RowOutputExcel += 1
    }
}

Function Get-FileProperties ($SelectedPath) {
    #Adds the Office assembly to the current Windows PowerShell session.
    Add-type -AssemblyName Office
    #Stores the BindingFlags enumeration in $Binding.
    $Binding = “System.Reflection.BindingFlags” -as [type]
    #Extensions that will be processed by the script.
    $WordExtensions = @(".doc", ".docx", ".dotm")
    $ExcelExtensions = @(".xlsx", ".xls", ".xltm")
    $VisioExtensions = @(".vdx", ".vsd", ".vdw")
    $PowerPointExtensions = @(".pptx", ".ppt")
    #Starts MS Word if necessary.
    $Word = Run-MSApplication -AppName "Word.Application" -AppExtensions $WordExtensions -SelectedPath $SelectedPath -Text "Word" 
    #If MS Word is started, makes it not visible to user.
    if ($Word -ne $null) {$Word.Visible = $false}
    #Starts MS Excel if necessary
    $Excel = Run-MSApplication -AppName "Excel.Application" -AppExtensions $ExcelExtensions -SelectedPath $SelectedPath -Text "Excel"
    #If MS Excel is started, makes it not visible to user.
    if ($Excel -ne $null) {$Excel.Visible = $false}
    #Starts MS Visio if necessary and makes it not visible to user.
    $Visio = Run-MSApplication -AppName "Visio.InvisibleApp" -AppExtensions $VisioExtensions -SelectedPath $SelectedPath -Text "Visio"
    #Starts MS PowerPoint if necessary. Invisibility is enabled when opening a presentation.
    $PowerPoint = Run-MSApplication -AppName "PowerPoint.Application" -AppExtensions $PowerPointExtensions -SelectedPath $SelectedPath -Text "PowerPoint" 
    #Opens another instance of Excel application, creates a workbook and activates the first sheet to output data to.
    $OutputExcel = New-Object -ComObject Excel.Application
    $OutputExcel.Visible = $false
    $OutputWorkbook = $OutputExcel.Workbooks.Add()
    $OutputWorksheet = $OutputWorkbook.Worksheets.Item(1)
    #Turns 'Text' data type for all columns where the output data will be kept.
    $OutputWorksheet.Range("A:A").NumberFormat = "@"
    $OutputWorksheet.Range("B:B").NumberFormat = "@"
    $OutputWorksheet.Range("C:C").NumberFormat = "@"
    $OutputWorksheet.Range("D:D").NumberFormat = "@"
    $OutputWorksheet.Range("E:E").NumberFormat = "@"
    #Create headers for columns.
    $OutputWorksheet.Cells.Item(1, 1) = "Property name"
    $OutputWorksheet.Cells.Item(1, 2) = "Property value"
    $OutputWorksheet.Cells.Item(1, 3) = "Property holder"
    $OutputWorksheet.Cells.Item(1, 4) = "Property type"
    $OutputWorksheet.Cells.Item(1, 5) = "Property holder extension"
    #Does a bit of formatting :)
    $OutputWorksheet.Columns.Item("A").ColumnWidth = 60
    $OutputWorksheet.Columns.Item("B").ColumnWidth = 60
    $OutputWorksheet.Columns.Item("C").ColumnWidth = 60
    $OutputWorksheet.Columns.Item("D").ColumnWidth = 20
    $OutputWorksheet.Columns.Item("E").ColumnWidth = 40
    $OutputWorksheet.Cells.Item(1, 1).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, 2).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, 3).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, 4).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, 5).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
    $OutputWorksheet.Cells.Item(1, 2).HorizontalAlignment = -4108
    $OutputWorksheet.Cells.Item(1, 3).HorizontalAlignment = -4108
    $OutputWorksheet.Cells.Item(1, 4).HorizontalAlignment = -4108
    $OutputWorksheet.Cells.Item(1, 5).HorizontalAlignment = -4108
    #All the data will be written into the table starting from this row (first row contains the column headers).
    $script:RowOutputExcel = 2 
    #Starts looping through each file in the folder specified by user.
    Get-ChildItem -Path $SelectedPath | % {
        #if extension of the processed file matches an extension in $ExcelExtensions array, the script will open it and extract its properties using Excel application.
        if ($ExcelExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted.
            $Workbook = $Excel.Workbooks.Open($_.FullName)
            #Gets a collection of properties and puts it in $FileProperties.
            $FileProperties = $Workbook.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            #Outputs data kept in $CollectedPropertiesData array to the Excel file
            Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            #Closes active workbook without saving it.
            $Workbook.Close()
        }
        #if extension of the processed file matches an extension in $WordExtensionss array, the script will open it and extract its properties using Word application.
        if ($WordExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted.
            $Document = $Word.Documents.Open($_.FullName)
            #Gets a collection of properties and puts it in $FileProperties.
            $FileProperties = $Document.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            #Outputs data kept in $CollectedPropertiesData array to the Excel file
            Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            #Closes active document without saving it.
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions array, the script will open it and extract its properties using Visio application.
        if ($VisioExtensions -contains ($_.Extension).ToLower()) {
            #Opens a document.
            $Document = $Visio.Documents.Open($_.FullName)
            #List of Built In Document Properties.
            $VisioDocumentBuiltInProperties = @("Subject", "Title", "Creator", "Manager", "Company", "Language", "Category", "Keywords", "Description", "HyperlinkBase", "TimeCreated", "TimeEdited", "TimePrinted", "TimeSaved", "Stat", "Version")
            #$VisioDocumentBuiltInProperties | % {$Document.$_}
            #Closes active document without saving.
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions array, the script will open it and extract its properties using Visio application.
        if ($PowerPointExtensions -contains ($_.Extension).ToLower()) {
            #Opens a presentation and makes it not visible to the user.
            $Presentation = $PowerPoint.Presentations.Open($_.FullName, $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            #Gets a collection of properties and puts it in $FileProperties.
            $FileProperties = $Presentation.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            #Outputs data kept in $CollectedPropertiesData array to the Excel file
            Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            #Closes active presentation without saving it.
            $Presentation.Close()
        }
    }
    #Saves the Excel file that keeps the output data and kills all started MS Office processes
    Write-Host "Closing opened MS Office applications... It may take some time..."
    $OutputWorkbook.SaveAs("$PSScriptRoot\Properties.xlsx")
    $OutputWorkbook.Close()
    Start-Sleep -Seconds 3
    Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
}
Get-FileProperties -SelectedPath $SelectedPath
