clear
$SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка"
$script:GetBuiltInProperties = $true
$script:GetCustomProperties = $true
$script:IgnorePropertiesWithNoValue = $false
$script:UseBlacklist = $true
$script:WhitelistEnabled = $true
$script:Blacklist = @("Title", "Author", "Subject", "Creation date")
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
            $PropertyName = [System.__ComObject].InvokeMember(“name”,$BindingFlags::GetProperty,$null,$Property,$null)
            trap [system.exception] {continue}
            $PropertyValue = [System.__ComObject].InvokeMember(“value”,$BindingFlags::GetProperty,$null,$Property,$null)
            if ($PropertyValue -eq $null) {$PropertyValue = ""}
            if ($script:UseBlacklist -eq $true -and $script:WhitelistEnabled -eq $true) {
                if ($script:Blacklist -contains $PropertyName) {
                    if ($script:IgnorePropertiesWithNoValue -eq $true -and $PropertyValue -eq "") {continue}
                    [array]$PropertyNames += $PropertyName 
                    [array]$PropertyValues += $PropertyValue 
                    continue  
                } else {
                    continue
                }
            } else {
                if ($script:UseBlacklist -eq $true -and $script:Blacklist -contains $PropertyName) {continue}  
                if ($script:IgnorePropertiesWithNoValue -eq $true -and $PropertyValue -eq "") {continue}
                [array]$PropertyNames += $PropertyName
                [array]$PropertyValues += $PropertyValue
            }
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

Function Format-ExcelTable ($OutputWorksheet, $ColumnLetter, $ColumnNumber, $Width, $Header) {
    $OutputWorksheet.Range("$ColumnLetter").NumberFormat = "@"
    $OutputWorksheet.Cells.Item(1, $ColumnNumber) = $Header
    $OutputWorksheet.Columns.Item("$ColumnLetter").ColumnWidth = $Width
    $OutputWorksheet.Cells.Item(1, $ColumnNumber).Font.Bold = $true
    $OutputWorksheet.Cells.Item(1, $ColumnNumber).HorizontalAlignment = -4108
}

Function Get-FileProperties ($SelectedPath) {
    #Adds the Office assembly to the current Windows PowerShell session.
    Add-type -AssemblyName Office
    #Stores the BindingFlags enumeration in $Binding.
    $Binding = “System.Reflection.BindingFlags” -as [type]
    #Extensions that will be processed by the script.
    $WordExtensions = @(".doc", ".docx", ".dotm")
    $ExcelExtensions = @(".xlsx", ".xls", ".xltm", ".xlsm")
    $VisioExtensions = @(".vdx", ".vsd", ".vdw")
    $PowerPointExtensions = @(".pptx", ".ppt", ".pptm", ".potx")
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
    $OutputExcel.Visible = $true
    $OutputWorkbook = $OutputExcel.Workbooks.Add()
    $OutputWorksheet = $OutputWorkbook.Worksheets.Item(1)
    #Turns 'Text' data type for all columns where the output data will be kept and does a bit of formatting :)
    Format-ExcelTable -OutputWorksheet $OutputWorksheet -ColumnLetter "A:A" -ColumnNumber 1 -Width 40 -Header "Property name"
    Format-ExcelTable -OutputWorksheet $OutputWorksheet -ColumnLetter "B:B" -ColumnNumber 2 -Width 70 -Header "Property value"
    Format-ExcelTable -OutputWorksheet $OutputWorksheet -ColumnLetter "C:C" -ColumnNumber 3 -Width 60 -Header "Property holder"
    Format-ExcelTable -OutputWorksheet $OutputWorksheet -ColumnLetter "D:D" -ColumnNumber 4 -Width 20 -Header "Property type"
    Format-ExcelTable -OutputWorksheet $OutputWorksheet -ColumnLetter "E:E" -ColumnNumber 5 -Width 40 -Header "Property holder extension"
    #All the data will be written into the table starting from this row (first row contains the column headers).
    $script:RowOutputExcel = 2 
    #Starts looping through each file in the folder specified by user.
    Get-ChildItem -Path $SelectedPath | % {
        #if extension of the processed file matches an extension in $ExcelExtensions array, the script will open it and extract its properties using Excel application.
        if ($ExcelExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted.
            $Workbook = $Excel.Workbooks.Open($_.FullName)
            #Extracts built-in properties if required
            if ($script:GetBuiltInProperties -eq $true) {
                #Gets a collection of built-in properties and puts it in $FileProperties.
                $FileProperties = $Workbook.BuiltInDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            }
            #Extracts custom properties if required
            if ($script:GetCustomProperties -eq $true) {
                #Gets a collection of custom properties and puts it in $FileProperties.
                $FileProperties = $Workbook.CustomDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "C" -PropertyHolderExtension $_.Extension
            }
            #Closes active workbook without saving it.
            $Workbook.Close()
        }
        #if extension of the processed file matches an extension in $WordExtensionss array, the script will open it and extract its properties using Word application.
        if ($WordExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted.
            $Document = $Word.Documents.Open($_.FullName)
            #Extracts built-in properties if required
            if ($script:GetBuiltInProperties -eq $true) {
                #Gets a collection of built-in properties and puts it in $FileProperties.
                $FileProperties = $Document.BuiltInDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            }
            #Extracts custom properties if required
            if ($script:GetCustomProperties -eq $true) {
                #Gets a collection of custom properties and puts it in $FileProperties.
                $FileProperties = $Document.CustomDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "C" -PropertyHolderExtension $_.Extension
            }
            #Closes active document without saving it.
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions array, the script will open it and extract its properties using Visio application.
        if ($VisioExtensions -contains ($_.Extension).ToLower()) {
            #Opens a document.
            $Document = $Visio.Documents.Open($_.FullName)
            #List of Built In Document Properties.
            $CollectedPropertiesData = @(), @()
            $VisioDocumentBuiltInPropertyNames = @("Subject", "Title", "Creator", "Manager", "Company", "Language", "Category", "Keywords", "Description", "HyperlinkBase", "TimeCreated", "TimeEdited", "TimePrinted", "TimeSaved", "Stat", "Version")
            foreach ($VisioPropertyName in $VisioDocumentBuiltInPropertyNames) {
                if ($script:UseBlacklist -eq $true -and $script:WhitelistEnabled -eq $true) {
                    if ($script:Blacklist -contains $VisioPropertyName) {
                        if ($script:IgnorePropertiesWithNoValue -eq $true -and $Document.$VisioPropertyName -eq "") {continue}
                        $CollectedPropertiesData[0] += $VisioPropertyName 
                        $CollectedPropertiesData[1] += $Document.$VisioPropertyName 
                        continue  
                    } else {
                        continue
                    }
                } else {
                    if ($script:UseBlacklist -eq $true -and $script:Blacklist -contains $VisioPropertyName) {continue}  
                    if ($script:IgnorePropertiesWithNoValue -eq $true -and $Document.$VisioPropertyName -eq "") {continue}
                        $CollectedPropertiesData[0] += $VisioPropertyName 
                        $CollectedPropertiesData[1] += $Document.$VisioPropertyName 
                }  
            }
            Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            #Closes active document without saving.
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions array, the script will open it and extract its properties using Visio application.
        if ($PowerPointExtensions -contains ($_.Extension).ToLower()) {
            #Opens a presentation and makes it not visible to the user.
            $Presentation = $PowerPoint.Presentations.Open($_.FullName, $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            #Extracts built-in properties if required
            if ($script:GetBuiltInProperties -eq $true) {
                #Gets a collection of properties and puts it in $FileProperties.
                $FileProperties = $Presentation.BuiltInDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            }
            #Extracts custom properties if required
            if ($script:GetCustomProperties -eq $true) {
                #Gets a collection of custom properties and puts it in $FileProperties.
                $FileProperties = $Presentation.CustomDocumentProperties
                #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array.
                $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
                #Outputs data kept in $CollectedPropertiesData array to the Excel file
                Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "C" -PropertyHolderExtension $_.Extension
            }
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
