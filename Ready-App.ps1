clear
###GET FILE PROPERTIES FUNCTIONS###
Function Run-MSApplication ($AppName, $AppExtensions, $Text, $PathToFolder) {
   #Gets each file's extension in the folder specified by user.
   $ExtensionsInSelectedFolder = @(Get-ChildItem -Path $PathToFolder | % {$_.Extension})
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
        #Gets property name and its value
        $PropertyName = [System.__ComObject].InvokeMember(“name”,$BindingFlags::GetProperty,$null,$Property,$null)
        trap [system.exception] {continue}
        $PropertyValue = [System.__ComObject].InvokeMember(“value”,$BindingFlags::GetProperty,$null,$Property,$null)
        #If the propery value has a $null value, converts it into an empty string
        if ($PropertyValue -eq $null) {$PropertyValue = ""}
        #Goes to this scenario, if user has enabled the whitelist
        if ($script:UseBlacklist -eq $true -and $script:WhitelistEnabled -eq $true) {
            #Checks if the current property name is on the whitelist
            if ($script:GetPropertiesBlacklist -contains $PropertyName) {
                #If this checkout enabled, checks if the property value is empty. If the value is an empty string, moves on to the next item in $CollectionOfProperties.
                if ($script:IgnorePropertiesWithNoValue -eq $true -and $PropertyValue -eq "") {continue}
                #Adds property name and value to arrays
                [array]$PropertyNames += $PropertyName 
                [array]$PropertyValues += $PropertyValue
                #Moves on to the next item in $CollectionOfProperties
                continue  
            } else {
                #If the current property name is not on the whitelist, moves on to the next item in $CollectionOfProperties
                continue
            }
        #Goes to this scenario, if user has enabled the blacklist
        } else {
            #Checks if the current property name is on the blacklist
            if ($script:UseBlacklist -eq $true -and $script:GetPropertiesBlacklist -contains $PropertyName) {continue}  
            #If this checkout enabled, checks if the property value is empty. If the value is an empty string, moves on to the next item in $CollectionOfProperties.
            if ($script:IgnorePropertiesWithNoValue -eq $true -and $PropertyValue -eq "") {continue}
            #Adds property name and value to arrays
            [array]$PropertyNames += $PropertyName
            [array]$PropertyValues += $PropertyValue
        }
    }
    #Returns collected data
    return $PropertyNames, $PropertyValues
}


Function Output-CollectedPropertiesToExcelTable ($OutputWorksheet, $CollectedPropertiesData, $PropertyHolder, $PropertyType, $PropertyHolderExtension) {
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

Function Get-FileProperties {
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
    $Word = Run-MSApplication -AppName "Word.Application" -AppExtensions $WordExtensions -Text "Word" -PathToFolder $script:GetPropertyPathToSelectedFolder
    #If MS Word is started, makes it not visible to user.
    if ($Word -ne $null) {$Word.Visible = $false}
    #Starts MS Excel if necessary
    $Excel = Run-MSApplication -AppName "Excel.Application" -AppExtensions $ExcelExtensions -Text "Excel" -PathToFolder $script:GetPropertyPathToSelectedFolder
    #If MS Excel is started, makes it not visible to user.
    if ($Excel -ne $null) {$Excel.Visible = $false}
    #Starts MS Visio if necessary and makes it not visible to user.
    $Visio = Run-MSApplication -AppName "Visio.InvisibleApp" -AppExtensions $VisioExtensions -Text "Visio" -PathToFolder $script:GetPropertyPathToSelectedFolder
    #Starts MS PowerPoint if necessary. Invisibility is enabled when opening a presentation.
    $PowerPoint = Run-MSApplication -AppName "PowerPoint.Application" -AppExtensions $PowerPointExtensions -Text "PowerPoint" -PathToFolder $script:GetPropertyPathToSelectedFolder
    #Creates another instance of Excel application, adds a workbook and activates the first sheet to output data to.
    $OutputExcel = New-Object -ComObject Excel.Application
    $OutputExcel.Visible = $false
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
    Get-ChildItem -Path $script:GetPropertyPathToSelectedFolder | % {
        Write-Host "Extracting properties from $($_.Name)..."
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
            $DocumentVisio = $Visio.Documents.Open($_.FullName)
            #List of Built In Document Properties.
            $CollectedPropertiesData = @(), @()
            $VisioDocumentBuiltInPropertyNames = @("Subject", "Title", "Creator", "Manager", "Company", "Language", "Category", "Keywords", "Description", "HyperlinkBase", "TimeCreated", "TimeEdited", "TimePrinted", "TimeSaved", "Stat", "Version")
            foreach ($VisioPropertyName in $VisioDocumentBuiltInPropertyNames) {
                if ($script:UseBlacklist -eq $true -and $script:WhitelistEnabled -eq $true) {
                    if ($script:GetPropertiesBlacklist -contains $VisioPropertyName) {
                        if ($script:IgnorePropertiesWithNoValue -eq $true -and $DocumentVisio.$VisioPropertyName -eq "") {continue}
                        $CollectedPropertiesData[0] += $VisioPropertyName 
                        $CollectedPropertiesData[1] += $DocumentVisio.$VisioPropertyName 
                        continue  
                    } else {
                        continue
                    }
                } else {
                    if ($script:UseBlacklist -eq $true -and $script:GetPropertiesBlacklist -contains $VisioPropertyName) {continue}  
                    if ($script:IgnorePropertiesWithNoValue -eq $true -and $DocumentVisio.$VisioPropertyName -eq "") {continue}
                    $CollectedPropertiesData[0] += $VisioPropertyName 
                    $CollectedPropertiesData[1] += $DocumentVisio.$VisioPropertyName 
                }  
            }
            Output-CollectedPropertiesToExcelTable -OutputWorksheet $OutputWorksheet -CollectedPropertiesData $CollectedPropertiesData -PropertyHolder $_.Name -PropertyType "B" -PropertyHolderExtension $_.Extension
            #Closes active document without saving.
            $DocumentVisio.Close()
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
###GET FILE PROPERTIES FUNCTIONS###

###SET FILE PROPERTIES FUNCTIONS###
Function Find-PropertiesForFile ($LastNonEmptyCell, $WorksheetWithProperties, $FileName) {
    $RangeToSearchThrough = $WorksheetWithProperties.Range("C2:C$LastNonEmptyCell")
    $Target = $RangeToSearchThrough.Find($FileName, [Type]::Missing, [Type]::Missing, 1)
        if ($Target -eq $null) {
        Write-Host "Properties.xlsx has no properties for $FileName. Moving on to the next document..."
        return $null
        } else {
            $FirstMatch = $Target
            Do {
            $CurrentAddress = $Target.AddressLocal($false, $false) -replace "C", ""
            [array]$PropertyNames += $WorksheetWithProperties.Cells.Item($CurrentAddress, "A").Value()
            [array]$PropertyValues += $WorksheetWithProperties.Cells.Item($CurrentAddress, "B").Value()
            [array]$PropertyTypes += $WorksheetWithProperties.Cells.Item($CurrentAddress, "D").Value()
            $Target = $RangeToSearchThrough.FindNext($Target)
            } While ($Target.AddressLocal() -ne $FirstMatch.AddressLocal())
        }
    Write-Host "Properties.xlsx has property(ies) for $FileName."
    return $PropertyNames, $PropertyValues, $PropertyTypes
}

Function Update-PropertiesInFile ($CollectionOfProperties, $PropertiesToBeUpdated, $Binding) {
    for ($t = 0; $t -lt $PropertiesToBeUpdated[0].Length; $t++) {
        $PropertyToBeUpdated = @{Name = $PropertiesToBeUpdated[0][$t]; Value = $PropertiesToBeUpdated[1][$t]}
        $InvokedProperty = [System.__ComObject].InvokeMember(“item”,$Binding::GetProperty,$null,$CollectionOfProperties,$PropertyToBeUpdated.Name)
        [System.__ComObject].InvokeMember(“value”,$Binding::SetProperty,$null,$InvokedProperty,$PropertyToBeUpdated.Value)
        Write-Host "Updated property: $($PropertyToBeUpdated.Name)." "New value: $($PropertyToBeUpdated.Value)" -ForegroundColor DarkGreen
    }
}

Function Filter-PropertiesByType ($FoundProperties, $RequiredPropertyType) {
    for ($t = 0; $t -lt $FoundProperties[2].Length; $t++) {
        if ($FoundProperties[2][$t] -eq "$RequiredPropertyType") {
        [array]$FilteredNames += $FoundProperties[0][$t]
        [array]$FilteredValues += $FoundProperties[1][$t]
        }
    }
    return $FilteredNames, $FilteredValues
}

Function Close-SavedDocument ($DocumentObject) {
$DocumentObject.Saved = $false
$DocumentObject.Save()
$DocumentObject.Close()
}

Function Set-FileProperties () {
    #Stores the BindingFlags enumeration in $Binding.
    $Binding = “System.Reflection.BindingFlags” -as [type]
    #Extensions that will be processed by the script.
    $WordExtensions = @(".doc", ".docx", ".dotm")
    $ExcelExtensions = @(".xlsx", ".xls", ".xltm", ".xlsm")
    $VisioExtensions = @(".vdx", ".vsd", ".vdw")
    $PowerPointExtensions = @(".pptx", ".ppt", ".pptm", ".potx")
    #Starts MS Word if necessary.
    $Word = Run-MSApplication -AppName "Word.Application" -AppExtensions $WordExtensions -Text "Word" -PathToFolder $script:SetPropertyPathToSelectedFolder
    #If MS Word is started, makes it not visible to user.
    if ($Word -ne $null) {$Word.Visible = $false}
    #Starts MS Excel if necessary
    $Excel = Run-MSApplication -AppName "Excel.Application" -AppExtensions $ExcelExtensions -Text "Excel" -PathToFolder $script:SetPropertyPathToSelectedFolder
    #If MS Excel is started, makes it not visible to user.
    if ($Excel -ne $null) {$Excel.Visible = $false}
    #Starts MS Visio if necessary and makes it not visible to user.
    $Visio = Run-MSApplication -AppName "Visio.InvisibleApp" -AppExtensions $VisioExtensions -Text "Visio" -PathToFolder $script:SetPropertyPathToSelectedFolder
    #Starts MS PowerPoint if necessary. Invisibility is enabled when opening a presentation.
    $PowerPoint = Run-MSApplication -AppName "PowerPoint.Application" -AppExtensions $PowerPointExtensions -Text "PowerPoint" -PathToFolder $script:SetPropertyPathToSelectedFolder
    #Creates another instance of Excel application, opens Excel file that contains the properties and activates the first sheet.
    $ExcelWithProperties = New-Object -ComObject Excel.Application
    $ExcelWithProperties.Visible = $false
    $WorkbookWithProperties = $ExcelWithProperties.WorkBooks.Open($script:SetPropertyPathToSelectedFile)
    $WorksheetWithProperties = $WorkbookWithProperties.Worksheets.Item(1)
    #Finds last non empty cell in column C. This value will be used later to create a search range in Excel.
    $LastNonemptyCellInColumn = $WorksheetWithProperties.Range("C:C").End(-4121).Row
    Write-Host "Getting ready..."
    Start-Sleep -Seconds 2
    Get-ChildItem -Path $script:SetPropertyPathToSelectedFolder | % {
        #Finds properties for the file being processed and puts them into array
        $FoundProperties = Find-PropertiesForFile -LastNonEmptyCell $LastNonemptyCellInColumn -WorksheetWithProperties $WorksheetWithProperties -FileName $_.Name
        #If $WorksheetWithProperties contains any properties (i.e. $FoundProperties is not equal to $null) for the file being processed, the script will go on to update them. Otherwise, the script will move to the next document.
        if ($FoundProperties -ne $null) {
                #if extension of the file being processed matches an extension in $ExcelExtensions array, the script will open it and update avaiable properties using Excel application
                if ($ExcelExtensions -contains ($_.Extension).ToLower()) {
                    #Opens the file whose properties are to be updated.
                    $Workbook = $Excel.Workbooks.Open($_.FullName)
                    #Updates built-in properties if required
                    if ($script:SetBuiltInProperties -eq $true) {
                        Write-Host "Updating built-in properties..."
                        #Filters $FoundProperties array to get only B (abbreviation for BuiltInProperties) type properties
                        [array]$FoundBuiltInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "B"
                        Write-Host $FoundBuiltInProperties
                        #If no BuiltInProperties were found, the script simply moves on to update CustomProperties
                        if ($FoundBuiltInProperties -eq $null) {
                            Write-Host "No built-in properties for this file" -ForegroundColor Red
                        } else {
                            #If some BuiltInProperties were found, the script will update them
                            Update-PropertiesInFile -CollectionOfProperties $Workbook.BuiltInDocumentProperties -PropertiesToBeUpdated $FoundBuiltInProperties -Binding $Binding
                        } 
                    }
                    #Updates custom properties if required
                    if ($script:SetCustomProperties -eq $true) {
                        Write-Host "Updating custom properties..."
                        #Filters $FoundProperties array to get only C (abbreviation for CustomProperties) type properties
                        [array]$FoundCustomInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "C"
                        #If no CustomProperties were found, the script simply moves on to update properties in the next document
                        if ($FoundCustomInProperties -eq $null) {
                            Write-Host "No custom properties for this file" -ForegroundColor Red
                        } else {
                            #If some BuiltInProperties were found, the script will update them
                            Update-PropertiesInFile -CollectionOfProperties $Workbook.CustomDocumentProperties -PropertiesToBeUpdated $FoundCustomInProperties -Binding $Binding
                        }
                    }
                    Close-SavedDocument -DocumentObject $Workbook
                }
            #if extension of the file being processed matches an extension in $WordExtensions array, the script will open it and update avaiable properties using Word application
            if ($WordExtensions -contains ($_.Extension).ToLower()) {
                #Opens the file whose properties are to be updated
                $Document = $Word.Documents.Open($_.FullName)
                #Updates built-in properties if required
                if ($script:SetBuiltInProperties -eq $true) {
                    Write-Host "Updating built-in properties..."
                    #Filters $FoundProperties array to get only B (abbreviation for BuiltInProperties) type properties
                    [array]$FoundBuiltInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "B"
                    #If no BuiltInProperties were found, the script simply moves on to update CustomProperties
                    if ($FoundBuiltInProperties -eq $null) {
                        Write-Host "No built-in properties for this file" -ForegroundColor Red
                    } else {
                        #If some BuiltInProperties were found, the script will update them
                        Update-PropertiesInFile -CollectionOfProperties $Document.BuiltInDocumentProperties -PropertiesToBeUpdated $FoundBuiltInProperties -Binding $Binding
                    }
                }
                #Updates custom properties if required
                if ($script:SetCustomProperties -eq $true) {
                    Write-Host "Updating custom properties..."
                    #Filters $FoundProperties array to get only C (abbreviation for CustomProperties) type properties
                    [array]$FoundCustomInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "C"
                    #If no CustomProperties were found, the script simply moves on to update properties in the next document
                    if ($FoundCustomInProperties -eq $null) {
                        Write-Host "No custom properties for this file" -ForegroundColor Red
                    } else {
                        #If some BuiltInProperties were found, the script will update them
                        Update-PropertiesInFile -CollectionOfProperties $Document.CustomDocumentProperties -PropertiesToBeUpdated $FoundCustomInProperties -Binding $Binding
                    }
                }
                Close-SavedDocument -DocumentObject $Document
            }
            #if extension of the file being processed matches an extension in $VisioExtensions array, the script will open it and update avaiable properties using Visio application
            if ($VisioExtensions -contains ($_.Extension).ToLower()) {
                #Opens the file whose properties are to be updated
                $DocumentVisio = $Visio.Documents.Open($_.FullName)
                Write-Host "Updating built-in properties..."
                for ($t = 0; $t -lt $FoundProperties[0].Length; $t++) {
                    $VisioPropertyName = $FoundProperties[0][$t]
                    $VisioPropertyNewValue = $FoundProperties[1][$t]
                    $DocumentVisio.$VisioPropertyName = $VisioPropertyNewValue
                    Write-Host "Updated property: $VisioPropertyName." "New value: $VisioPropertyNewValue" -ForegroundColor DarkGreen
                }
                Close-SavedDocument -DocumentObject $DocumentVisio   
            }
            #if extension of the file being processed matches an extension in $PowerPointExtensions array, the script will open it and update avaiable properties using PowerPoint application
            if ($PowerPointExtensions -contains ($_.Extension).ToLower()) {
            #Opens a presentation and makes it not visible to the user
            $Presentation = $PowerPoint.Presentations.Open($_.FullName, $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
                #Updates built-in properties if required
                if ($script:SetBuiltInProperties -eq $true) {
                    Write-Host "Updating built-in properties..."
                    #Filters $FoundProperties array to get only B (abbreviation for BuiltInProperties) type properties
                    [array]$FoundBuiltInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "B"
                    #If no BuiltInProperties were found, the script simply moves on to update CustomProperties
                    if ($FoundBuiltInProperties -eq $null) {
                        Write-Host "No built-in properties for this file" -ForegroundColor Red
                    } else {
                        #If some BuiltInProperties were found, the script will update them
                        Update-PropertiesInFile -CollectionOfProperties $Presentation.BuiltInDocumentProperties -PropertiesToBeUpdated $FoundBuiltInProperties -Binding $Binding
                    }
                }
                #Updates custom properties if required
                if ($script:SetCustomProperties -eq $true) {
                    Write-Host "Updating custom properties..."
                    #Filters $FoundProperties array to get only C (abbreviation for CustomProperties) type properties
                    [array]$FoundCustomInProperties = Filter-PropertiesByType -FoundProperties $FoundProperties -RequiredPropertyType "C"
                    #If no CustomProperties were found, the script simply moves on to update properties in the next document
                    if ($FoundCustomInProperties -eq $null) {
                        Write-Host "No custom properties for this file" -ForegroundColor Red
                    } else {
                        #If some BuiltInProperties were found, the script will update them
                        Update-PropertiesInFile -CollectionOfProperties $Presentation.CustomDocumentProperties -PropertiesToBeUpdated $FoundCustomInProperties -Binding $Binding
                    }
                }
                Close-SavedDocument -DocumentObject $Presentation
            }
        }   
    Write-Host "==========DOCUMENT PROPERTIES UPDATE COMPLETE========="
    }
    $WorkbookWithProperties.Close()
    Start-Sleep -Seconds 3
    Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
}
###SET FILE PROPERTIES FUNCTIONS###

###UI###
Function Custom-Form
{
    #variables
    $script:GetPropertyPathToSelectedFolder = $null
    $script:SetPropertyPathToSelectedFolder = $null
    $script:SetPropertyPathToSelectedFile = $null
    Add-Type -AssemblyName System.Windows.Forms
    #FORM
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon = $false
    $Form.AutoSize = $true
    $Form.Text = "Script Settings"
    $Form.AutoSizeMode = "GrowAndShrink"
    $Form.WindowState = "Normal"
    $Form.SizeGripStyle = "Hide"
    $Form.ShowInTaskbar = $true
    $Form.StartPosition = "CenterScreen"
    $Form.MinimizeBox = $false
    $Form.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #TAB CONTROL
    $TabControl = New-object System.Windows.Forms.TabControl
    $TabControl.Location = New-Object System.Drawing.Size(5,5)
    $TabControl.Size = New-Object System.Drawing.Size(560,545) #width,height
    $Form.Controls.Add($TabControl)
    #GET PROPERTIES PAGE
    $GetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $GetPropertiesPage.Text = "Get Properties”
    $TabControl.Controls.Add($GetPropertiesPage)
    #GET PROPERTIES PAGE ELEMENTS
    #Button 'Browse...'
    $GetPropertyButtonBrowse = New-Object System.Windows.Forms.Button
    $GetPropertyButtonBrowse.Location = New-Object System.Drawing.Point(25,25)
    $GetPropertyButtonBrowse.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonBrowse.Text = "Browse..."
    $GetPropertyButtonBrowse.Add_Click({
        $script:GetPropertyPathToSelectedFolder = Select-Folder -Description "Select folder with files whose properties will be extracted"
        if ($script:GetPropertyPathToSelectedFolder -ne $null) {
            if ($script:GetPropertyPathToSelectedFolder.Length -gt 85) {
                $GetPropertyLabelButtonBrowse.Text = "Specified directory's name is too long to display it here. Hover to see the full path."
                $ToolTip.SetToolTip($GetPropertyLabelButtonBrowse, "$script:GetPropertyPathToSelectedFolder")
            } else {
                $GetPropertyLabelButtonBrowse.Text = "Specified directory: '$(Split-Path -Path $script:GetPropertyPathToSelectedFolder -Leaf)'. Hover to see the full path."
                $ToolTip.SetToolTip($GetPropertyLabelButtonBrowse, "$script:GetPropertyPathToSelectedFolder")
            }
        }
        if ($script:GetPropertyPathToSelectedFolder -ne $null -and ($GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true -or $GetPropertyCheckboxGetCustomProperties.Checked -eq $true)) {
            $GetPropertyButtonExtract.Enabled = $true
        } else {
            $GetPropertyButtonExtract.Enabled = $false
        }
    })
    $GetPropertiesPage.Controls.Add($GetPropertyButtonBrowse)
    #Label for 'Browse...' button
    $GetPropertyLabelButtonBrowse = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonBrowse.Location = New-Object System.Drawing.Point(110,32)
    $GetPropertyLabelButtonBrowse.Width = 400
    $GetPropertyLabelButtonBrowse.Text = "Specify folder with documents whose properties will be extracted"
    $GetPropertyLabelButtonBrowse.AutoSize = $true
    $GetPropertyLabelButtonBrowse.MaximumSize = New-Object System.Drawing.Point(430,38)
    $GetPropertiesPage.Controls.Add($GetPropertyLabelButtonBrowse)
    #Checkbox 'Get Built-In Properties'
    $GetPropertyCheckboxGetBuiltInProperties = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxGetBuiltInProperties.Location = New-Object System.Drawing.Point(25,65)
    $GetPropertyCheckboxGetBuiltInProperties.Width = 300
    $GetPropertyCheckboxGetBuiltInProperties.Text = "Get Built-In Properties"
    $GetPropertyCheckboxGetBuiltInProperties.Add_CheckStateChanged({
        if ($GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetCustomProperties.Checked -eq $false) {
            $GetPropertyButtonExtract.Enabled = $true
        } elseif ($GetPropertyCheckboxGetBuiltInProperties.Checked -eq $false -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetCustomProperties.Checked -eq $true) {
            $GetPropertyButtonExtract.Enabled = $true
        } elseif ($GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetCustomProperties.Checked -eq $true) {
            $GetPropertyButtonExtract.Enabled = $true
        } else {
            $GetPropertyButtonExtract.Enabled = $false
        }
        })
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxGetBuiltInProperties)
    #Checkbox 'Get Custom Properties'
    $GetPropertyCheckboxGetCustomProperties = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxGetCustomProperties.Location = New-Object System.Drawing.Point(25,90)
    $GetPropertyCheckboxGetCustomProperties.Width = 300
    $GetPropertyCheckboxGetCustomProperties.Text = "Get Custom Properties"
    $GetPropertyCheckboxGetCustomProperties.Add_CheckStateChanged({
        if ($GetPropertyCheckboxGetCustomProperties.Checked -eq $true -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetBuiltInProperties.Checked -eq $false) {
            $GetPropertyButtonExtract.Enabled = $true
        } elseif ($GetPropertyCheckboxGetCustomProperties.Checked -eq $false -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true) {
            $GetPropertyButtonExtract.Enabled = $true
        } elseif ($GetPropertyCheckboxGetCustomProperties.Checked -eq $true -and $script:GetPropertyPathToSelectedFolder -ne $null -and $GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true) {
            $GetPropertyButtonExtract.Enabled = $true
        } else {
            $GetPropertyButtonExtract.Enabled = $false
        }
        })
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxGetCustomProperties)
    #Checkbox 'Ignore Properties That Have No Value'
    $GetPropertyCheckboxIgnorePropertiesWithNoValue = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Location = New-Object System.Drawing.Point(25,115)
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Width = 300
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Text = "Ignore Properties That Have No Value"
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Add_CheckStateChanged({})
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxIgnorePropertiesWithNoValue)
    #Checkbox 'Use Blacklist'
    $GetPropertyCheckboxUseBlacklist = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxUseBlacklist.Location = New-Object System.Drawing.Point(25,140)
    $GetPropertyCheckboxUseBlacklist.Width = 300
    $GetPropertyCheckboxUseBlacklist.Text = "Use Blacklist"
    $GetPropertyCheckboxUseBlacklist.Add_CheckStateChanged({
    if ($GetPropertyCheckboxUseBlacklist.Checked -eq $true) {$GetPropertyGroupboxBlacklistSettings.Enabled = $true} else {$GetPropertyGroupboxBlacklistSettings.Enabled = $false}
    })
    $GetPropertiesPage.Controls.Add($GetPropertyCheckboxUseBlacklist)
    #Groupbox 'Blacklist settings'
    $GetPropertyGroupboxBlacklistSettings = New-Object System.Windows.Forms.GroupBox
    $GetPropertyGroupboxBlacklistSettings.Location = New-Object System.Drawing.Point(25,170) #x,y
    $GetPropertyGroupboxBlacklistSettings.Size = New-Object System.Drawing.Point(500,300) #width,height
    $GetPropertyGroupboxBlacklistSettings.Text = "Blacklist Settings"
    $GetPropertyGroupboxBlacklistSettings.Enabled = $false
    $GetPropertiesPage.Controls.Add($GetPropertyGroupboxBlacklistSettings) 
    #Label for 'Black list' listbox
    $GetPropertyLabelBlacklistListBox = New-Object System.Windows.Forms.Label
    $GetPropertyLabelBlacklistListBox.Location =  New-Object System.Drawing.Point(13,19) #x,y
    $GetPropertyLabelBlacklistListBox.Width = 250
    $GetPropertyLabelBlacklistListBox.Height = 13
    $GetPropertyLabelBlacklistListBox.Text = "Properties blacklisted by default:"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelBlacklistListBox)
    #Listbox 'Black list'
    $DefaultBlackList = @("Last author", "Template", "Security", "Revision number", "Application name", "Last print date", "Number of bytes", "Number of characters (with spaces)", "Number of multimedia clips", "Number of hidden Slides", "Number of notes", "Number of slides", "Number of paragraphs", "Number of lines", "Number of characters", "Number of words", "Number of pages", "Total editing time", "Last save time", "Creation date")
    $GetPropertyListBoxBlackList = New-Object System.Windows.Forms.ListBox
    $GetPropertyListBoxBlackList.Location = New-Object System.Drawing.Point(15,35) #x,y
    $GetPropertyListBoxBlackList.Size = New-Object System.Drawing.Point(210,260) #width,height
    $DefaultBlackList | % {$GetPropertyListBoxBlackList.Items.Add($_)} | Out-Null
    $GetPropertyListBoxBlackList.Add_SelectedIndexChanged({
        if ($GetPropertyListBoxBlackList.SelectedIndex -ne -1) {
            #Write-Host "$($GetPropertyListBoxBlackList.SelectedIndex)"
            $GetPropertyInputboxEditItem.Text = $GetPropertyListBoxBlackList.SelectedItem
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyListBoxBlackList)
    #Button 'Add'
    $GetPropertyButtonAddItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonAddItem.Location = New-Object System.Drawing.Point(235,34) #x,y
    $GetPropertyButtonAddItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonAddItem.Text = "Add"
    $GetPropertyButtonAddItem.Add_Click({
        if ($GetPropertyInputboxAddItem.Text -ne "Type in property name to add it...") {
            if ($GetPropertyListBoxBlackList.Items.Contains($GetPropertyInputboxAddItem.Text)) {
                Show-MessageBox -Title "Item already exists" -Type OK -Message "The property you are attempting to add ($($GetPropertyInputboxAddItem.Text)) is already on the list."
            } else {
                $GetPropertyListBoxBlackList.Items.Insert(0, $GetPropertyInputboxAddItem.Text)
                $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
                $GetPropertyInputboxAddItem.ForeColor = "Gray"
            }
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonAddItem)
    #Inputbox 'Add item to the list'
    $GetPropertyInputboxAddItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxAddItem.Location = New-Object System.Drawing.Size(295,35) #x,y
    $GetPropertyInputboxAddItem.Width = 190
    $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
    $GetPropertyInputboxAddItem.ForeColor = "Gray"
    $GetPropertyInputboxAddItem.Add_GotFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "Type in property name to add it...") {
            $GetPropertyInputboxAddItem.Text = ""
            $GetPropertyInputboxAddItem.ForeColor = "Black"
        }
        })
    $GetPropertyInputboxAddItem.Add_LostFocus({
        if ($GetPropertyInputboxAddItem.Text -eq "") {
            $GetPropertyInputboxAddItem.Text = "Type in property name to add it..."
            $GetPropertyInputboxAddItem.ForeColor = "Gray"
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxAddItem)
    #Button 'Edit'
    $GetPropertyButtonEditItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonEditItem.Location = New-Object System.Drawing.Point(235,62) #x,y
    $GetPropertyButtonEditItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonEditItem.Text = "Edit"
    $GetPropertyButtonEditItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -ne "Select item on the blacklist to edit it...") {
            Disable-AllExceptEditing -BooleanRest $false -BooleanEditing $true
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonEditItem)
    #Inputbox 'Edit the selected item'
    $GetPropertyInputboxEditItem = New-Object System.Windows.Forms.TextBox 
    $GetPropertyInputboxEditItem.Location = New-Object System.Drawing.Size(295,63) #x,y
    $GetPropertyInputboxEditItem.Width = 190
    $GetPropertyInputboxEditItem.Enabled = $false
    $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it..."
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyInputboxEditItem)
    #Button 'Apply'
    $GetPropertyButtonApplyItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonApplyItem.Location = New-Object System.Drawing.Point(295,85) #x,y
    $GetPropertyButtonApplyItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonApplyItem.Text = "Apply"
    $GetPropertyButtonApplyItem.Enabled = $false
    $GetPropertyButtonApplyItem.Add_Click({
        if ($GetPropertyInputboxEditItem.Text -eq $GetPropertyListBoxBlackList.SelectedItem) {
            Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
        } else {
            $SelectedIndex = $GetPropertyListBoxBlackList.SelectedIndex
            $GetPropertyListBoxBlackList.Items.Insert($GetPropertyListBoxBlackList.SelectedIndex, ($GetPropertyInputboxEditItem.Text).Trim(' '))
            $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
            $GetPropertyListBoxBlackList.SelectedIndex = $SelectedIndex
            Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonApplyItem)
    #Button 'Cancel'
    $GetPropertyButtonCancelItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonCancelItem.Location = New-Object System.Drawing.Point(347,85) #x,y
    $GetPropertyButtonCancelItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonCancelItem.Text = "Cancel"
    $GetPropertyButtonCancelItem.Enabled = $false
    $GetPropertyButtonCancelItem.Add_Click({
        Disable-AllExceptEditing -BooleanRest $true -BooleanEditing $false    
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonCancelItem)
    #Button 'Delete'
    $GetPropertyButtonDeleteItem = New-Object System.Windows.Forms.Button
    $GetPropertyButtonDeleteItem.Location = New-Object System.Drawing.Point(235,113) #x,y
    $GetPropertyButtonDeleteItem.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonDeleteItem.Text = "Delete"
    $GetPropertyButtonDeleteItem.Add_Click({
        $GetPropertyListBoxBlackList.Items.Remove($GetPropertyListBoxBlackList.SelectedItem)
        $GetPropertyInputboxEditItem.Text = "Select item on the blacklist to edit it..."
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonDeleteItem)
    #Label for 'Delete' button
    $GetPropertyLabelButtonDelete = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonDelete.Location =  New-Object System.Drawing.Point(295,116) #x,y
    $GetPropertyLabelButtonDelete.Size =  New-Object System.Drawing.Point(203,15) #width,height
    $GetPropertyLabelButtonDelete.Text = "Delete unwanted item from the blacklist"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonDelete)
    #Button 'Export'
    $GetPropertyButtonExportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExportList.Location = New-Object System.Drawing.Point(235,141) #x,y
    $GetPropertyButtonExportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonExportList.Text = "Export"
    $GetPropertyButtonExportList.Add_Click({
        $PathToExportedFile = Save-File
        if ($PathToExportedFile -ne $null) {
            If (Test-Path -Path $PathToExportedFile) {Remove-Item -Path $PathToExportedFile -Force}
            $GetPropertyListBoxBlackList.Items | % {Add-Content -Path $PathToExportedFile -Value $_}
        }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonExportList)
    #Label for 'Export' button
    $GetPropertyLabelButtonExport = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonExport.Location =  New-Object System.Drawing.Point(295,144) #x,y
    $GetPropertyLabelButtonExport.Size =  New-Object System.Drawing.Point(200,15) #width,height
    $GetPropertyLabelButtonExport.Text = "Export your blacklist for later use"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonExport)
    #Button 'Import'
    $GetPropertyButtonImportList = New-Object System.Windows.Forms.Button
    $GetPropertyButtonImportList.Location = New-Object System.Drawing.Point(235,169) #x,y
    $GetPropertyButtonImportList.Size = New-Object System.Drawing.Point(50,22) #width,height
    $GetPropertyButtonImportList.Text = "Import"
    $GetPropertyButtonImportList.Add_Click({
        $PathToImportedFile = Open-File -Filter "Text file (*.txt)| *.txt"
        Write-Host $PathToImportedFile
            if ($PathToImportedFile -ne $null) {
                $TxtBlackListContent = @(Get-Content -Path $PathToImportedFile)
                $GetPropertyListBoxBlackList.Items.Clear()
                $TxtBlackListContent | % {$GetPropertyListBoxBlackList.Items.Add($_)}
            }
        })
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyButtonImportList)
    #Label for 'Import' button
    $GetPropertyLabelButtonImport = New-Object System.Windows.Forms.Label
    $GetPropertyLabelButtonImport.Location =  New-Object System.Drawing.Point(295,172) #x,y
    $GetPropertyLabelButtonImport.Size =  New-Object System.Drawing.Point(200,15) #width,height
    $GetPropertyLabelButtonImport.Text = "Import your blacklist"
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyLabelButtonImport)
    #Checkbox 'Get Only Blacklisted Properties'
    $GetPropertyCheckboxTurnIntoWhite = New-Object System.Windows.Forms.CheckBox
    $GetPropertyCheckboxTurnIntoWhite.Location = New-Object System.Drawing.Point(235,200) #x,y
    $GetPropertyCheckboxTurnIntoWhite.Width = 200
    $GetPropertyCheckboxTurnIntoWhite.Text = "Get Only Blacklisted Properties"
    $GetPropertyCheckboxTurnIntoWhite.Add_CheckStateChanged({})
    $GetPropertyGroupboxBlacklistSettings.Controls.Add($GetPropertyCheckboxTurnIntoWhite)
    #Button 'Extract Properties'
    $GetPropertyButtonExtract = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExtract.Location = New-Object System.Drawing.Point(25,480) #x,y
    $GetPropertyButtonExtract.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonExtract.Text = "Extract"
    $GetPropertyButtonExtract.Enabled = $false
    $GetPropertyButtonExtract.Add_Click({
        $ClickResult = Show-MessageBox -Title "Warning" -Type OKCancel -Message "The script will close the following applications: Excel, Word, PowerPoint, Visio.$([System.Environment]::NewLine)To prevent data loss, make sure you have saved and closed any documents opened in the listed applications before clicking 'OK'."
        if ($ClickResult -eq "OK") {
            Write-Host "Script started"
            Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
            if ($GetPropertyCheckboxGetBuiltInProperties.Checked -eq $true) {$script:GetBuiltInProperties = $true} else {$script:GetBuiltInProperties = $false}
            if ($GetPropertyCheckboxGetCustomProperties.Checked -eq $true) {$script:GetCustomProperties = $true} else {$script:GetCustomProperties = $false}
            if ($GetPropertyCheckboxIgnorePropertiesWithNoValue.Checked -eq $true) {$script:IgnorePropertiesWithNoValue = $true} else {$script:IgnorePropertiesWithNoValue = $false}
            if ($GetPropertyCheckboxUseBlacklist.Checked -eq $true) {
                $script:UseBlacklist = $true
                $script:GetPropertiesBlacklist = @()
                $GetPropertyListBoxBlackList.Items | % {$script:GetPropertiesBlacklist += $_}
            } else {
                $script:UseBlacklist = $false
            }
            if ($GetPropertyCheckboxTurnIntoWhite.Checked -eq $true) {$script:WhitelistEnabled = $true} else {$script:WhitelistEnabled = $false}
            Get-FileProperties
        } else {
            Write-Host "Script not started"
        }
    })
    $GetPropertiesPage.Controls.Add($GetPropertyButtonExtract)
    #Button 'Exit'
    $GetPropertyButtonExit = New-Object System.Windows.Forms.Button
    $GetPropertyButtonExit.Location = New-Object System.Drawing.Point(115,480) #x,y
    $GetPropertyButtonExit.Size = New-Object System.Drawing.Point(80,30)
    $GetPropertyButtonExit.Text = "Exit"
    $GetPropertyButtonExit.Add_Click({
    $Form.Close();
    })
    $GetPropertiesPage.Controls.Add($GetPropertyButtonExit)
    #SET PROPERTIES PAGE
    $SetPropertiesPage = New-Object System.Windows.Forms.TabPage
    $SetPropertiesPage.Text = "Set Properties”
    $TabControl.Controls.Add($SetPropertiesPage)
    #SET PROPERTIES PAGE ELEMENTS
    
    #Button 'Browse...' (for file)
    $SetPropertyButtonBrowseFile = New-Object System.Windows.Forms.Button
    $SetPropertyButtonBrowseFile.Location = New-Object System.Drawing.Point(25,25) #x,y
    $SetPropertyButtonBrowseFile.Size = New-Object System.Drawing.Point(80,30)
    $SetPropertyButtonBrowseFile.Text = "Browse..."
    $SetPropertyButtonBrowseFile.Add_Click({
        $script:SetPropertyPathToSelectedFile = Open-File -Filter "Excel file (*.xlsx)| *.xlsx"
        Write-Host $script:SetPropertyPathToSelectedFile
        if ($script:SetPropertyPathToSelectedFile -ne $null) {
            if ($script:SetPropertyPathToSelectedFile.Length -gt 85) {
                $SetPropertyLabelButtonBrowseFile.Text = "Specified files's name is too long to display it here. Hover to see the full path."
                $ToolTip.SetToolTip($SetPropertyLabelButtonBrowseFile, "$script:SetPropertyPathToSelectedFile")
            } else {
                $SetPropertyLabelButtonBrowseFile.Text = "Specified file: '$(Split-Path -Path $script:SetPropertyPathToSelectedFile -Leaf)'. Hover to see the full path."
                $ToolTip.SetToolTip($SetPropertyLabelButtonBrowseFile, "$script:SetPropertyPathToSelectedFile")
            }
        }
        if ($script:SetPropertyPathToSelectedFolder -ne $null -and ($SetPropertyCheckboxSetBProperties.Checked -eq $true -or $SetPropertyCheckboxSetCProperties.Checked -eq $true)) {
            $SetPropertyButtonSet.Enabled = $true
        }
    })
    $SetPropertiesPage.Controls.Add($SetPropertyButtonBrowseFile)
    #Label for 'Browse...' (for file) button 
    $SetPropertyLabelButtonBrowseFile = New-Object System.Windows.Forms.Label
    $SetPropertyLabelButtonBrowseFile.Location =  New-Object System.Drawing.Point(110,32) #x,y
    $SetPropertyLabelButtonBrowseFile.Width = 400
    $SetPropertyLabelButtonBrowseFile.Text = "Specify *.xlsx that contains properties you want to insert"
    $SetPropertyLabelButtonBrowseFile.AutoSize = $true
    $SetPropertyLabelButtonBrowseFile.MaximumSize = New-Object System.Drawing.Point(430,38)
    $SetPropertiesPage.Controls.Add($SetPropertyLabelButtonBrowseFile)
    #Checkbox 'Do Not Clear Filtering'
    $SetPropertyCheckboxDoNotClearFiltering = New-Object System.Windows.Forms.CheckBox
    $SetPropertyCheckboxDoNotClearFiltering.Location = New-Object System.Drawing.Point(25,65) #x,y
    $SetPropertyCheckboxDoNotClearFiltering.Width = 500
    $SetPropertyCheckboxDoNotClearFiltering.Text = "Do Not Clear Applied Filtering"
    $SetPropertyCheckboxDoNotClearFiltering.Add_CheckStateChanged({})
    $SetPropertiesPage.Controls.Add($SetPropertyCheckboxDoNotClearFiltering)
    #Button 'Browse...' (for folder)
    $SetPropertyButtonBrowseFolder = New-Object System.Windows.Forms.Button
    $SetPropertyButtonBrowseFolder.Location = New-Object System.Drawing.Point(25,97) #x,y
    $SetPropertyButtonBrowseFolder.Size = New-Object System.Drawing.Point(80,30)
    $SetPropertyButtonBrowseFolder.Text = "Browse..."
    $SetPropertyButtonBrowseFolder.Add_Click({
    $script:SetPropertyPathToSelectedFolder = Select-Folder -Description "Select folder with files whose properties will be extracted"
        Write-Host $script:SetPropertyPathToSelectedFolder
        if ($script:SetPropertyPathToSelectedFolder -ne $null) {
            if ($script:SetPropertyPathToSelectedFolder.Length -gt 85) {
                $SetPropertyLabelButtonBrowseFolder.Text = "Specified directory's name is too long to display it here. Hover to see the full path."
                $ToolTip.SetToolTip($SetPropertyLabelButtonBrowseFolder, "$script:SetPropertyPathToSelectedFolder")
            } else {
                $SetPropertyLabelButtonBrowseFolder.Text = "Specified directory: '$(Split-Path -Path $script:SetPropertyPathToSelectedFolder -Leaf)'. Hover to see the full path."
                $ToolTip.SetToolTip($SetPropertyLabelButtonBrowseFolder, "$script:SetPropertyPathToSelectedFolder")
            }
        }
        if ($script:SetPropertyPathToSelectedFile -ne $null -and ($SetPropertyCheckboxSetBProperties.Checked -eq $true -or $SetPropertyCheckboxSetCProperties.Checked -eq $true)) {
            $SetPropertyButtonSet.Enabled = $true
        }
    })
    $SetPropertiesPage.Controls.Add($SetPropertyButtonBrowseFolder)
    #Label for 'Browse...' (for folder) button 
    $SetPropertyLabelButtonBrowseFolder = New-Object System.Windows.Forms.Label
    $SetPropertyLabelButtonBrowseFolder.Location =  New-Object System.Drawing.Point(110,105) #x,y
    $SetPropertyLabelButtonBrowseFolder.Width = 400
    $SetPropertyLabelButtonBrowseFolder.Text = "Specify folder with documents whose properties will be updated"
    $SetPropertyLabelButtonBrowseFolder.AutoSize = $true
    $SetPropertyLabelButtonBrowseFolder.MaximumSize = New-Object System.Drawing.Point(430,38)
    $SetPropertiesPage.Controls.Add($SetPropertyLabelButtonBrowseFolder)
    #Checkbox 'Set Built-In Properties'
    $SetPropertyCheckboxSetBProperties = New-Object System.Windows.Forms.CheckBox
    $SetPropertyCheckboxSetBProperties.Location = New-Object System.Drawing.Point(25,137) #x,y
    $SetPropertyCheckboxSetBProperties.Width = 500
    $SetPropertyCheckboxSetBProperties.Text = "Set Built-In Properties"
    $SetPropertyCheckboxSetBProperties.Checked = $true
    $SetPropertyCheckboxSetBProperties.Add_CheckStateChanged({
        if ($SetPropertyCheckboxSetBProperties.Checked -eq $true -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $false) {
            $SetPropertyButtonSet.Enabled = $true
        } elseif ($SetPropertyCheckboxSetBProperties.Checked -eq $false -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $true) {
            $SetPropertyButtonSet.Enabled = $true
        } elseif ($SetPropertyCheckboxSetBProperties.Checked -eq $true -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $true) {
            $SetPropertyButtonSet.Enabled = $true
        } else {
            $SetPropertyButtonSet.Enabled = $false
        }
    })
    $SetPropertiesPage.Controls.Add($SetPropertyCheckboxSetBProperties)
    #Checkbox 'Set Custom Properties'
    $SetPropertyCheckboxSetCProperties = New-Object System.Windows.Forms.CheckBox
    $SetPropertyCheckboxSetCProperties.Location = New-Object System.Drawing.Point(25,162) #x,y
    $SetPropertyCheckboxSetCProperties.Width = 500
    $SetPropertyCheckboxSetCProperties.Text = "Set Custom Properties"
    $SetPropertyCheckboxSetCProperties.Checked = $true
    $SetPropertiesPage.Controls.Add($SetPropertyCheckboxSetCProperties)
    $SetPropertyCheckboxSetCProperties.Add_CheckStateChanged({
        if ($SetPropertyCheckboxSetBProperties.Checked -eq $false -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $true) {
            $SetPropertyButtonSet.Enabled = $true
        } elseif ($SetPropertyCheckboxSetBProperties.Checked -eq $true -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $false) {
            $SetPropertyButtonSet.Enabled = $true
        } elseif ($SetPropertyCheckboxSetBProperties.Checked -eq $true -and $script:SetPropertyPathToSelectedFile -ne $null -and $script:SetPropertyPathToSelectedFolder -ne $null -and $SetPropertyCheckboxSetCProperties.Checked -eq $true) {
            $SetPropertyButtonSet.Enabled = $true
        } else {
            $SetPropertyButtonSet.Enabled = $false
        }
    })
    #Button 'Set'
    $SetPropertyButtonSet = New-Object System.Windows.Forms.Button
    $SetPropertyButtonSet.Location = New-Object System.Drawing.Point(25,480) #x,y
    $SetPropertyButtonSet.Size = New-Object System.Drawing.Point(80,30)
    $SetPropertyButtonSet.Text = "Update"
    $SetPropertyButtonSet.Enabled = $false
    $SetPropertyButtonSet.Add_Click({
    $ClickResult = Show-MessageBox -Title "Warning" -Type OKCancel -Message "The script will close the following applications: Excel, Word, PowerPoint, Visio.$([System.Environment]::NewLine)To prevent data loss, make sure you have saved and closed any documents opened in the listed applications before clicking 'OK'."
    if ($ClickResult -eq "OK") {
        Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
        if ($SetPropertyCheckboxDoNotClearFiltering.Checked -eq $true) {$script:DoNotClearAppliedFiltering = $true} else {$script:DoNotClearAppliedFiltering = $false}
        if ($SetPropertyCheckboxSetBProperties.Checked -eq $true) {$script:SetBuiltInProperties = $true} else {$script:SetBuiltInProperties = $false}
        if ($SetPropertyCheckboxSetCProperties.Checked -eq $true) {$script:SetCustomProperties = $true} else {$script:SetCustomProperties = $false}
        Set-FileProperties
        Write-Host "Script started"
    } else {
        Write-Host "Script not started"
    }
    })
    $SetPropertiesPage.Controls.Add($SetPropertyButtonSet)
    #Button 'Exit'
    $SetPropertyButtonExit = New-Object System.Windows.Forms.Button
    $SetPropertyButtonExit.Location = New-Object System.Drawing.Point(115,480) #x,y
    $SetPropertyButtonExit.Size = New-Object System.Drawing.Point(80,30)
    $SetPropertyButtonExit.Text = "Exit"
    $SetPropertyButtonExit.Add_Click({
        $Form.Close();
    })
    $SetPropertiesPage.Controls.Add($SetPropertyButtonExit)
    $Form.ShowDialog()
}
Function Disable-AllExceptEditing ($BooleanRest, $BooleanEditing) 
{
    $GetPropertyButtonBrowse.Enabled = $BooleanRest
    $GetPropertyLabelButtonBrowse.Enabled = $BooleanRest
    $GetPropertyCheckboxGetBuiltInProperties.Enabled = $BooleanRest
    $GetPropertyCheckboxGetCustomProperties.Enabled = $BooleanRest
    $GetPropertyCheckboxIgnorePropertiesWithNoValue.Enabled = $BooleanRest
    $GetPropertyCheckboxUseBlacklist.Enabled = $BooleanRest
    $GetPropertyListBoxBlackList.Enabled = $BooleanRest
    $GetPropertyButtonAddItem.Enabled = $BooleanRest
    $GetPropertyInputboxAddItem.Enabled = $BooleanRest
    $GetPropertyButtonDeleteItem.Enabled = $BooleanRest
    $GetPropertyLabelButtonDelete.Enabled = $BooleanRest
    $GetPropertyButtonExportList.Enabled = $BooleanRest
    $GetPropertyLabelButtonExport.Enabled = $BooleanRest
    $GetPropertyButtonImportList.Enabled = $BooleanRest
    $GetPropertyLabelButtonImport.Enabled = $BooleanRest
    $GetPropertyInputboxEditItem.Enabled = $BooleanEditing
    $GetPropertyButtonApplyItem.Enabled = $BooleanEditing
    $GetPropertyButtonCancelItem.Enabled = $BooleanEditing
    $GetPropertyButtonEditItem.Enabled = $BooleanRest
    $GetPropertyCheckboxTurnIntoWhite.Enabled = $BooleanRest
    $GetPropertyButtonExtract.Enabled = $BooleanRest
    $GetPropertyButtonExit.Enabled = $BooleanRest
}
Function Save-File
{ 
    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "Text file (*.txt)| *.txt"
    $DialogResult = $SaveFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SaveFileDialog.FileName} else {return $null}
}
Function Open-File ($Filter)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = $Filter
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $OpenFileDialog.FileName} else {return $null}
}
Function Select-Folder ($Description)
{
    Add-Type -AssemblyName System.Windows.Forms
    $SelectFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $SelectFolderDialog.Rootfolder = "Desktop"
    $SelectFolderDialog.Description = $Description
    $DialogResult = $SelectFolderDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SelectFolderDialog.SelectedPath} else {return $null}
}
Function Show-MessageBox 
{ 
    param($Message, $Title, [ValidateSet("OK", "OKCancel")]$Type)
    Add-Type –AssemblyName System.Windows.Forms 
    if ($Type -eq "OK") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title")}  
    if ($Type -eq "OKCancel") {[System.Windows.Forms.MessageBox]::Show("$Message","$Title",[System.Windows.Forms.MessageBoxButtons]::OKCancel)}
}
Custom-Form
###UI###
