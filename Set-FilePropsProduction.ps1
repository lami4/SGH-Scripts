clear
Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
$script:SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка\Test2"
$script:SelectedFile = "C:\Users\Tsedik\Desktop\Properties.xlsx"
$script:SetBuiltInProperties = $true
$script:SetCustomProperties = $true

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
    $Word = Run-MSApplication -AppName "Word.Application" -AppExtensions $WordExtensions -Text "Word" -PathToFolder $script:SelectedPath
    #If MS Word is started, makes it not visible to user.
    if ($Word -ne $null) {$Word.Visible = $false}
    #Starts MS Excel if necessary
    $Excel = Run-MSApplication -AppName "Excel.Application" -AppExtensions $ExcelExtensions -Text "Excel" -PathToFolder $script:SelectedPath
    #If MS Excel is started, makes it not visible to user.
    if ($Excel -ne $null) {$Excel.Visible = $false}
    #Starts MS Visio if necessary and makes it not visible to user.
    $Visio = Run-MSApplication -AppName "Visio.InvisibleApp" -AppExtensions $VisioExtensions -Text "Visio" -PathToFolder $script:SelectedPath
    #Starts MS PowerPoint if necessary. Invisibility is enabled when opening a presentation.
    $PowerPoint = Run-MSApplication -AppName "PowerPoint.Application" -AppExtensions $PowerPointExtensions -Text "PowerPoint" -PathToFolder $script:SelectedPath
    #Creates another instance of Excel application, opens Excel file that contains the properties and activates the first sheet.
    $ExcelWithProperties = New-Object -ComObject Excel.Application
    $ExcelWithProperties.Visible = $false
    $WorkbookWithProperties = $ExcelWithProperties.WorkBooks.Open($script:SelectedFile)
    $WorksheetWithProperties = $WorkbookWithProperties.Worksheets.Item(1)
    #Finds last non empty cell in column C. This value will be used later to create a range in Excel.
    $LastNonemptyCellInColumn = $WorksheetWithProperties.Range("C:C").End(-4121).Row
    Write-Host "Getting ready..."
    Start-Sleep -Seconds 2
    Get-ChildItem -Path $script:SelectedPath | % {
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
    Write-Host "=============================================================="
    }
    $WorkbookWithProperties.Close()
    Start-Sleep -Seconds 3
    Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
}
Set-FileProperties
