clear
$script:SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка"
$script:SelectedFile = "C:\Users\Tsedik\Desktop\Properties.xlsx"

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
        Write-Host "No properties for $FileName"
    } else {
        $FirstMatch = $Target
        Do
        {
        Write-Host "Property for $FileName found"
        $Target = $RangeToSearchThrough.FindNext($Target)
        }
        While ($Target.AddressLocal() -ne $FirstMatch.AddressLocal())
    }
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
    Get-ChildItem -Path $script:SelectedPath | % {
    #Finds all properties for the file being processed and puts it into array
    Find-PropertiesForFile -LastNonEmptyCell $LastNonemptyCellInColumn -WorksheetWithProperties $WorksheetWithProperties -FileName $_.Name
    }
    Start-Sleep -Seconds 3
    Kill -Name VISIO, POWERPNT, EXCEL, WINWORD -ErrorAction SilentlyContinue
}
Set-FileProperties
