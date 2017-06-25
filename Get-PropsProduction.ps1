clear
Function Run-MSApplication ($AppName, $AppExtensions, $SelectedPath, $Text) {
   #Gets each file's extension in the folder specified by user
   $ExtensionsInSelectedFolder = @(Get-ChildItem -Path $SelectedPath | % {$_.Extension})
   #If a file's extension from the folder specified by user matches an extension in the $AppExtensions array, opens the required application
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
$SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка"

Function Output-CollectedPropertiesToExcelTable () {

}

Function Get-FileProperties ($SelectedPath) {
    #Adds the Office assembly to the current Windows PowerShell session
    Add-type -AssemblyName Office
    #Stores the BindingFlags enumeration in $Binding
    $Binding = “System.Reflection.BindingFlags” -as [type]
    #Extensions that will be processed by the script
    $WordExtensions = @(".doc", ".docx", ".dotm")
    $ExcelExtensions = @(".xlsx", ".xls", ".xltm")
    $VisioExtensions = @(".vdx", ".vsd", ".vdw")
    $PowerPointExtensions = @(".pptx", ".ppt")
    #Starts MS Word if necessary
    $Word = Run-MSApplication -AppName "Word.Application" -AppExtensions $WordExtensions -SelectedPath $SelectedPath -Text "Word" 
    #If MS Word is started, makes it not visible to user
    if ($Word -ne $null) {$Word.Visible = $false}
    #Starts MS Excel if necessary
    $Excel = Run-MSApplication -AppName "Excel.Application" -AppExtensions $ExcelExtensions -SelectedPath $SelectedPath -Text "Excel"
    #If MS Excel is started, makes it not visible to user
    if ($Excel -ne $null) {$Excel.Visible = $false}
    #Starts MS Visio if necessary and makes it not visible to user
    $Visio = Run-MSApplication -AppName "Visio.InvisibleApp" -AppExtensions $VisioExtensions -SelectedPath $SelectedPath -Text "Visio"
    #Starts MS PowerPoint if necessary. Invisibility is enabled when opening a presentation.
    $PowerPoint = Run-MSApplication -AppName "PowerPoint.Application" -AppExtensions $PowerPointExtensions -SelectedPath $SelectedPath -Text "PowerPoint"
    <#
    #Opens excel worksheet to output data to
    $ExcelOutput = New-Object -ComObject Excel.Application
    $ExcelOutput.Visible = $true
    $WorkbookOutput = $ExcelOutput.Workbooks.Add()
    $WorksheetOutput = $WorkbookOutput.Worksheets.Item(1)
    #>
    #Starts looping through each file in the folder specified by user
    Get-ChildItem -Path $SelectedPath | % {
        #if extension of the processed file matches an extension in $ExcelExtensions, the script will open it and extract its properties using Excel application
        if ($ExcelExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted
            $Workbook = $Excel.Workbooks.Open("$SelectedPath\Лист Microsoft Excel.xlsx")
            #Gets a collection of properties and puts it in $FileProperties
            $FileProperties = $Workbook.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            Write-Host $CollectedPropertiesData[0][0]
            Write-Host $CollectedPropertiesData[1][0]
            #Closes active workbook without saving it
            $Workbook.Close()
        }
        #if extension of the processed file matches an extension in $WordExtensionss, the script will open it and extract its properties using Word application
        if ($WordExtensions -contains ($_.Extension).ToLower()) {
            #Opens the file whose properties are to be extracted
            $Document = $Word.Documents.Open("$SelectedPath\Документ Microsoft Word.docx")
            #Gets a collection of properties and puts it in $FileProperties
            $FileProperties = $Document.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            Write-Host $CollectedPropertiesData[0]
            Write-Host $CollectedPropertiesData[1][0]
            #Closes active document without saving it
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions, the script will open it and extract its properties using Visio application
        if ($VisioExtensions -contains ($_.Extension).ToLower()) {
            #Opens a document
            $Document = $Visio.Documents.Open("$SelectedPath\Документ Microsoft Visio.vsd")
            #List of Built In Document Properties
            $VisioDocumentBuiltInProperties = @("Subject", "Title", "Creator", "Manager", "Company", "Language", "Category", "Keywords", "Description", "HyperlinkBase", "TimeCreated", "TimeEdited", "TimePrinted", "TimeSaved", "Stat", "Version")
            $VisioDocumentBuiltInProperties | % {$Document.$_}
            #Closes active document without saving
            $Document.Close()
        }
        #if extension of the processed file matches an extension in $VisioExtensions, the script will open it and extract its properties using Visio application
        if ($PowerPointExtensions -contains ($_.Extension).ToLower()) {
            #Opens a presentation and makes it not visible to the user
            $Presentation = $PowerPoint.Presentations.Open("$SelectedPath\Презентация Microsoft PowerPoint.pptx", $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            #Gets a collection of properties and puts it in $FileProperties
            $FileProperties = $Presentation.BuiltInDocumentProperties
            #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
            $CollectedPropertiesData = Extract-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
            Write-Host $CollectedPropertiesData[0][0]
            Write-Host $CollectedPropertiesData[1][0]
            #Closes active presentation without saving it
            $Presentation.Close()
        }
    }
    Write-Host "Closing opened MS Office applications... It may take up to 30 seconds... Ignore any warning messages if appear..."
    Start-Sleep -Seconds 3
    if ($Excel -ne $null) {$Excel.Quit(); $Excel = $null; $Workbook = $null; $FileProperties = $null; [gc]::collect(); [gc]::WaitForPendingFinalizers()}
    if ($Word -ne $null) {$Word.Quit()}
    if ($Visio -ne $null) {$Visio.Quit()}
    if ($PowerPoint -ne $null) {$PowerPoint.Quit(); $PowerPoint = $null; $Presentation = $null; $FileProperties = $null; [gc]::collect(); [gc]::WaitForPendingFinalizers()}
}
Get-FileProperties -SelectedPath $SelectedPath
