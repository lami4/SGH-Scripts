clear
Function Get-FileProperties ($BindingFlags, $CollectionOfProperties) {
    foreach ($Property in $CollectionOfProperties) {
        [array]$PropertyNames += [System.__ComObject].InvokeMember(“name”,$BindingFlags::GetProperty,$null,$Property,$null)
        trap [system.exception] {continue}
        [array]$PropertyValues += [System.__ComObject].InvokeMember(“value”,$BindingFlags::GetProperty,$null,$Property,$null)
    }
    return $PropertyNames, $PropertyValues
}

#Stores the BindingFlags enumeration in $Binding
$Binding = “System.Reflection.BindingFlags” -as [type]
$SelectedPath = "C:\Users\Tsedik\Desktop\Новая папка"
$ExcelOutput = New-Object -ComObject Excel.Application
$ExcelOutput.Visible = $true
$WorkbookOutput = $ExcelOutput.Workbooks.Add()
$WorksheetOutput = $WorkbookOutput.Worksheets.Item(1)
Get-ChildItem -Path $SelectedPath | % {
    if ($_.Extension -eq ".xlsx") {
        #Starts MS Excel
        $Excel = New-Object -ComObject Excel.Application
        #Makes it not visible to the user
        $Excel.Visible = $false
        #Opens the file whose properties are to be extracted
        $Workbook = $Excel.Workbooks.Open("C:\Users\Tsedik\Desktop\Новая папка\Лист Microsoft Excel.xlsx")
        #Gets a collection of properties and puts it in $FileProperties
        $FileProperties = $Workbook.BuiltInDocumentProperties
        #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
        $CollectedPropertiesData = Get-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
        Write-Host $CollectedPropertiesData[0][0]
        Write-Host $CollectedPropertiesData[1][0]
        #Closes active workbook without saving and quits MS Word
        $Workbook.Close($false)
        $Excel.Quit()
    }
    if ($_.Extension -eq ".docx") {
        #Starts MS Word
        $Word = New-Object -ComObject Word.Application
        #Makes it not visible to the user
        $Word.Visible = $false
        #Opens the file whose properties are to be extracted
        $Document = $Word.Documents.Open("C:\Users\Tsedik\Desktop\Новая папка\Документ Microsoft Word.docx")
        #Gets a collection of properties and puts it in $FileProperties
        $FileProperties = $Document.BuiltInDocumentProperties
        #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
        $CollectedPropertiesData = Get-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
        Write-Host $CollectedPropertiesData[0]
        Write-Host $CollectedPropertiesData[1][0]
        #Closes active document without saving and quits MS Word
        $Document.Close([ref]0)
        $Word.Quit()
    }
    if ($_.Extension -eq ".vdw") {
        #Starts MS Visio and makes it invisible
        $Visio = New-Object -ComObject Visio.InvisibleApp
        #Opens a document
        $Document = $Visio.Documents.Open("C:\Users\Tsedik\Desktop\Новая папка\Документ Microsoft Visio.vsd")
        #List of Built In Document Properties
        $VisioDocumentBuiltInProperties = @("Subject", "Title", "Creator", "Manager", "Company", "Language", "Category", "Keywords", "Description", "HyperlinkBase", "TimeCreated", "TimeEdited", "TimePrinted", "TimeSaved", "Stat", "Version")
        $VisioDocumentBuiltInProperties | % {$Document.$_}
        #Closes active document without saving and quits MS Visio
        $Document.Close()
        $Visio.Quit()
    }
    if ($_.Extension -eq ".pptx") {
        #Add the Office assembly to the current Windows PowerShell session
        Add-type -AssemblyName Office
        #Starts MS PowerPoint
        $PowerPoint = New-Object -ComObject PowerPoint.Application
        #Opens a presentation and makes it not visible to the user
        $Presentation = $PowerPoint.Presentations.Open("C:\Users\Tsedik\Desktop\Новая папка\Презентация Microsoft PowerPoint.pptx", $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
        #Gets a collection of properties and puts it in $FileProperties
        $FileProperties = $Presentation.BuiltInDocumentProperties
        #Uses Get-FileProperties function to extract file properties to $CollectedPropertiesData array
        $CollectedPropertiesData = Get-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
        Write-Host $CollectedPropertiesData[0][0]
        Write-Host $CollectedPropertiesData[1][0]
        #Closes active presentation without saving, quits PowerPoint, nulls out all variables related to PowerPoint and calls garbage collection
        $Presentation.Close()
        $PowerPoint.Quit()
        $PowerPoint = $null
        $Presentation = $null
        $FileProperties = $null
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
    }
}
