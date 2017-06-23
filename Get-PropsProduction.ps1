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

Get-ChildItem -Path $SelectedPath | % {
    Start-Sleep -Seconds 2
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
        $Workbook.Close()
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
        Write-Host $CollectedPropertiesData[0][0]
        Write-Host $CollectedPropertiesData[1][0]
        $Document.Close([ref]0)
        $Word.Quit()
    }
    if ($_.Extension -eq ".vsd") {
        $Visio = New-Object -ComObject Visio.Application
        $Visio.Visible = $false
        $Document = $Visio.Documents.Open("C:\Users\Tsedik\Desktop\Новая папка\Документ Microsoft Visio.vsd")
        $Document.Subject
        $Document.Close()
        $Visio.Quit()
    }
    if ($_.Extension -eq ".pptx"){
        $PowerPoint = New-Object -ComObject PowerPoint.Application
        #try {$PowerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse} catch {"kek"}
        $Presentation = $PowerPoint.Presentations.Open("C:\Users\Tsedik\Desktop\Новая папка\Презентация Microsoft PowerPoint.pptx")
        $FileProperties = $Presentation.BuiltInDocumentProperties
        $CollectedPropertiesData = Get-FileProperties -BindingFlags $Binding -CollectionOfProperties $FileProperties
        Write-Host $CollectedPropertiesData[0][0]
        Write-Host $CollectedPropertiesData[1][0]
        $Presentation.Close()
        $PowerPoint.Quit()
    }
}

