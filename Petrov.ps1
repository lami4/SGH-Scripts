#Сделал этот скрипт на базе вот этой статьи
#Я не понимаю для параметра $HtmlOutput нужно указывать полный путь к файлу, в оторый выгружается результаты преобразовани, или только папку? 
#Можешь в переменоой указать папку прямо вот тут сверху.
#https://paregov.net/blog/19-powershell/24-xslt-processor-with-powershell


$script:PathToXmlFile = $null
$script:PathToXsltFile = $null
$script:PathToOutput = $null

#Функция преобразования
Function Transform-XmlFile 
(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$XmlPath,
     
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$XslPath,
     
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$HtmlOutput
)
{
    try
    {
        $XslPatht = New-Object System.Xml.Xsl.XslCompiledTransform
        $XslPatht.Load($XslPath)
        $XslPatht.Transform($XmlPath, $HtmlOutput)
     
        Write-Host "Generated output is on path: $HtmlOutput"
    }
    catch
    {
        Write-Host $_.Exception -ForegroundColor Red
    }
}

Function Save-File
{ 
    Add-Type -AssemblyName System.Windows.Forms
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "XML file (*.xml)| *.xml"
    $DialogResult = $SaveFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $SaveFileDialog.FileName} else {return $null}
}

Function Open-File ($Filter, $MultipleSelectionFlag)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = $Filter
    if ($MultipleSelectionFlag -eq $true) {$OpenFileDialog.Multiselect = $true}
    if ($MultipleSelectionFlag -eq $false) {$OpenFileDialog.Multiselect = $false}
    $DialogResult = $OpenFileDialog.ShowDialog()
    if ($DialogResult -eq "OK") {return $OpenFileDialog.FileNames} else {return $null}
}


Function ApplyChangesForm ()
{
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $ApplyChangesForm = New-Object System.Windows.Forms.Form
    $ApplyChangesForm.Padding = New-Object System.Windows.Forms.Padding(0,0,10,10)
    $ApplyChangesForm.ShowIcon = $false
    $ApplyChangesForm.AutoSize = $true
    $ApplyChangesForm.Text = "Внести изменения"
    $ApplyChangesForm.AutoSizeMode = "GrowAndShrink"
    $ApplyChangesForm.WindowState = "Normal"
    $ApplyChangesForm.SizeGripStyle = "Hide"
    $ApplyChangesForm.ShowInTaskbar = $true
    $ApplyChangesForm.StartPosition = "CenterScreen"
    $ApplyChangesForm.MinimizeBox = $false
    $ApplyChangesForm.MaximizeBox = $false
    #TOOLTIP
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    #Кнопка обзор
    $ApplyChangesFormFilesBeingPublished = New-Object System.Windows.Forms.Button
    $ApplyChangesFormFilesBeingPublished.Location = New-Object System.Drawing.Point(10,10) #x,y
    $ApplyChangesFormFilesBeingPublished.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormFilesBeingPublished.Text = "Обзор..."
    $ApplyChangesFormFilesBeingPublished.TabStop = $false
    $ApplyChangesFormFilesBeingPublished.Add_Click({
        $script:PathToXmlFile = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:PathToXmlFile -ne $null) {
            $ApplyChangesFormFilesBeingPublishedLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToXmlFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesFormFilesBeingPublishedLabel, $script:PathToXmlFile)
            #Write-Host $script:PathToXmlFile
        }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormFilesBeingPublished)
    #Поле к кнопке Обзор
    $ApplyChangesFormFilesBeingPublishedLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesFormFilesBeingPublishedLabel.Location =  New-Object System.Drawing.Point(95,14) #x,y
    $ApplyChangesFormFilesBeingPublishedLabel.Width = 500
    $ApplyChangesFormFilesBeingPublishedLabel.Text = "Укажите XML-файл"
    $ApplyChangesFormFilesBeingPublishedLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesFormFilesBeingPublishedLabel)
    #Кнопка обзор
    $ApplyChangesFormCurrentVersion = New-Object System.Windows.Forms.Button
    $ApplyChangesFormCurrentVersion.Location = New-Object System.Drawing.Point(10,42) #x,y
    $ApplyChangesFormCurrentVersion.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormCurrentVersion.Text = "Обзор..."
    $ApplyChangesFormCurrentVersion.TabStop = $false
    $ApplyChangesFormCurrentVersion.Add_Click({
        $script:PathToXsltFile = Open-File -Filter "All files (*.*)| *.*" -MultipleSelectionFlag $false
        if ($script:PathToXsltFile -ne $null) {
            $ApplyChangesFormCurrentVersionLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToXsltFile -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesFormCurrentVersionLabel, $script:PathToXsltFile)
            #Write-Host $script:PathToXsltFile
        } 
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCurrentVersion)
    #Поле к кнопке Обзор
    $ApplyChangesFormCurrentVersionLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesFormCurrentVersionLabel.Location =  New-Object System.Drawing.Point(95,46) #x,y
    $ApplyChangesFormCurrentVersionLabel.Width = 500
    $ApplyChangesFormCurrentVersionLabel.Text = "Укажите XSLT-файл"
    $ApplyChangesFormCurrentVersionLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCurrentVersionLabel)
    #Кнопка обзор
    $ApplyChangesArchiveFolder = New-Object System.Windows.Forms.Button
    $ApplyChangesArchiveFolder.Location = New-Object System.Drawing.Point(10,74) #x,y
    $ApplyChangesArchiveFolder.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesArchiveFolder.Text = "Обзор..."
    $ApplyChangesArchiveFolder.TabStop = $false
    $ApplyChangesArchiveFolder.Add_Click({
        $script:PathToOutput = Save-File
        if ($script:PathToOutput -ne $null) {
            $ApplyChangesArchiveFolderLabel.Text = "Указанный файл: $(Split-Path -Path $script:PathToOutput -Leaf). Наведите курсором, чтобы увидеть полный путь."
            $ToolTip.SetToolTip($ApplyChangesArchiveFolderLabel, $script:PathToOutput)
            #Write-Host $script:PathToOutput
        }
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesArchiveFolder)
    #Поле к кнопке Обзор
    $ApplyChangesArchiveFolderLabel = New-Object System.Windows.Forms.Label
    $ApplyChangesArchiveFolderLabel.Location =  New-Object System.Drawing.Point(95,78) #x,y
    $ApplyChangesArchiveFolderLabel.Width = 500
    $ApplyChangesArchiveFolderLabel.Text = "Сохранить результаты преобразования в..."
    $ApplyChangesArchiveFolderLabel.TextAlign = "TopLeft"
    $ApplyChangesForm.Controls.Add($ApplyChangesArchiveFolderLabel)
    #Кнопка Начать
    $ApplyChangesFormApplyButton = New-Object System.Windows.Forms.Button
    $ApplyChangesFormApplyButton.Location = New-Object System.Drawing.Point(10,190) #x,y
    $ApplyChangesFormApplyButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormApplyButton.Text = "Начать"
    $ApplyChangesFormApplyButton.Enabled = $true
    $ApplyChangesFormApplyButton.Add_Click({
    Transform-XmlFile -XmlPath $script:PathToXmlFile -XslPath $script:PathToXsltFile -HtmlOutput $script:PathToOutput
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormApplyButton)
    #Кнопка закрыть
    $ApplyChangesFormCancelButton = New-Object System.Windows.Forms.Button
    $ApplyChangesFormCancelButton.Location = New-Object System.Drawing.Point(100,190) #x,y
    $ApplyChangesFormCancelButton.Size = New-Object System.Drawing.Point(80,22) #width,height
    $ApplyChangesFormCancelButton.Text = "Закрыть"
    $ApplyChangesFormCancelButton.Add_Click({
        $ApplyChangesForm.Close()
    })
    $ApplyChangesForm.Controls.Add($ApplyChangesFormCancelButton)
    $ApplyChangesForm.ShowDialog()
}

ApplyChangesForm
