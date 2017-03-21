clear
$script:yesNoUserInput = 0
#functions
Function Select-Folder ($Description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
        If ($Show -eq "OK") {
        Return $objForm.SelectedPath
        } Else {
        Exit
        }
}

Function Input-YesOrNo ($Question, $BoxTitle) 
{
    $a = New-Object -ComObject wscript.shell
    $intAnswer = $a.popup($Question,0,$BoxTitle,4)
    If ($intAnswer -eq 6) {
        $script:yesNoUserInput = 1
    } else {
        $script:yesNoUserInput = 0
    }
}

Function Prompt-Overwrite ($FileName, $Path) 
{
    if (Test-Path -Path "$PSScriptRoot\KPD\# For publication\$FileName") {
        $nl = [System.Environment]::NewLine
        Input-YesOrNo -Question "$FileName already exists in '...\KPD\# For publication'.$nl`Do you want to overwrite it?"
        if ($script:yesNoUserInput -eq 1) {
        Copy-Item -Path "$Path\$FileName" -Destination "$PSScriptRoot\KPD\# For publication"  
        } else {return}
    } else {
    Copy-Item -Path "$Path\$FileName" -Destination "$PSScriptRoot\KPD\# For publication"
    }
}

$SelectedPath = Select-Folder -Description "Please specify a path to '# For publication' folder of the previous project."
Get-ChildItem -Path "$PSScriptRoot\KPD\# Matches against previous version" | % {
    $FileBaseName = ""
    if ($_.BaseName -match "GL-RU") {$FileBaseName = $_.BaseName -replace "GL-RU", "GL-EN"}
    if ($_.BaseName -match "RU-RU") {$FileBaseName = $_.BaseName -replace "RU-RU", "RU-EN"}
        Get-ChildItem -Path "$SelectedPath\$FileBaseName.*" | % {
        Prompt-Overwrite -FileName $_.Name -Path $SelectedPath
        }
}
