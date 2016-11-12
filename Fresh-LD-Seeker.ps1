clear
Function Select-FolderDialog
{
    param([string]$Description="Select Folder",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Exit
        }
    }

$parsedCurrentProjectPath = ($PSScriptRoot | Split-Path -Leaf) -split " "
$projectPath = $parsedCurrentProjectPath[0]
$table = @()
$pathToFilesToBeTranslated = Select-FolderDialog
$filesToBeTranslated = @(Get-ChildItem "$pathToFilesToBeTranslated" | % {$_.BaseName})
$foldersToLoopThrough = @(Get-ChildItem "C:\Users\Анник\Desktop\# projects" | Where-Object {$_ -Match "$projectPath .*\d$"} | Sort-Object -Descending)

Add-Content "$PSScriptRoot\LiveDoc Report.html" '<!doctype html>
<html lang="en">

<head>
<meta charset="utf-8">
<title>LiveDoc Report</title>
</head>

<body>
<h3>Hello. Here is what I managed to find.</h3>
<table style="width:50%" border=1px align=center>
<tr>
        <th>Document</th>
        <th>Status</th> 
        <th>Found In</th>
</tr>'

for ($i = 0; $i -lt $filesToBeTranslated.Length; $i++)
{
    $currentFile = $filesToBeTranslated[$i]
    for ($t = 0; $t -lt $foldersToLoopThrough.Length; $t++)
        {
        $currentFolder = $foldersToLoopThrough[$t]
        $boolean = Test-Path "C:\Users\Анник\Desktop\# projects\$currentFolder\LiveDocs\$currentFile.txt"
        if ($boolean -eq $true) {
        #copy a shit file to a shit folder and breaks
        $table += @{Document="$currentFile"; Status="FOUND"; FoundIn="$currentFolder"}
        Add-Content "$PSScriptRoot\LiveDoc Report.html" "<tr>
        <td><font color=""black"">$currentFile</font></td>
        <td><font color=""green""><b>FOUND</b></font></td>
        <td><font color=""black"">$currentFolder</font></td>
</tr>"
        break}
        }
        if ($t -eq $foldersToLoopThrough.Length) {
        $table += @{Document="$currentFile"; Status="NOT FOUND"; FoundIn="-NONE-"}
        Add-Content "$PSScriptRoot\LiveDoc Report.html" "<tr>
        <td><font color=""black"">$currentFile</font></td>
        <td><font color=""red""><b>NOT FOUND</b></font></td>
        <td><font color=""black"">-none-</font></td>
</tr>"
        }
} 
$table.ForEach({[PSCustomObject]$_}) | Format-Table Document, Status, FoundIn -AutoSize

Add-Content "$PSScriptRoot\LiveDoc Report.html" "</table>
</body>
</html>"
Invoke-Item "$PSScriptRoot\LiveDoc Report.html"
