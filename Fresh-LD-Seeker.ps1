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

    Function Make-Choice
{
$a = new-object -comobject wscript.shell 
$errorBox = $a.popup("LiveDoc Report.html already exists in $PSScriptRoot!
Do you want to overwrite it?
Clicking 'No' will stop the script.
",0,"Delete Files",4) 
If ($errorBox -eq 6) {
Remove-Item -Path "$PSScriptRoot\LiveDoc Report.html"} else {Exit} 
}

$parsedCurrentProjectPath = ($PSScriptRoot | Split-Path -Leaf) -split " "
$projectPath = $parsedCurrentProjectPath[0]
$table = @()
$pathToFilesToBeTranslated = Select-FolderDialog
$reportExistenceCheck = Test-Path "$PSScriptRoot\LiveDoc Report.html"
if ($reportExistenceCheck -eq $true) {
Make-Choice
}
$filesToBeTranslated = @(Get-ChildItem "$pathToFilesToBeTranslated" | % {$_.BaseName})
$foldersToLoopThrough = @(Get-ChildItem "C:\Users\Анник\Desktop\# projects" | Where-Object {$_ -Match "$projectPath .*\d$"} | Sort-Object -Descending)

Add-Content "$PSScriptRoot\LiveDoc Report.html" "<!DOCTYPE html>
<html lang=""en"">

<head>
<meta charset=""utf-8"">
<title>LiveDoc Report</title>
<style type=""text/css"">
   div {
    font-family: Verdana, Arial, Helvetica, sans-serif;
   }

table {
    border-collapse: collapse;
}

table, td, th {
    border: 1px solid black;
    padding: 3px;
}
td {
    background-color: #FFC;
}
</style>
</head>

<body>
<div>
<h3>Hello.</h3>
<h3>I searched through the following folders:</h3>
<ul style=""list-style-type:square"">"
$foldersToLoopThrough | % {Add-Content "$PSScriptRoot\LiveDoc Report.html" "<li>$_</li>"}
Add-Content "$PSScriptRoot\LiveDoc Report.html" "</ul>
<h3>Here is what I managed to find.</h3>
<table>
<tr>
        <th>Document</th>
        <th>Status</th> 
        <th>Found In</th>
</tr>"

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
        <td align=""center""><font color=""black"">-none-</font></td>
</tr>"
        }
} 
$table.ForEach({[PSCustomObject]$_}) | Format-Table Document, Status, FoundIn -AutoSize

Add-Content "$PSScriptRoot\LiveDoc Report.html" "</table>
</div>
</body>
</html>"
Invoke-Item "$PSScriptRoot\LiveDoc Report.html"
