#global arrays
$fileNames = @()
$fileMD5s = @()
$fileFullNames = @()
#functions
Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
        If ($Show -eq "OK") {
        Return $objForm.SelectedPath
        } Else {
        Exit
        }
}
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "
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
<br>
<table>
<tr>
        <th>Document name</th>
        <th>Old version</th> 
        <th>MD5</th>
        <th>Action</th>
</tr>" -Encoding UTF8
#==========Statistics==========#
#script code
$path = Select-Folder -description "Specify a path to '# Source documents doc, docx, xls, xlsx' folder you want to compare new documents against"
Write-Host "Calculating MD5 checksums..."
Get-ChildItem -Path "$PSScriptRoot\KPD\# Source documents docx, doc, xls, xlsx\*.*" -Include "*.doc*", "*.xls*" | % {
$fileMD5s += (Get-FileHash -Path $_.FullName -Algorithm MD5).Hash
$fileNames += $_.Name
$fileFullNames += $_.FullName
}
$fileData = $fileNames, $fileMD5s, $fileFullNames
#compares
for ($i = 0; $i -lt $fileData[0].Length ; $i++) {
    $searchFor = $path + "\" + $fileData[0][$i]
    $existence = Test-Path -Path $searchFor
    $statFileName = $fileData[0][$i]
    if ($existence) {
        Write-Host "Comparing" $fileData[0][$i] "against its previous version..."
        Write-Host "Previous version found" -ForegroundColor Green
        $foundFileMD5 = (Get-FileHash -Path $searchFor -Algorithm MD5).Hash
        if ($fileData[1][$i] -eq $foundFileMD5) {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td>$statFileName</td>
        <td>Found</td>
        <td>Match</td>
        <td>Copied to '# Matches against previous version'</td>
</tr>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs match" -ForegroundColor Green
        Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Matches against previous version\" 
        } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td>$statFileName</td>
        <td>Found</td>
        <td>No match</td>
        <td>Copied to '# Source documents to be translated'</td>
</tr>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs DO NOT match" -ForegroundColor Red
        Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Source documents to be translated\"
        }
    } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td>$statFileName</td>
        <td>Not found</td>
        <td>---</td>
        <td>Copied to '# Source documents to be translated'</td>
</tr>" -Encoding UTF8
#==========Statistics==========#
    Write-Host "Previous version NOT found" -ForegroundColor Red
    Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Source documents to be translated\"
    }
Write-Host "-----------"
}
Add-Content "$PSScriptRoot\Comparison Report.html" "</table>
</div>
</body>
</html>" -Encoding UTF8
$fileNames = @()
$fileMD5s = @()
$fileFullNames = @()
