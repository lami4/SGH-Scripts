$script:yesNoUserInput = 0
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

Function Input-YesOrNo ($Question, $BoxTitle) {
$a = New-Object -ComObject wscript.shell
$intAnswer = $a.popup($Question,0,$BoxTitle,4)
If ($intAnswer -eq 6) {
  $script:yesNoUserInput = 1
}
}

Function Prompt-Overwrite ($FileName, $TestPath, $FileFullName, $Statistics) {
$check = Test-Path -Path "$TestPath\$FileName"
    if ($check -eq $false) {
    Copy-Item -Path $FileFullName -Destination "$TestPath"
    Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td>Copied to '$Statistics'</td>
</tr>"
    } else {
    Input-YesOrNo -Question "$FileName already exists in $TestPath. Do you want to overwrite it?"
        if ($script:yesNoUserInput -eq 1) {
        Copy-Item -Path $FileFullName -Destination "$TestPath"
        Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td>Copied to '$Statistics'</td>
</tr>"
        } else {
        Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td>Copying to '$Statistics' failed.<br>Overwrite cancelled by user.</td>
</tr>"
}
    $script:yesNoUserInput = 0
    }
}

#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "
<html lang=""en"">
<head>
<meta charset=""utf-8"">
<title>Compare and Distribute</title>
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
        <td>Match</td>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs match" -ForegroundColor Green
        Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Matches against previous version" -FileFullName $fileData[2][$i] -Statistics "# Matches against previous version"
        #Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Matches against previous version\" 
        } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td>$statFileName</td>
        <td>Found</td>
        <td>No match</td>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs DO NOT match" -ForegroundColor Red
        Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Source documents to be translated" -FileFullName $fileData[2][$i] -Statistics "# Source documents to be translated"
        #Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Source documents to be translated\"
        }
    } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td>$statFileName</td>
        <td>Not found</td>
        <td>---</td>" -Encoding UTF8
#==========Statistics==========#
    Write-Host "Previous version NOT found" -ForegroundColor Red
    Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Source documents to be translated" -FileFullName $fileData[2][$i] -Statistics "# Source documents to be translated"
    #Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Source documents to be translated\"
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
