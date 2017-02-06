clear
#global arrays
$fileNames = @()
$fileMD5s = @()
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
#script code
$path = Select-Folder -description "Specify a path to '# Source documents doc, docx, xls, xlsx' folder you want to compare updated files against"
Get-ChildItem "$PSScriptRoot\KPD\# Source documents doc, docx, xls, xlsx" | % {
$fileMD5s += (Get-FileHash -Path $_.FullName -Algorithm MD5).Hash
$fileNames += $_.BaseName
}
$fileData = $fileNames, $fileMD5s
#add for loop to check for files with the same name, but different extensions
for ($i = 0; $i -lt $fileData[0].Length; $i++) {
    $searchFor = $path + "\" + $fileData[0][$i] + ".*"
    $existence = Test-Path -Path $searchFor
    if ($existence) {
        Write-Host "File with the same name found"
        $foundFileMD5 = (Get-FileHash -Path $searchFor -Algorithm MD5).Hash
        if ($fileData[1][$i] -eq $foundFileMD5) {
        Write-Host "MDs match"
        } else {Write-Host "MDs DO NOT match"}
    } else {Write-Host "File with the same name NOT found"}
Write-Host "-----------"
}
$fileNames = @()
$fileMD5s = @()
