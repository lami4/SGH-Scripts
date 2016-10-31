clear
$var = "PABKRF"
$filesToBeTranslated = @(Get-ChildItem "C:\Users\Анник\Desktop\# projects\PABKRF 7.0.0\Source documents to be translated" | % {$_.BaseName})
$foldersToLoopThrough = @(Get-ChildItem "C:\Users\Анник\Desktop\# projects" | Where-Object {$_ -Match "$var .*\d$"} | Sort-Object -Descending)
for ($i = 0; $i -lt $filesToBeTranslated.Length; $i++)
{
    $currentFile = $filesToBeTranslated[$i]
    for ($t = 0; $t -lt $foldersToLoopThrough.Length; $t++)
        {
        $currentFolder = $foldersToLoopThrough[$t]
        $boolean = Test-Path "C:\Users\Анник\Desktop\# projects\$currentFolder\LiveDocs\$currentFile.txt"
        if ($boolean -eq $true) {
        #copy a shit file to a shit folder and breaks 
        break}
        }
    Write-Host $currentFile "found in" $currentFolder
}
