clear
function WrongInput ($inum, $enum)
{
Clear
Write-Host "Input a number in the range from $inum to $enum
"
UserPrompt
}

function ShortCutsFP ($doc)
{
$CurrentLocation = Get-Location
$TargetFile = "$CurrentLocation\$doc\# For publication"
$ProjectName = Split-Path $CurrentLocation -Leaf
$ShortcutFile = "Z:\OTD.Translate\Переводчики\!! for publication\$ProjectName $doc EN.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()
}

function ShortCutsI ($doc)
{
$CurrentLocation = Get-Location
$TargetFile = "$CurrentLocation\$doc\# Images"
$ProjectName = Split-Path $CurrentLocation -Leaf
$ShortcutFile = "Z:\OTD.Translate\Переводчики\# images\Images $ProjectName $doc.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()
}

function UserPrompt ()
{
Write-Host 'What do you want to get published?

Type 1 to publish KPD
Type 2 to publish RPD
'
$input = Read-Host 'Make your choice'
    if ($input -eq 1) {
    ShortCutsFP KPD
    ShortCutsI KPD
    } elseif ($input -eq 2) {
    ShortCutsFP RPD
    ShortCutsI RPD
    } else {
    WrongInput 1  2
    }
}
UserPrompt
