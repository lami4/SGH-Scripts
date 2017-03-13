$script:yesNoUserInput = 0

Function Select-Folder ($description)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = "Desktop"
    $objForm.Description = $description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK") {Return $objForm.SelectedPath} Else {Exit}
}

Function Input-YesOrNo ($Question, $BoxTitle) 
{
    $a = New-Object -ComObject wscript.shell
    $intAnswer = $a.popup($Question,0,$BoxTitle,4)
    If ($intAnswer -eq 6) {$script:yesNoUserInput = 1} else {Exit}
}

$SelectedFolder = Select-Folder -description "Укажите папку, в которой необходимо снять контрольный суммы файлов."
if ((Test-Path -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt") -eq $true) {
$nl = [System.Environment]::NewLine
Input-YesOrNo  -Question "Список 'MD5 файлов текущего релиза.txt' уже сущетвует. Продолжить?$nl$nl`Да - перезаписать и продолжить исполнение скрипта.$nl`Нет - не перезаписывать и остановить исполнение скрипта.$nl$nl`Если вы не хотите перезаписывать существующий список, но хотите продолжить исполнение скрипта - переместите список из папки, где расположен файл скрипта, в любое удобное для вас место и нажмите 'Да'." -BoxTitle "Список уже существует"
if ($script:yesNoUserInput -eq 1) {Remove-Item -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt"}
$script:yesNoUserInput = 0
}
Get-ChildItem -Path "$SelectedFolder\*.*" -Exclude "*.pdf", "*.doc*", "*.xls*" | % {
Add-Content -Path "$PSScriptRoot\MD5 файлов текущего релиза.txt" -Value "$($_.Name):$((Get-FileHash -Path $_.FullName -Algorithm MD5).Hash)"
}
