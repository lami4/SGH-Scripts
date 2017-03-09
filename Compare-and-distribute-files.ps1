#global arrays and variables
$script:yesNoUserInput = 0
$fileNames = @()
$fileMD5s = @()
$fileFullNames = @()
$script:SourceDocumentsToBeTranslated = 0
$script:MathcesAgainstPreviousVersion = 0
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
    Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td id=""action"">Copied to '$Statistics'</td>
</tr>" -Encoding UTF8
    } else {
    $nl = [System.Environment]::NewLine
    Input-YesOrNo -Question "$FileName already exists in $TestPath.$nl`Do you want to overwrite it?"
        if ($script:yesNoUserInput -eq 1) {
        Copy-Item -Path $FileFullName -Destination "$TestPath"
        Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td id=""action"">Copied to '$Statistics'</td>
</tr>" -Encoding UTF8
        } else {
        Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "        <td id=""action"">Copying to '$Statistics' failed.<br>Overwrite cancelled by user.</td>
</tr>" -Encoding UTF8
}
    $script:yesNoUserInput = 0
    }
}

#script code
$path = Select-Folder -description "Specify a path to '# Source documents doc, docx, xls, xlsx' folder you want to compare new documents against"
Write-Host "Calculating MD5 checksums..."
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
	padding: 3px;
	border: 1px solid black;
    text-align:center;
    background-color: #FFC;
}
#indication, #action {
text-align: left;
}
</style>
<script>
function nameInputFilter() { 
  var input, filter, table, tr, td, i;
  input = document.getElementById(""nameInput"");
  filter = input.value.toUpperCase();
  table = document.getElementById(""Stats"");
  tr = table.getElementsByTagName(""tr"");
  for (i = 1; i < tr.length; i++) {
    td = tr[i].getElementsByTagName(""td"")[0];
    if (td) {
      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
        tr[i].style.display = """";
      } else {
        tr[i].style.display = ""none"";
      }
    } 
  }
}
function filter(filterid, columnnumber) {
    var filterValue = document.getElementById(filterid).value;
    table = document.getElementById(""Stats"");
    tr = table.getElementsByTagName(""tr"");
        if (filterValue == ""Discard filter"") {
            for (i = 1; i < tr.length; i++) {
                td = tr[i].getElementsByTagName(""td"")[columnnumber];
                if (td) {
                tr[i].style.display = """";
                }  
            }
        } else {
            for (i = 1; i < tr.length; i++) {
                td = tr[i].getElementsByTagName(""td"")[columnnumber];
                    if (td) {
                        if (td.innerText == filterValue) {
        		tr[i].style.display = """";
                        } else {
                        tr[i].style.display = ""none"";
                        }
                    } 
             }
        }
}
</script>
</head>
<body>
<div>
<h3>Hello.</h3>
<br>
<table id=""Stats"">
<tr>
        <td id=""indication"">
	        <input type=""text"" id=""nameInput"" onkeyup=""nameInputFilter()"" placeholder=""Search for names..."">
	    </td>
        <td>
            <select id=""OldVersionFilter"" onchange=""filter('OldVersionFilter', 1)"">
            <option value=""Discard filter"">Discard filter</option>
            <option value=""Found"">Found</option>
            <option value=""Not found"">Not found</option>
            </select>
	    </td>
        <td>
	        <select id=""MD5filter"" onchange=""filter('MD5filter', 2)"">
            <option value=""Discard filter"">Discard filter</option>
            <option value=""Match"">Match</option>
            <option value=""No match"">No match</option>
            </select>
	    </td>
        <td>This cell is so lovely.</td>
</tr>
<tr>
        <th>Document name</th>
        <th>Old version</th> 
        <th>MD5</th>
        <th>Action</th>
</tr>" -Encoding UTF8
#==========Statistics==========#
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
    Start-Sleep -Milliseconds 50
    if ($existence) {
        Write-Host "Comparing" $fileData[0][$i] "against its previous version..."
        Write-Host "Previous version found" -ForegroundColor Green
        $foundFileMD5 = (Get-FileHash -Path $searchFor -Algorithm MD5).Hash
        if ($fileData[1][$i] -eq $foundFileMD5) {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td id=""indication"">$statFileName</td>
        <td><font color=""green""><b>Found</b></font></td>
        <td><font color=""green""><b>Match</b></font></td>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs match" -ForegroundColor Green
        Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Matches against previous version" -FileFullName $fileData[2][$i] -Statistics "# Matches against previous version"
        $script:MathcesAgainstPreviousVersion += 1
        #Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Matches against previous version\" 
        } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td id=""indication"">$statFileName</td>
        <td><font color=""green""><b>Found</b></font></td>
        <td><font color=""red""><b>No match</b></font></td>" -Encoding UTF8
#==========Statistics==========#
        Write-Host "MDs DO NOT match" -ForegroundColor Red
        Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Source documents to be translated" -FileFullName $fileData[2][$i] -Statistics "# Source documents to be translated"
        $script:SourceDocumentsToBeTranslated += 1
        #Copy-Item -Path $fileData[2][$i] -Destination "$PSScriptRoot\KPD\# Source documents to be translated\"
        }
    } else {
#==========Statistics==========#
Add-Content -Path "$PSScriptRoot\Comparison Report.html" -Value "<tr>
        <td id=""indication"">$statFileName</td>
        <td><font color=""red""><b>Not found</b></font></td>
        <td>---</td>" -Encoding UTF8
#==========Statistics==========#
    Write-Host "Previous version NOT found" -ForegroundColor Red
    Prompt-Overwrite -FileName $fileData[0][$i] -TestPath "$PSScriptRoot\KPD\# Source documents to be translated" -FileFullName $fileData[2][$i] -Statistics "# Source documents to be translated"
    $script:SourceDocumentsToBeTranslated += 1
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
$totalCount = $script:SourceDocumentsToBeTranslated + $script:MathcesAgainstPreviousVersion
$string = "<h3>Hello.</h3>`r`n<h4>Documents processed: $totalCount</h4>`r`n<h4>Copied to '# Source documents to be translated': $script:SourceDocumentsToBeTranslated</h4>`r`n<h4>Copied to '# Matches against previous version': $script:MathcesAgainstPreviousVersion</h4>"
(Get-Content -Path "$PSScriptRoot\Comparison Report.html").Replace("<h3>Hello.</h3>", $string) | Set-Content("$PSScriptRoot\Comparison Report.html")
Invoke-Item "$PSScriptRoot\Comparison Report.html"
