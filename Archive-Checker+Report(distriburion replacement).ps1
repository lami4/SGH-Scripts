#Global variables
$desktopPath = [Environment]::GetFolderPath("Desktop")
$folderWithProcessedDocuments = "Processed documents"
$pathToImageStorage = "C:\Users\Анник\Desktop\2\# chest of images"
$folderWithOldDocuments = "Source documents from previous version"
#Filter (images that are less than specified values will not be watermarked)
$imageWidth = 110
$imageHeight = 40

#Functions
Function Replace-FilesInArchive ($currentDirectoryName)
{
    #Moves files from the current archive to the "Temporary zip" folder
    Write-Host "Removing original images from the archive..."
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\Temporary zip").MoveHere("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media")
    Start-Sleep -Seconds 5
    Write-Host "Copying watermarked and translated images to the archive..."
    if (Test-Path -Path "$desktopPath\$folderWithProcessedDocuments\media\Thumbs.db") {Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments\media\Thumbs.db" -Force}
    #Copies processed files to now empty "media" folder in archive
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word").CopyHere("$desktopPath\$folderWithProcessedDocuments\media")
    Start-Sleep -Seconds 5
}

Function Select-Folder ($description)
{
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     
$objForm = New-Object System.Windows.Forms.FolderBrowserDialog
$objForm.Rootfolder = "Desktop"
$objForm.Description = $description
$Show = $objForm.ShowDialog()
If ($Show -eq "OK")
    {
    Return $objForm.SelectedPath
    } Else {
    Exit
    }
}

Function Make-Choice ($newFiles, $oldFiles)
{
$a = new-object -comobject wscript.shell 
$errorBox = $a.popup("Folder named '$folderWithProcessedDocuments' already exists on your desktop!
Do you want to overwrite it?
Clicking 'No' will stop the script.
",0,"Delete Files",4) 
If ($errorBox -eq 6) {
   Write-Host "Copying documents to the ""$folderWithProcessedDocuments"" folder..."
   Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments" -Recurse -Force
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments" -type directory
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments" -type directory > $null
   Get-ChildItem -Path "$newFiles\*.doc*" | % {
   Copy-Item -Path $_.FullName -Destination "$desktopPath\$folderWithProcessedDocuments"
   $currentBaseName = $_.BaseName
   Copy-Item -Path "$oldFiles\$currentBaseName.*" -Destination "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments"
   }
} else { 
   Exit
} 
}

Function Unzip-Archive ($folderName)
{
Get-ChildItem -path "$desktopPath\$folderName\*.zip" | % {
$destination = join-path $_.DirectoryName  $_.BaseName
New-Item $destination -type directory
$shell = new-object -com shell.application
$zip = $shell.NameSpace("$_")
    foreach($item in $zip.items())
    {
    $shell.Namespace("$destination").copyhere($item)
    }
}
}

Function Write-TextWaterMark 
{
   [CmdletBinding()] 
 
   Param ( 
 
      [Parameter( 
      ValueFromPipeline=$False, 
      Mandatory=$True, 
      HelpMessage="A path to original image")] 
      [string]$SourceImage, 
       
      [Parameter( 
      ValueFromPipeline=$False, 
      Mandatory=$True, 
      HelpMessage="A path to target image")] 
      [string]$TargetImage, 
       
      [Parameter( 
      ValueFromPipeline=$False, 
      Mandatory=$True, 
      HelpMessage="Text to write on image")] 
      [string]$MessageText 
 
      ) 
 
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null 
 
    #read source image and create new target image 
    $srcImg = [System.Drawing.Image]::FromFile($SourceImage) 
    $tarImg = new-object System.Drawing.Bitmap([int]($srcImg.width)),([int]($srcImg.height)) 
 
    #Intialize Graphics 
    $Image = [System.Drawing.Graphics]::FromImage($tarImg) 
    $Image.SmoothingMode = "AntiAlias" 
 
    $Rectangle = New-Object Drawing.Rectangle 0, 0, $srcImg.Width, $srcImg.Height 
    $Image.DrawImage($srcImg, $Rectangle, 0, 0, $srcImg.Width, $srcImg.Height, ([Drawing.GraphicsUnit]::Pixel)) 
 
    #Write MessageText (10 in from left, 1 down from top, white semi transparent text) 
    $Font = new-object System.Drawing.Font("Verdana", 24) 
    $Brush = New-Object Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(100, 255, 0,255)) 
    $Image.DrawString($MessageText, $Font, $Brush, 10, 1) 
     
    #Save and close the files 
    $tarImg.save($targetImage, [System.Drawing.Imaging.ImageFormat]::png) 
    $srcImg.Dispose() 
    $tarImg.Dispose() 
}

#Get hashes of images in each unzipped document one by one, looks for images with the same hash in 'chest of images'
#If image with the same hash not found, make watermarks
Function Process-ImagesFromDocument
{
#Arrays that will keep md5, image full name (image3.png), image name (image3) and image extension (.png),
$imageHashes = @()
$imageFullNames = @()
$imageNames = @()
$imageExtensions = @()
$imageFullPaths = @()
$oldImageHashes = @()
#========Statistics========
$oldDocumentExistence = $false
$totalRepetitionsCount = 0
$filteredRepetitionsCount = 0
$totalNumberOfImagesInDocument = 0
$totalNumberOfFilterdImagesInDocument = 0
$totalNumberOfLootedImages = 0
$totaNumberOfFilteredLootedImages = 0
$totalNumberOfFilteredRepetitionInAnalysis = 0
#========Statistics========
#Creates Count table
Add-Content "$PSScriptRoot\Test Report.html" "
<h3>Counts</h3>
<table>
    <tr>
        <th>Document</th>
        <th>Repetition</th>
        <th>Looted</th>
        <th>Total</th>
        <th>Repetition (Filtered)</th>
        <th>Looted (Filtered)</th>
        <th>Total (Filtered)</th>
    </tr>"
#Creates Analysis table
Add-Content "$PSScriptRoot\Analysis.html" "
<h3>Analysis</h3>
<table>
    <tr>
        <th>Document</th>
        <th>Looted</th>
        <th>Repitition</th>
        <th>New</th>
        <th>Total</th>
    </tr>"
#Gets the list of unzipped documents
Get-ChildItem -Path "$desktopPath\$folderWithProcessedDocuments" -Directory | Where-Object {$_ -NotMatch "$folderWithOldDocuments"} | % {
    #Gets md5, image name, image extension, image full name and then adds them to appropriate arrays in each unzipped document one by one
    $currentDirectory = $_
    Write-Host "==============================================================================="
    Write-Host "Just started working on $_..."

    #========Statistics========
    #Checks if the document existed in the previous version
    $oldDocumentExistence = Test-Path -Path "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments\$_"
    Write-Host $oldDocumentExistence
    if ($oldDocumentExistence -eq $true) {
        Get-FileHash -Path "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments\$_\word\media\*" -Algorithm MD5 | % {
        $oldImageHash = $_.Hash
        $oldImageHashes += $oldImageHash
        }
    }
    #Total number of images in a document
    $totalNumberOfImagesInDocument = (Get-ChildItem -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\*" -Exclude "*.wdp").Count
    #Total number of filtered images in a document
    Get-ChildItem -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\*" -Exclude "*.wdp" | % {
        [int]$currentWidthForTotal = magick identify -ping -format "%w" $_.FullName
        [int]$currentHeightForTotal = magick identify -ping -format "%h" $_.FullName
        if ($currentWidthForTotal -gt $imageWidth -and $currentHeightForTotal -gt $imageHeight) {
        $totalNumberOfFilterdImagesInDocument += 1
        }
    }
    #========Statistics========
        Get-FileHash -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\*" -Algorithm MD5 | % {
        $imageHash = $_.Hash
        $imageHashes += $imageHash
        $imageFullName = Split-Path $_.Path -Leaf
        $imageFullNames += $imageFullName
        $imageName = [IO.Path]::GetFileNameWithoutExtension($_.Path)
        $imageNames += $imageName
        $imageFullPath = $_.Path
        $imageFullPaths += $imageFullPath
        $imageExtension = [IO.Path]::GetExtension($_.Path)
        $imageExtensions += $imageExtension
        #add width and height to arrays!!!!!
    }
    #Creates temporary folders
    New-Item "$desktopPath\$folderWithProcessedDocuments\Temporary", "$desktopPath\$folderWithProcessedDocuments\Temporary WM", "$desktopPath\$folderWithProcessedDocuments\Temporary zip"  -Type directory
    New-Item "$desktopPath\$folderWithProcessedDocuments\Temporary bmp", "$desktopPath\$folderWithProcessedDocuments\Temporary bmp for WM", "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp" -Type directory
    #Joins together arrays in the multidimensional array called imageProperties
    $imageProperties = $imageHashes, $imageFullNames, $imageNames, $imageExtensions, $imageFullPaths
    Write-Host "Searching for translated images in the chest..."
    #Processes each image stored in 'imageProperties' array
    for ($i = 0; $i -lt $imageProperties[0].Length; $i++) 
        {
        $currentMD5 = $imageProperties[0][$i]
        $currentFullName = $imageProperties[1][$i]
        $currentName = $imageProperties[2][$i]
        $currentExtension = $imageProperties[3][$i]
        $currentFullPath = $imageProperties[4][$i]

        #========Statistics========
        #checks total repetitions
        if ($oldDocumentExistence -eq $true) {
            if ($currentExtension -ne ".wdp") {
                if ($oldImageHashes -contains $currentMD5) {
                    Write-Host "Repitition"
                    $totalRepetitionsCount += 1
                    } else {
                    Write-Host "No repitition"
                }
            }
        }
        #checks true repetitions
        if ($oldDocumentExistence -eq $true) {
            if ($currentExtension -ne ".wdp") {
                if ($oldImageHashes -contains $currentMD5) {
                    [int]$currentWidthForTrueRepetitions = magick identify -ping -format "%w" $currentFullPath
                    [int]$currentHeightForTrueRepetitions = magick identify -ping -format "%h" $currentFullPath
                    if ($currentWidthForTrueRepetitions -lt $imageWidth -and $currentHeightForTrueRepetitions -lt $imageHeight) {
                    Write-Host "Little Repititon Found"
                    } else {
                    Write-Host "Big Repition Found"
                    $filteredRepetitionsCount += 1
                    }
                }
            }
        }
        #========Statistics========

        #Checks if a file in the image storage has a name eqalling MD5 sum of the image being processed
        $existenceInImageStorage = Test-Path -Path "$pathToImageStorage\*\$currentMD5.*"
        
        #========Statistics========
            if ($currentExtension -ne ".wdp") {
                if ($existenceInImageStorage -eq $true) {
                    #Adds +1 to total number of looted images
                    $totalNumberOfLootedImages += 1
                    } else {
                    #Adds +1 to filtered repitions in analysis
                    if ($oldImageHashes -contains $currentMD5) {
                        [int]$currentWidthForTrueLootedAnalysis = magick identify -ping -format "%w" $currentFullPath
                        [int]$currentHeightForTrueLootedAnalysis = magick identify -ping -format "%h" $currentFullPath
                        if ($currentWidthForTrueLootedAnalysis -gt $imageWidth -and $currentHeightForTrueLootedAnalysis -gt $imageHeight) {
                            $totalNumberOfFilteredRepetitionInAnalysis += 1
                        }
                    }
                }
            }
            if ($currentExtension -ne ".wdp") {
                #Adds +1 to total number of filtered looted images
                if ($existenceInImageStorage -eq $true) {
                    [int]$currentWidthForTrueLooted = magick identify -ping -format "%w" $currentFullPath
                    [int]$currentHeightForTrueLooted = magick identify -ping -format "%h" $currentFullPath
                    if ($currentWidthForTrueLooted -gt $imageWidth -and $currentHeightForTrueLooted -gt $imageHeight) {
                        $totaNumberOfFilteredLootedImages += 1
                    }
                }
            }
        #========Statistics========

        if ($existenceInImageStorage -eq $true)
            {
            #Write-Host "EN file for" $currentFullName "was found in the image storage."
            #checks if extesions are equal: if yes, copies file to temporary folder and renames it as required; if no, changes extension as required and saves a file in Temporary folder using PNG format
            $imageInStorage = Get-ChildItem "$pathToImageStorage\*\$currentMD5.*"
            if ($currentExtension -eq $imageInStorage[0].Extension)
                {
                #Write-Host "Extension matches! File from the storage was copied to Temporary folder and renamed."
                Copy-Item -Path $imageInStorage[0] "$desktopPath\$folderWithProcessedDocuments\Temporary"
                $fileToBeRenamed = $imageInStorage[0].Name
                Rename-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary\$fileToBeRenamed" -NewName "$currentFullName"
                } else {
                    if ($currentExtension -eq ".wdp") {
                    #convert to bmp first as WDP converter can work only with BMP files
                    magick convert $imageInStorage[0] "$desktopPath\$folderWithProcessedDocuments\Temporary bmp\$currentName.bmp"
                    #then BMP files is converted to WDP
                    Start-Process -FilePath 'C:\WDP Converter\JXREncApp\x64\JXREncApp.exe' -ArgumentList "-i ""$desktopPath\$folderWithProcessedDocuments\Temporary bmp\$currentName.bmp"" -o ""$desktopPath\$folderWithProcessedDocuments\Temporary\$currentName.wdp"" -c 0"
                    } else {
                    #Write-Host "Extension does not match! File from the storage was converted to match the required extension and then copied to Temporary folder"
                    magick convert $imageInStorage[0] "$desktopPath\$folderWithProcessedDocuments\Temporary\$currentFullName"
                    }
                    }
            } else {
            if ($currentExtension -eq ".wdp") {
            Start-Process -FilePath 'C:\WDP Converter\JXRDecApp\x64\JXRDecApp.exe' -ArgumentList "-i ""$desktopPath\$folderWithProcessedDocuments\$_\word\media\$currentFullName"" -o ""$desktopPath\$folderWithProcessedDocuments\Temporary bmp for WM\$currentName.bmp"" -c 0"
            } else {
            Copy-Item -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\$currentFullName" "$desktopPath\$folderWithProcessedDocuments\Temporary WM"
            #Write-Host "EN file for" $currentFullName "was NOT found in the image storage. Image will be watermarked."
            }
            }
        }
    Write-Host "Watermarking images that were not found in the chest..."
    #Watermarks images in Temporary WM folder and copies them to Temporary folder
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary WM" | % {
        [int]$width = magick identify -ping -format "%w" $_.FullName
        [int]$height = magick identify -ping -format "%h" $_.FullName
        if ($width -lt $imageWidth -and $height -lt $imageHeight) { 
        Copy-Item -Path $_.FullName -Destination "$desktopPath\$folderWithProcessedDocuments\Temporary"
        Write-Host "Little image copied to temporary"
        } else {
        $currentImage = $_.Name
        Write-TextWaterMark -SourceImage "$desktopPath\$folderWithProcessedDocuments\Temporary WM\$currentImage" -TargetImage "$desktopPath\$folderWithProcessedDocuments\Temporary\$currentImage" -MessageText “*”
        }
    }
    #Watermarks bmp images in Temporary for WM folder and copies them to Temporary marked bmp folder
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary bmp for WM" | % {
        [int]$width = magick identify -ping -format "%w" $_.FullName
        [int]$height = magick identify -ping -format "%h" $_.FullName
        if ($width -lt $imageWidth -and $height -lt $imageHeight) { 
        Copy-Item -Path $_.FullName -Destination "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp"
        Write-Host "Little WDP image copied to temporary"
        } else {
        $currentImage = $_.Name
        Write-TextWaterMark -SourceImage "$desktopPath\$folderWithProcessedDocuments\Temporary bmp for WM\$currentImage" -TargetImage "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp\$currentImage" -MessageText “*”
        }
    }
    #Renames marked bmp images back to WDP and copies them to Temporary folder
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp\*.bmp" | % {
    $baseName = $_.BaseName
    Rename-Item -Path $_.FullName -NewName "$baseName.wdp"
    Copy-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp\$baseName.wdp" -Destination "$desktopPath\$folderWithProcessedDocuments\Temporary"
    }
    #Renames the Temporary folder to media
    Rename-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary" -NewName "media"
    #Moves images from the current archive (word/media) to Temporary zip folder
    Start-Sleep -Seconds 5
    Replace-FilesInArchive -currentDirectoryName "$_"
    Start-Sleep -Seconds 5
    Write-Host "Document processed."

    #========Statistics========
    $completelyNewImages = $totalNumberOfFilterdImagesInDocument - $totaNumberOfFilteredLootedImages - $totalNumberOfFilteredRepetitionInAnalysis
    #Adds collected statistics to Counts table
    Add-Content "$PSScriptRoot\Test Report.html" "
    <tr>
        <td>$currentDirectory</td>
        <td>$totalRepetitionsCount</td>
        <td>$totalNumberOfLootedImages</td>
        <td>$totalNumberOfImagesInDocument</td>
        <td>$filteredRepetitionsCount</td>
        <td>$totaNumberOfFilteredLootedImages</td>
        <td>$totalNumberOfFilterdImagesInDocument</td>
    </tr>"
    #Adds collected statistics to Analysis table
    Add-Content "$PSScriptRoot\Analysis.html" "
    <tr>
        <td>$currentDirectory</td>
        <td>$totaNumberOfFilteredLootedImages</td>
        <td>$totalNumberOfFilteredRepetitionInAnalysis</td>
        <td>$completelyNewImages</td>
        <td>$totalNumberOfFilterdImagesInDocument</td>
    </tr>"
    #$currentDirectory contains $totalNumberOfFilterdImagesInDocument images: $totaNumberOfFilteredLootedImages - looted from chest, $totalNumberOfFilteredRepetitionInAnalysis - repetitions, $completelyNewImages - completely new images"
    #========Statistics========
    #clears arrays and resets boolean values
    $imageHashes = @()
    $imageFullNames = @()
    $imageNames = @()
    $imageExtensions = @()
    $imageFullPaths = @()
    $oldImageHashes = @()

    #========Statistics========
    $oldDocumentExistence = $false
    $totalRepetitionsCount = 0
    $filteredRepetitionsCount = 0
    $totalNumberOfImagesInDocument = 0
    $totalNumberOfFilterdImagesInDocument = 0
    $totalNumberOfLootedImages = 0
    $totaNumberOfFilteredLootedImages = 0
    $totalNumberOfFilteredRepetitionInAnalysis = 0
    #========Statistics========

    #removes temporary folders
    Start-Sleep -Seconds 5
    Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments\media", "$desktopPath\$folderWithProcessedDocuments\Temporary WM" -Recurse -Force
    Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary bmp", "$desktopPath\$folderWithProcessedDocuments\Temporary bmp for WM", "$desktopPath\$folderWithProcessedDocuments\Temporary marked bmp" -Recurse -Force
    Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary zip" -Recurse -Force
    Write-Host "==============================================================================="
}
}


Function Get-WordFileExtension 
{
$documentNames = @()
$documentExtensions = @()
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\*.doc*" | % {
$documentName = [IO.Path]::GetFileNameWithoutExtension($_)
$documentNames += $documentName
$documentExtension = [IO.Path]::GetExtension($_).Trim(".")
$documentExtensions += $documentExtension
}
$global:documentNameForRenaming = $documentNames, $documentExtensions
}

Function RenameBack-ZipFile 
{
for ($i = 0; $i -lt $documentNameForRenaming[0].Length; $i++) {
$zipToBeRenamed = $documentNameForRenaming[0][$i]
$documentWordExtension = $documentNameForRenaming[1][$i]
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\$zipToBeRenamed.zip" | Rename-Item -NewName { [io.path]::ChangeExtension($_.name, "$documentWordExtension") }
}
}

#Gets path to a folder with files to be processed
$folderWithNewFiles = Select-Folder -description "Please, specify a path to a folder with MS Word files to be processed."
#Gets path to a folder with files to be processed
$folderWithOldFiles = Select-Folder -description "Please, specify a path to a '# Source documents doc, docx, xls, xlsx' folder of the previous project"
#Checks if the 'desktop/$folderWithProcessedDocuments' folder exists and copies files to it
$processedFilesExistenceCheck = Test-Path -Path "$desktopPath\$folderWithProcessedDocuments"
if ($processedFilesExistenceCheck -eq $true)
{
   #If folder named '$folderWithProcessedDocuments' already exists on the user's desktop, the script asks for permition to overwrite it
   Make-Choice -newFiles $folderWithNewFiles -oldFiles $folderWithOldFiles > $null
} else {
   #If folder named '$folderWithProcessedDocuments' does not exist, the script creates it and copies files to it
   Write-Host "Copying documents to the ""$folderWithProcessedDocuments"" folder..."
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments" -type directory > $null
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments" -type directory > $null
   Get-ChildItem -Path "$folderWithNewFiles\*.doc*" | % {
   Copy-Item -Path $_.FullName -Destination "$desktopPath\$folderWithProcessedDocuments"
   $currentBaseName = $_.BaseName
   Copy-Item -Path "$folderWithOldFiles\$currentBaseName.*" -Destination "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments"
   }
}
#Memorises word document name to use it later for the renaming
Get-WordFileExtension

#========Statistics========
#Creates html report
Add-Content "$PSScriptRoot\Test Report.html" "<!DOCTYPE html>
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
    text-align:center;
    background-color: #FFC;
}
</style>
</head>
<body>
<div>
<h3>Hello.</h3>
<h3>Here is some statistics...</h3>"
#========Statistics========

#Renames doc/docx as zip in '$folderWithProcessedDocuments'
$totalFiles = @(Get-ChildItem -Path "$desktopPath\$folderWithProcessedDocuments\*.doc*")
Write-Host "There are" $totalFiles.Length "document(s) to process." 
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\*.doc*" | Rename-Item -NewName { [io.path]::ChangeExtension($_.name, "zip") } > $null
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\$folderWithOldDocuments\*.doc*" | Rename-Item -NewName { [io.path]::ChangeExtension($_.name, "zip") } > $null
#Unzips files in '$folderWithProcessedDocuments'
Write-Host "Renaming and unzipping them..."
Unzip-Archive -folderName "$folderWithProcessedDocuments" > $null
Unzip-Archive -folderName "$folderWithProcessedDocuments\$folderWithOldDocuments" > $null
#Gets all folders in '$folderWithProcessedDocuments', then computes MD5 sums in '*/media', checks them in 'chest of images' - hits are copied in zip files where no hits are water marked by *
Process-ImagesFromDocument > $null

#========Statistics========
Add-Content "$PSScriptRoot\Test Report.html" "</table>
<br>"
Add-Content "$PSScriptRoot\Analysis.html" "</table>"
Get-Content -Path "$PSScriptRoot\Analysis.html" | Add-Content -Path "$PSScriptRoot\Test Report.html"
#========Statistics========

#Deletes all folders in '$folderWithProcessedDocuments' folder
Start-Sleep -Seconds 1
Remove-Item "$desktopPath\$folderWithProcessedDocuments\*" -Exclude "*.zip" -Force -Recurse
#Renames zip as doc/docx
RenameBack-ZipFile
Add-Content "$PSScriptRoot\Test Report.html" "
</div>
</body>
</html>"
Write-Host "Exiting the script"
Start-Sleep -Seconds 1
