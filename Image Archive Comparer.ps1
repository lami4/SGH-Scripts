clear
#Global variables
$desktopPath = [Environment]::GetFolderPath("Desktop")
$folderWithProcessedDocuments = "Processed files"
$pathToImageStorage = "C:\Users\Анник\Desktop\Batch\# chest of images"

#Functions
Function Select-Folder
{
    param([string]$Description="Please, specify a path to a folder with MS Word files to be processed",[string]$RootFolder="Desktop")

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

Function Make-Choice($filesToBeCopied)
{
$a = new-object -comobject wscript.shell 
$errorBox = $a.popup("Folder named '$folderWithProcessedDocuments' already exists on your desktop!
Do you want to overwrite it?
Clicking 'No' will stop the script.
",0,"Delete Files",4) 
If ($errorBox -eq 6) { 
   Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments" -Recurse
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments" -type directory
   for ($i = 0; $i -lt $filesToBeCopied.Length; $i++)
   {
   $currentItem = $filesToBeCopied[$i]
   Copy-Item "$currentItem" "$desktopPath\$folderWithProcessedDocuments"
   }
} else { 
   Exit
} 
}

Function Unzip-Archive
{
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\*.zip" | % {
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
    $tarImg.save($targetImage, [System.Drawing.Imaging.ImageFormat]::Jpeg) 
    $srcImg.Dispose() 
    $tarImg.Dispose() 
}

Function Convert-Image ($pathToImageInStorage, $saveTo, $ImageFormat) {
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms"); 
$i = new-object System.Drawing.Bitmap($pathToImageInStorage); 
$i.Save($saveTo, $ImageFormat);
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
#Gets the list of unzipped documents
Get-ChildItem -Path "$desktopPath\$folderWithProcessedDocuments" -Directory | % {
    #Gets md5, image name, image extension, image full name and then adds them to appropriate arrays in each unzipped document one by one
    Get-FileHash -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\*" -Algorithm MD5 | % {
        $imageHash = $_.Hash
        $imageHashes += $imageHash
        $imageFullName = Split-Path $_.Path -Leaf
        $imageFullNames += $imageFullName
        $parsedImageFullName = $imageFullName -split "\."
        $imageName = $parsedImageFullName[0]
        $imageNames += $imageName
        $imageExtension = "." + $parsedImageFullName[1]
        $imageExtensions += $imageExtension
    }
    #Creates temporary folders
    New-Item "$desktopPath\$folderWithProcessedDocuments\Temporary" -Type directory
    New-Item "$desktopPath\$folderWithProcessedDocuments\Temporary WM" -Type directory
    New-Item "$desktopPath\$folderWithProcessedDocuments\Temporary zip" -Type directory
    #Joins together arrays in the multidimensional array called imageProperties
    $imageProperties = $imageHashes, $imageFullNames, $imageNames, $imageExtensions
    #Processes each image stored in 'imageProperties' array
    for ($i = 0; $i -lt $imageProperties[0].Length; $i++) 
        {
        #========================================================
        #Uncomment 4 strings below to check if parsing goes well on your PC
        #Write-Host "Image MD5 sum:" $imageProperties[0][$i]
        #Write-Host "Image full name:" $imageProperties[1][$i]
        #Write-Host "Image name:" $imageProperties[2][$i]
        #Write-Host "Image extension:" $imageProperties[3][$i]
        #Write-Host "Image extension without '.':" $imageProperties[4][$i]
        #Write-Host "-----next image-----"
        #========================================================
        $currentMD5 = $imageProperties[0][$i]
        $currentFullName = $imageProperties[1][$i]
        $currentName = $imageProperties[2][$i]
        $currentExtension = $imageProperties[3][$i]
        #Checks if a file in the image storage has a name eqalling MD5 sum of the image being processed
        $existenceInImageStorage = Test-Path -Path "$pathToImageStorage\*\$currentMD5.*"
        if ($existenceInImageStorage -eq $true)
            {
            Write-Host "EN file for" $currentFullName "was found in the image storage."
            #checks if extesions are equal: if yes, copies file to temporary folder and renames it as required; if no, changes extension as required and saves a file in Temporary folder using PNG format
            $imageInStorage = Get-ChildItem "$pathToImageStorage\*\$currentMD5.*"
            if ($currentExtension -eq $imageInStorage.Extension)
                {
                Write-Host "Extension matches! File from the storage was copied to Temporary folder and renamed."
                Copy-Item -Path "$imageInStorage" "$desktopPath\$folderWithProcessedDocuments\Temporary"
                $fileToBeRenamed = $imageInStorage.Name
                Rename-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary\$fileToBeRenamed" -NewName "$currentFullName"
                } else {
                Write-Host "Extension does not match! File from the storage was converted to match the required extension and then copied to Temporary folder"
                Convert-Image -pathToImageInStorage "$imageInStorage" -saveTo "$desktopPath\$folderWithProcessedDocuments\Temporary\$currentFullName" -ImageFormat "Jpeg"
                }
            } else {
            Copy-Item -Path "$desktopPath\$folderWithProcessedDocuments\$_\word\media\$currentFullName" "$desktopPath\$folderWithProcessedDocuments\Temporary WM"
            Write-Host "EN file for" $currentFullName "was NOT found in the image storage. Image will be watermarked."
            }
        }
    #Watermarks images in Temporary WM folder and copies them to Temporary folder
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary WM" | % {
    $currentImage = $_.Name
    Write-TextWaterMark -SourceImage "$desktopPath\$folderWithProcessedDocuments\Temporary WM\$currentImage" -TargetImage "$desktopPath\$folderWithProcessedDocuments\Temporary\$currentImage" -MessageText “*”

    }
    #Moves images from the current archive (word/media) to Temporary zip folder
    $currentDirectoryName = $_ 
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary" | % {
    $currentImageName = $_.Name
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\Temporary zip").MoveHere("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media\$currentImageName")
    }
    #Copies processed images to the current archive
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary" | % {
    $currentImageName = $_.Name
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media").CopyHere("$desktopPath\$folderWithProcessedDocuments\Temporary\$currentImageName")
    Start-Sleep -Seconds 1
    }
    Write-Host "--------LOOP FOR FILE STOPS HERE---------"
    #clears arrays
    $imageHashes = @()
    $imageFullNames = @()
    $imageNames = @()
    $imageExtensions = @()
    #removes temporary folders
    Start-Sleep -Seconds 1
    Remove-Item -Path "$desktopPath\$folderWithProcessedDocuments\Temporary", "$desktopPath\$folderWithProcessedDocuments\Temporary WM", "$desktopPath\$folderWithProcessedDocuments\Temporary zip" -Recurse
}
}

Function Get-WordFileExtension 
{
$documentNames = @()
$documentExtensions = @()
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\*.doc*" | % {
$parsedDocumentFullName = $_.Name -split "\."
$documentName = $parsedDocumentFullName[0]
$documentNames += $documentName
$documentExtension = $parsedDocumentFullName[1]
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
$folderPath = Select-Folder
#Gets files in the specified folder ignoring everything except for *.doc and *.docx files
$listOfFiles = @(Get-ChildItem -path "$folderPath\*.doc*")
#Checks if the 'desktop/$folderWithProcessedDocuments' folder exists and copies files to it
$processedFilesExistenceCheck = Test-Path -Path "$desktopPath\$folderWithProcessedDocuments"
if ($processedFilesExistenceCheck -eq $true)
{
   #If folder named '$folderWithProcessedDocuments' already exists on the user's desktop, the script asks for permition to overwrite it
   Make-Choice $listOfFiles
} else {
   #If folder named '$folderWithProcessedDocuments' does not exist, the script creates it and copies files to it
   New-Item -Path "$desktopPath\$folderWithProcessedDocuments" -type directory
   for ($i = 0; $i -lt $listOfFiles.Length; $i++)
   {
   $currentItem = $listOfFiles[$i]
   Copy-Item "$currentItem" "$desktopPath\$folderWithProcessedDocuments"
   }
}
#Memorises word document name to use it later for the renaming
Get-WordFileExtension
#Renames doc/docx as zip in '$folderWithProcessedDocuments'
Get-ChildItem -path "$desktopPath\$folderWithProcessedDocuments\*.doc*" | Rename-Item -NewName { [io.path]::ChangeExtension($_.name, "zip") }
#Unzips files in '$folderWithProcessedDocuments'
Unzip-Archive
#Gets all folders in '$folderWithProcessedDocuments', then computes MD5 sums in '*/media', checks them in 'chest of images' - hits are copied in zip files where no hits are water marked by *
Process-ImagesFromDocument > $null
#Deletes all folders in '$folderWithProcessedDocuments' folder
Start-Sleep -Seconds 1
Remove-Item "$desktopPath\$folderWithProcessedDocuments\*" -Exclude "*.zip" -Recurse
#Renames zip as doc/docx
RenameBack-ZipFile
Write-Host "Exiting the script"
Start-Sleep -Seconds 1
