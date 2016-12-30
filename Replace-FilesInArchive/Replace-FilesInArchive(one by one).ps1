Function Replace-FilesInArchive ($currentDirectoryName)
{
    #Creates temporary *.txt file to prevent the "media" folder from being delete after the script deletes the last file in it
    New-Item -Path "$desktopPath\$folderWithProcessedDocuments\temporary.txt" -ItemType "file"
Start-Sleep -Seconds 1
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media").MoveHere("$desktopPath\$folderWithProcessedDocuments\temporary.txt")
Start-Sleep -Seconds 2
    #Moves files from the current archive to the "Temporary zip" folder
    Write-Host "Removing original images from the archive..."
    $filesInTemporary = Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary"
    Write-Host $filesInTemporary.Count
    Write-Host $filesInTemporary
    Start-Sleep -Seconds 5
    $filesInTemporary | % {
    Start-Sleep -Seconds 1
    $currentImageNameMove = $_.Name
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\Temporary zip").MoveHere("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media\$currentImageNameMove")
    }
    Write-Host "Copying watermarked and translated images to the archive..."
Start-Sleep -Seconds 5
    #Copies processed files to now empty "media" folder in archive
    Get-ChildItem "$desktopPath\$folderWithProcessedDocuments\Temporary" | % {
    $currentImageNameCopy = $_.Name
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media").CopyHere("$desktopPath\$folderWithProcessedDocuments\Temporary\$currentImageNameCopy")
    Start-Sleep -Seconds 4
    }
Start-Sleep -Seconds 1
    #Deletes temporary *.txt file in the "media" folder
    (New-Object -COM Shell.Application).NameSpace("$desktopPath\$folderWithProcessedDocuments\Temporary zip").MoveHere("$desktopPath\$folderWithProcessedDocuments\$currentDirectoryName.zip\word\media\temporary.txt")
    }
