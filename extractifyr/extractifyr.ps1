<#
    Script: extractifyr
    By: Stephen Chapman
    Site: http://dsasmblr.com/blog
    GitHub: http://github.com/dsasmblr
#>


#-----------------#
# -- FUNCTIONS -- #
#-----------------#


# Script header
Function Script-Header()
{
    Write-Host "#-----------------------#`r"
    Write-Host "  e-x-t-r-a-c-t-i-f-y-r `r"
    Write-Host "   By Stephen Chapman   `r"
    Write-Host "      dsasmblr.com      `r"
    Write-Host "#-----------------------#`n"
}

# Allows user to choose file to extract files from
Function Get-FileName()
{
    If ($StoredFilePath -Eq "") {
        $SteamPath = "C:\Program Files (x86)\Steam\steamapps\common"
        $SteamPathTest = Test-Path $SteamPath
        $FilePath = If ($SteamPathTest) {$SteamPath} Else {"C:\"}
    } Else {
        $FilePath = $StoredFilePath
    }

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $FilePath
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName

    $Global:StoredFilePath = $OpenFileDialog.FileName.TrimEnd([System.IO.Path]::GetFileName($OpenFileDialog.FileName))
}

# Allows user to choose folder to extract files to
Function Get-FolderName($Origin)
{
    If ($StoredDirPath -Eq ""){
        $SteamPath = "C:\Stephen\Steam\steamapps\common"
        $SteamPathTest = Test-Path $SteamPath
        $DesktopPath = [Environment]::GetFolderPath("Desktop")
        $DesktopPathTest = Test-Path $DesktopPath
        $DirPath = If ($SteamPathTest) {$SteamPath} ElseIf ($DesktopPathTest) {$DesktopPath} Else {"C:\"}
    } Else {
        $DirPath = $StoredDirPath
    }

    Add-Type -Assembly "System.Windows.Forms"
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.SelectedPath = $DirPath
    #$FolderBrowser.ShowDialog() | Out-Null
    If ($FolderBrowser.ShowDialog() -eq "Cancel") {
        User-Choice
    }
    $FolderBrowser.SelectedPath

    If ($Origin -Ne "ExtractTo") {
        $Global:StoredDirPath = $FolderBrowser.SelectedPath
    }
}

# Extracts files from selected file, to selected folder
Function Extract-Files($Origin, $FilesToExtract, $DirToExtractTo)
{
    If ($Origin -Eq "UserChoice") {

        # This block is ran if user specifies a particular file to extract

        Do {
            $File = Get-FileName # Assign path to file
            If ($File -Eq "") {
                Write-Host "`nAre you sure you want to cancel?"
                $Answer = Read-Host "[Y] to cancel, [Enter] to choose a file"
                If ($Answer -Eq "Y") {
                    User-Choice
                }
            }
        } While ($File -Eq "")

        Read-Host "`nFile selected! Now press [Enter] to choose a directory to extract to"
        
        Do {
            $Folder = Get-FolderName # Assign path to extraction folder
            If ($Folder -Eq "") {
                Write-Host "`nAre you sure you want to cancel?"
                $Answer = Read-Host "[Y] to cancel, [Enter] to choose a file"
                If ($Answer -Eq "Y") {
                    User-Choice
                }
            }
        } While ($Folder -Eq "")
        
        $FileName = [System.IO.Path]::GetFileName($File) # Assign original filename for use and preservation

        # Check for .zip extension and provision for Expand-Archive's file extension requirement
        If ($File -NotLike "*.zip") {
            $FileNameWithZip = $FileName + ".zip" # Create filename with .zip extension
            Rename-Item -Path $File -NewName $FileNameWithZip # Rename the file in the folder
            $NewPath = $File.TrimEnd($FileName) # Create new path to renamed file (1)
            $NewPath += $FileNameWithZip # Create new path to renamed file (2)
            $NewDir = New-Item -Path $Folder -Name $FileName -Type Directory
            Expand-Archive -LiteralPath $NewPath -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue # Extract files
            Rename-Item -Path $NewPath -NewName $FileName # Rename the file in the folder to its original name
        } Else {
            $NewDir = New-Item -Path $Folder -Name $FileName -Type Directory
            Expand-Archive -LiteralPath $File -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue # Extract files if extension is already .zip
        }
    } Else {

        # This block is ran if user had a directory scanned

        ForEach ($FileName in $FilesToExtract) {
            $OriginalFileName = [System.IO.Path]::GetFileName($FileName)

            # Check for .zip extension and provision for Expand-Archive's file extension requirement
            If ($FileName -NotLike "*.zip") {
                $FileNameWithZip = $FileName + ".zip" # Create filename with .zip extension
                Rename-Item -Path $FileName -NewName $FileNameWithZip # Rename the file in the folder
                $NewPath = $FileName.TrimEnd($FileName) # Create new path to renamed file for Expand-Archive
                $NewPath += $FileNameWithZip # New path creation continued
                $NewDir = New-Item -Path $DirToExtractTo -Name $OriginalFileName -Type Directory
                Expand-Archive -LiteralPath $NewPath -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue # Extract files (-Force overwrites)
                Rename-Item -Path $NewPath -NewName $FileName # Rename the file in the folder to its original name
            } Else {
                $NewDir = New-Item -Path $DirToExtractTo -Name $OriginalFileName -Type Directory
                Expand-Archive -LiteralPath $FileName -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue # Extract files if extension is already .zip
            }
        }
    }
    #Check for errors in extraction and let user know to try extracting to root.    
}

# Scans a folder and all sub-folders for files that can be extracted
Function Scan-Files()
{
    cls
    Script-Header
    Write-Host "/////////////////////////////////////////////////////////////////////"
    Write-Host "--------------------------------------------------------------------"
    Write-Host " Preparing for file scan.`r"
    Write-Host " Files identified as matches will pop up below as the scan occurs.`r"
    Write-Host " This can take a while if scanning a lot of files, so be patient! =)`r"
    Write-Host " You will be presented with options after the scan is complete.`r"
    Write-Host "--------------------------------------------------------------------"
    Write-Host "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\`n"
  
    Do {
        $FilesToScan = Get-FolderName | Get-ChildItem -Recurse -File | Select-Object -ExpandProperty FullName
        If ($FilesToScan.Count -Lt 1) {
            cls
            Script-Header
            Write-Host "You selected an empty folder. Please select a different folder."
            #$Answer = Read-Host "[Y] to cancel, [Enter] to choose a folder"
            #If ($Answer -Eq "Y") {
            #    User-Choice
            #}
        }
    } While ($FilesToScan.Count -Lt 1)

    $FilesToScan = ForEach ($FileName in $FilesToScan) {
        Try {
            $i = [System.BitConverter]::ToString((Get-Content $FileName -ReadCount 1 -TotalCount 4 -Encoding Byte))
        } Catch {
            # Ignore errors related to empty files
        }
        Try {
            If ($i.startsWith("50-4B")) {
                $Discovered = [System.IO.Path]::GetFileName($FileName)
                Write-Host "File discovered: $Discovered"
                $Global:FileArray += ,$FileName
            }
        } Catch {
            # Ignore errors related to non-zip files
        }
    }

    #Check for empty array
    If ($FileArray.Count -Lt 1) {

        cls
        Script-Header
        Write-Host "No zip files found! Would you like to try again or scan another folder?"
        
        Do {
            $UserSelection = Read-Host "[Y] to choose a directory or [N] to start over"

            If ($UserSelection -Eq "Y") {
                Scan-Files
            } ElseIf ($UserSelection -Eq "N") {
                User-Choice
            } Else {
                Write-Host "`nWat? Please choose one of the following:"
            }
        } While ($UserSelection -Ne "Y" -And $UserSelection -Ne "N") 
    }

    Write-Host "`nScan completed!"
    Write-Host "`nPress [Enter] to select a folder to extract files to or [S] to start over"
    $UserSelection = Read-Host "(Tip: Press [S] to start over if you want to extract individual files from the results)"

    If ($UserSelection -Eq "S") {
        User-Choice    
    } Else {
        $Global:ExtractTo = Get-FolderName "ExtractTo"
        Extract-Files "FromScanFiles" $FileArray $ExtractTo
        $Global:FileArray = @()
    }
}

# Choose option
Function User-Choice()
{
    cls
    Script-Header
    Write-Host "Would you like to scan a directory for files to extract, or select a file to extract?"

    Do {
        $UserSelection = Read-Host "[S] to choose a directory to scan or [E] to choose a file to extract"

        If ($UserSelection -Eq "S") {
            Scan-Files
        } ElseIf ($UserSelection -Eq "E") {
            Extract-Files "UserChoice"
        } Else {
            Write-Host "`nWat? Please choose one of the following:"
        }
    } While ($UserSelection -Ne "S" -And $UserSelection -Ne "E")

}


#--------------------------#
# -- SCRIPT ENTRY POINT -- #
#--------------------------#


$Global:StoredFilePath = ""
$Global:StoredDirPath = ""
$Global:FileArray = @()
$Global:ExtractTo = ""
User-Choice

Do {
    #cls
    #Script-Header
    $UserYesNo = Read-Host "`nExtraction complete! Start over? [Y] to start over or any other key to exit"
    If ($UserYesNo -Eq "Y") {
        Clear-Host
        Script-Header
        User-Choice
    } Else {
        Exit
    }
} While ($UserYesNo = "Y")
