<#
    Script: extractifyr
    By: Stephen Chapman
    Site: http://dsasmblr.com/blog
    GitHub: http://github.com/dsasmblr
#>


#-----------------#
# -- FUNCTIONS -- #
#-----------------#


#Script header
Function Script-Header()
{
    Write-Host "#-----------------------#`r"
    Write-Host "  e-x-t-r-a-c-t-i-f-y-r `r"
    Write-Host "   By Stephen Chapman   `r"
    Write-Host "      dsasmblr.com      `r"
    Write-Host "#-----------------------#`n"
}

#Allows user to choose file to extract files from
Function Get-FileName()
{
    If ($StoredFilePath -Eq "") {
        #Tries to default to Steam directory
        $SteamPath = "C:\Program Files (x86)\Steam\steamapps\common"
        $SteamPathTest = Test-Path $SteamPath
        $FilePath = If ($SteamPathTest) {$SteamPath} Else {"C:\"}
    } Else {
        $FilePath = $StoredFilePath
    }

    #Dialog box to select file
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $FilePath
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName

    #Store file path
    $Global:StoredFilePath = $OpenFileDialog.FileName.TrimEnd([System.IO.Path]::GetFileName($OpenFileDialog.FileName))
}

#Allows user to choose folder to extract files to
Function Get-FolderName($Origin)
{
    If ($StoredDirPath -Eq ""){
        #Tries to default to Steam directory first, then desktop
        $SteamPath = "C:\Stephen\Steam\steamapps\common"
        $SteamPathTest = Test-Path $SteamPath
        $DesktopPath = [Environment]::GetFolderPath("Desktop")
        $DesktopPathTest = Test-Path $DesktopPath
        $DirPath = If ($SteamPathTest) {$SteamPath} ElseIf ($DesktopPathTest) {$DesktopPath} Else {"C:\"}
    } Else {
        $DirPath = $StoredDirPath
    }

    #Dialog box to select file
    Add-Type -Assembly "System.Windows.Forms"
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.SelectedPath = $DirPath
    If ($FolderBrowser.ShowDialog() -eq "Cancel") {
        User-Choice
    }
    $FolderBrowser.SelectedPath

    If ($Origin -Ne "ExtractTo") {
        $Global:StoredDirPath = $FolderBrowser.SelectedPath
    }
}

#Extracts files from selected file, to selected folder
Function Extract-Files($Origin, $FilesToExtract, $DirToExtractTo)
{
    If ($Origin -Eq "UserChoice") {
        # This block is ran if user specifies a particular file to extract
        Do {
            #Assign path to file
            $File = Get-FileName
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
            #Assign path to extraction folder
            $Folder = Get-FolderName
            If ($Folder -Eq "") {
                Write-Host "`nAre you sure you want to cancel?"
                $Answer = Read-Host "[Y] to cancel, [Enter] to choose a file"
                If ($Answer -Eq "Y") {
                    User-Choice
                }
            }
        } While ($Folder -Eq "")
        
        #Assign original filename for use and preservation
        $FileName = [System.IO.Path]::GetFileName($File)

        #Check for .zip extension and provision for Expand-Archive's file extension requirement
        If ($File -NotLike "*.zip") {
            #Create filename with .zip extension
            $FileNameWithZip = $FileName + ".zip"
            #Rename the file in the folder
            Rename-Item -Path $File -NewName $FileNameWithZip
            #Create new path to renamed file
            $NewPath = $File.TrimEnd($FileName) 
            $NewPath += $FileNameWithZip
            #Create directory based on file name
            $NewDir = New-Item -Path $Folder -Name $FileName -Type Directory
            #Extract file(s) from archive
            Expand-Archive -LiteralPath $NewPath -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue
            #Rename the file in the folder to its original name
            Rename-Item -Path $NewPath -NewName $FileName
        } Else {
            #Create directory based on file name
            $NewDir = New-Item -Path $Folder -Name $FileName -Type Directory
            #Extract file(s) from archive if extension is already .zip
            Expand-Archive -LiteralPath $File -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue
        }
    } Else {
        #This block is ran if user had a directory scanned
        ForEach ($FileName in $FilesToExtract) {
            
            #Assign original filename for use and preservation
            $OriginalFileName = [System.IO.Path]::GetFileName($FileName)

            # Check for .zip extension and provision for Expand-Archive's file extension requirement
            If ($FileName -NotLike "*.zip") {
                #Create filename with .zip extension
                $FileNameWithZip = $FileName + ".zip"
                #Rename the file in the folder
                Rename-Item -Path $FileName -NewName $FileNameWithZip
                #Create new path to renamed file for Expand-Archive
                $NewPath = $FileName.TrimEnd($FileName)
                $NewPath += $FileNameWithZip
                #Create directory based on file name
                $NewDir = New-Item -Path $DirToExtractTo -Name $OriginalFileName -Type Directory
                #Extract file(s) from archive
                Expand-Archive -LiteralPath $NewPath -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue

                Rename-Item -Path $NewPath -NewName $FileName # Rename the file in the folder to its original name
            } Else {
                #Create directory based on file name
                $NewDir = New-Item -Path $DirToExtractTo -Name $OriginalFileName -Type Directory
                #Extract file(s) from archive if extension is already .zip
                Expand-Archive -LiteralPath $FileName -DestinationPath $NewDir -Force -ErrorAction SilentlyContinue
            }
        }
    }
    #TO DO: Check for errors in extraction and let user know to try extracting to root.    
}

#Scans a folder and all sub-folders for files that can be extracted
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
  
    #Get directory and all children to scan, then scan them and place all file names and paths into an array
    Do {
        $FilesToScan = Get-FolderName | Get-ChildItem -Recurse -File | Select-Object -ExpandProperty FullName
        If ($FilesToScan.Count -Lt 1) {
            cls
            Script-Header
            Write-Host "You selected an empty folder. Please select a different folder."
        }
    } While ($FilesToScan.Count -Lt 1)

    #Magic number check to see if each file in the array is zip-compressed or not
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

    #Check for empty array and send user into a rescan flow accordingly
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

    #Ask user if they want to extract files or start over 
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

#Initial user choice to scan directory or extract file
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
    $UserYesNo = Read-Host "`nExtraction complete! Start over? [Y] to start over or any other key to exit"
    If ($UserYesNo -Eq "Y") {
        Clear-Host
        Script-Header
        User-Choice
    } Else {
        Exit
    }
} While ($UserYesNo = "Y")
