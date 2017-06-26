# -- FUNCTIONS -- #

# Script header
Function Script-Header()
{
    Write-Host "#------------------#`r"
    Write-Host "   pngifierScript   `r"
    Write-Host " By Stephen Chapman`r"
    Write-Host "    dsasmblr.com   `r"
    Write-Host "#------------------#`n"
}

# File selector dialog box
Function Get-FileName($DirToPath)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $DirToPath
    $OpenFileDialog.showDialog() | Out-Null
    $OpenFileDialog.filename
}

# -- SCRIPT ENTRY POINT -- #

# Global variable
$InputFile = ""
$StoredPath = ""

# Recursive anonymous function to repeat script
$RunScript = {
    Script-Header
    Read-Host "Press enter to select a file for pngification"

    # Check if path has already been specified; if not, check if Steam path exists, else default to C:\
    If (!$StoredPath)
    {
        $SteamPath = Test-Path "C:\Program Files (x86)\Steam\steamapps\common"
        $PathTest = Test-Path $SteamPath
        $InputFile = If ($PathTest) {Get-FileName $SteamPath} Else {Get-FileName "C:\"}
    }
    Else
    {
        $InputFile = Get-FileName $StoredPath
    }

    # Prompt user for file again if they hit cancel
    While (!$InputFile) {
        Read-Host "Why you no choose file? Press enter to choose a file"
        $InputFile = Get-FileName $InputFile
    }

    # If directory doesn't exist, create using filename of specified file
    # If directory does exist, ask if overwrite, create another dir, or start over
    $global:DirName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)

    If (Test-Path $DirName)
    {
        Do
        {
            Write-Host "This directory already exists. What would you like to do?`r"
            Write-Host "[O] Overwrite files [C] Create another directory [S] Start over`n"
            $AlreadyExists = Read-Host "Choose one of the options above"

            Switch ($AlreadyExists)
            {
                "O" {
                        Break
                    }
                "C" {
                        $dnCount = 0
                        $CreateAnotherDir = {
                            $dnCount += 1
                            $DirName += $dnCount.ToString()
                            If (Test-Path $DirName)
                            {
                                $DirName = $DirName.TrimEnd($dnCount.ToString())
                                .$CreateAnotherDir
                            }
                            Else
                            {
                                New-Item $DirName -Type directory | Out-Null
                                $global:DirName = $DirName
                            }
                        }
                        &$CreateAnotherDir
                    }
                "S" {
                        Clear-Host
                        $StoredPath = $InputFile.TrimEnd([System.IO.Path]::GetFileName($InputFile))
                        .$RunScript
                    }
                Default {Write-Host "Wat? O, C, or S, please...`n"}
            }            
        } While ($AlreadyExists -Ne "O" -And $AlreadyExists -Ne "C" -And $AlreadyExists -Ne "S")
    }
    Else
    {
        New-Item $DirName -Type directory | Out-Null
    }

    # Scrape file for PNG data, then convert result to an array 
    $PngString = Select-String -Path $InputFile -Pattern "url(data:image/png;base64" -SimpleMatch
    $PngString = $PngString -Split "\s"

    # Loop through each item in the array and export PNG files
    $Count = 0
    ForEach ($i in $PngString){
        $i = $i.Trim()
        $i = $i.TrimStart("url(data:image/png;base64")
        $i = $i -Replace ",", ""
        $i = $i -Replace ";", ""
        $i = $i.TrimEnd(")*")
        If ($i.StartsWith("iVBOR")){
            $Count += 1
            $i = [Convert]::FromBase64CharArray($i, 0, $i.Length)
            [io.file]::WriteAllBytes("$DirName/$Count.png", $i)
        }
    }

    Do
    {
        II $DirName # Open directory files were saved to
        $RunAgainOrExit = Read-Host "`nScript completed! Run again? [Y] to run again or [N] to exit"

        If ($RunAgainOrExit -Eq "Y")
        {
            cls
            $StoredPath = $InputFile.TrimEnd([System.IO.Path]::GetFileName($InputFile))
            .$RunScript
        }

        If ($RunAgainOrExit -Eq "N")
        {
            Return
        }
    } While ($RunAgainOrExit -Ne "Y" -And $RunAgainOrExit -Ne "N")
}

&$RunScript