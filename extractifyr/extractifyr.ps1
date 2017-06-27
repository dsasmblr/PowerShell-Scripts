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
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "C:\"
    $OpenFileDialog.showDialog() | Out-Null
    $OpenFileDialog.filename
}

# Allows user to choose folder to extract files to
Function Get-FolderName()
{
    Add-Type -Assembly "System.Windows.Forms"
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog() | Out-Null
    $FolderBrowser.SelectedPath
}

# Extracts files from selected file, to selected folder
Function Extract-Files()
{
    $File = Get-FileName #Assign path to file
    $Folder = Get-FolderName #Assign path to extraction folder
    $FileName = [System.IO.Path]::GetFileName($File) #Assign original filename for use and preservation
    
    If ($File -NotLike "*.zip") #Check for .zip extension
    {
        $FileNameWithZip = $FileName + ".zip" #Create filename with .zip extension
        Rename-Item -Path $File -NewName $FileNameWithZip #Rename the file in the folder
        $NewPath = $File.TrimEnd($FileName) #Create new path to renamed file for Expand-Archive
        $NewPath += $FileNameWithZip #New path creation continued
        Expand-Archive -LiteralPath $NewPath -DestinationPath $Folder -Force #Extract files (-Force overwrites)
        Rename-Item -Path $NewPath -NewName $FileName #Rename the file in the folder to its original name
    }
    Else #If file is already *.zip
    {
        Expand-Archive -LiteralPath $File -DestinationPath $Folder -Force #Extract files (-Force overwrites)
    }
}


#--------------------------#
# -- SCRIPT ENTRY POINT -- #
#--------------------------#


Script-Header
Extract-Files

Do {
    $UserYesNo = Read-Host "Extraction complete! Select another file? [Y] to run or any other key to exit"
    If ($UserYesNo -Eq "Y")
    {
        Clear-Host
        Script-Header
        Extract-Files
    }
    Else
    {
        Exit
    }
} While ($UserYesNo = "Y")