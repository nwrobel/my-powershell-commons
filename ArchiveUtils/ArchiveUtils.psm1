<#
.SYNOPSIS
Creates a compressed archive out of the given directory.

.DESCRIPTION
This cmdlet creates a compressed 7zip archive out of the given directory. The compression is done
with the highest setting (mx=9) for the smallest file possible. All files' timestamps (date modified,
date created, date accessed) and files' attributes will be preserved when added to the archive.  

.PARAMETER SourceDirectory
Path of the directory containing the source files to be archived.

.PARAMETER DestinationDirectory
Path of the directory to which the newly created archive file will be saved.

.PARAMETER ArchiveBaseFilename
The base filename that the archive should have (optional) instead of having the archive file be
named the name of the source directory (default).

.PARAMETER TimestampFilename
Whether or not to include the current timestamp (date/time of archive creation) in the archive
filename. Timestamp will be at the beginning of the filename and be formatted like:
[YYYY-MM-DD HH_mm_SS]

.PARAMETER Suffix
Suffix string to be added to the filename, just before the file extension. Useful for specifying the
type of archive (ex: old data archival, data backup, etc).
Example: myArchive.suffix.7z

.EXAMPLE
Compress-Directory -SourceDirectory "C:\Users\me\Documents" -DestinationDirectory "D:\Backups" -TimestampFilename -Suffix "backup"

#>
function Compress-Directory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory=$true)]
        [string]$DestinationDirectory,

        [Parameter(Mandatory=$false)]
        [string]$ArchiveBaseFilename,

        [Parameter(Mandatory=$false)]
        [switch]$TimestampFilename,

        [Parameter(Mandatory=$false)]
        [string]$Suffix = ''
    )
    $thisFn = $MyInvocation.MyCommand.Name

    $szProgramFilepath = Get-SevenZipProgramFilepath

    if ($ArchiveBaseFilename) {
        $baseFilename = $ArchiveBaseFilename
    
    } else {
        $srcFolderName = (Split-Path $SourceDirectory -Leaf)
        $baseFilename = $srcFolderName
    }

    if ($TimestampFilename) {
        $filenameTimestamp = (Get-FilenameCurrentTimestampTag)

        if ($Suffix) {
            $archiveFilename = "$($filenameTimestamp) $($baseFilename).$($Suffix).7z"
        } else {
            $archiveFilename = "$($filenameTimestamp) $($baseFilename).7z"
        }
    } else {
        if ($Suffix) {
            $archiveFilename = "$($baseFilename).$($Suffix).7z"
        } else {
            $archiveFilename = "$($baseFilename).7z"
        }
    }
    
    $targetArchiveFilepath = (Join-Path $DestinationDirectory -ChildPath $archiveFilename)
    Write-Verbose "$($thisFn): Ready to create archive:`n  Source Dir: $($SourceDirectory)`n  Target Archive File: $($targetArchiveFilepath)"

    $args = @(
        'a', 
        '-mx=9', 
        '-t7z', 
        '-mtm=on',
        '-mtc=on',
        '-mta=on',
        '-mtr=on',
        $targetArchiveFilepath,
        $SourceDirectory
    )
    $formattedArgs = $args | ForEach-Object { return "`"" + $_ + "`"" }

    Write-Verbose "$($thisFn): Launching compression operation with highest settings" 
    Start-Process -FilePath $szProgramFilepath -ArgumentList $formattedArgs -Wait -NoNewWindow

    Write-Verbose "$($thisFn): Archive of directory created successfully"
}

<#
.SYNOPSIS
Tests the given 7zip archive for integrity.

.DESCRIPTION
This cmdlet uses the 7zip application to test the given archive file for integrity. This is useful
to ensure there are no errors with the archive.

.PARAMETER ArchiveFilepath
Path to the 7zip archive file to be tested

#>
function Test-Archive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$ArchiveFilepath  
    )
    $thisFn = $MyInvocation.MyCommand.Name
    $szProgramFilepath = Get-SevenZipProgramFilepath

    Write-Verbose "$($thisFn): Starting test on archive: $ArchiveFilepath"

    $args = @(
        't',
        $ArchiveFilepath
    )
    $formattedArgs = $args | ForEach-Object { return "`"" + $_ + "`"" }
    Start-Process -FilePath $szProgramFilepath -ArgumentList $formattedArgs -Wait -NoNewWindow

    Write-Verbose "$($thisFn): Test complete for this archive"
}

<#
.SYNOPSIS
Creates a compressed archive out of each directory contained in the given root directory.

.DESCRIPTION
This cmdlet creates a 7zip archive out of each directory contained within in the given root 
directory. Only directories that are depth 1 within the root directory are made into archives: i.e.
the archives are created on the uppermost level of the root directory's subfolders.

.PARAMETER RootDirectory
The path of the root folder in which to search for directories in and archive them.

.PARAMETER TimestampFilename
Whether or not to include the current timestamp (date/time of archive creation) in the archive
filename. Timestamp will be at the beginning of the filename and be formatted like:
[YYYY-MM-DD HH_mm_SS]

.PARAMETER Suffix
Suffix string to be added to the filename, just before the file extension. Useful for specifying the
type of archive (ex: old data archival, data backup, etc).
Example: myArchive.suffix.7z

.PARAMETER Confirm
Whether or not to prompt the user to confirm the folders archive creation action.

.EXAMPLE
Compress-AllDirectoriesInRootDirectory -RootDirectory "C:\Program Files" -TimestampFilename -Suffix "backup" -Confirm

.NOTES
This cmdlet uses the "Compress-Directory" cmdlet from this module to compress and create the archive 
from each directory.

#>
function Compress-AllDirectoriesInRootDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$RootDirectory,

        [Parameter(Mandatory=$false)]
        [switch]$TimestampFilename,

        [Parameter(Mandatory=$false)]
        [string]$Suffix = '',    

        [Parameter(Mandatory=$false)]
        [switch]$Confirm
    )    
    $thisFn = $MyInvocation.MyCommand.Name

    $foldersToZip = (Get-ChildItem -Path $RootDirectory -Directory | Select-Object -ExpandProperty FullName)
    $numFoldersToZip = @($foldersToZip).Count

    Write-Host "The following $numFoldersToZip folders were found within the root folder and are ready to archive:" -ForegroundColor Yellow
    $foldersToZip

    if ($Confirm) {
        do {
            $continue = Read-Host -Prompt "Proceed with archiving these folders? [y/n]"
    
            if ($continue -eq 'n') {
                Write-Host "Exiting due to choice of user. No archives were created."
                return
            
            } elseif ($continue -ne 'y') {
                Write-Host "Incorrect response: Enter y or n" -ForegroundColor Red
            }
        } while ($continue -ne 'y')
    }

    Write-Host "`nStarting archival process on $numFoldersToZip folders" -ForegroundColor Yellow

    foreach ($folder in $foldersToZip) {
        Write-Host "`n---------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Compressing directory: $folder" -ForegroundColor Cyan

        $params = @{
            SourceDirectory = $folder;
            DestinationDirectory = $RootDirectory;
            TimestampFilename = $TimestampFilename
            Suffix = $Suffix
        }
        Compress-Directory @params

        Write-Host "`nArchive creation complete for this directory" -ForegroundColor Cyan
    }

    Write-Host "`nCompression process completed on $numFoldersToZip directories" -ForegroundColor Green
}

<#
.SYNOPSIS
Tests all of the 7zip archive files contained within the given root directory.

.DESCRIPTION
This cmdlet tests all of the 7zip archive files contained within the root directory for integrity.
The root directory is searched for 7zip files recursively, so all subfolders are searched.

.PARAMETER RootDirectory
Path of the directory to check all of the 7zip files contained within.

.PARAMETER Confirm
Whether or not to ask user confimation before starting the testing process.

#>
function Test-AllArchivesInRootDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$RootDirectory,

        [Parameter(Mandatory=$false)]
        [switch]$Confirm
    )
    $thisFn = $MyInvocation.MyCommand.Name

    $archivesToTest = Get-ChildItem -Path $RootDirectory -Filter '*.7z' -Recurse | Select-Object -ExpandProperty FullName
    $numArchivesToTest = @($archivesToTest).Count

    Write-Host "$numArchivesToTest 7z archives were found within the parent folder, ready to test:" -ForegroundColor Yellow
    $archivesToTest

    if ($Confirm) {
        do {
            $continue = Read-Host -Prompt "Proceed with testing these archives? [y/n]"
    
            if ($continue -eq 'n') {
                Write-Host "Exiting due to choice of user. No archives were tested."
                return
            
            } elseif ($continue -ne 'y') {
                Write-Host "Incorrect response: Enter y or n" -ForegroundColor Red
            }
        
        } while ($continue -ne 'y')
    }

    Write-Host "`nStarting testing process on $numArchivesToTest archives"

    foreach ($archive in $archivesToTest) {
        Write-Host "`n---------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Test-Archive -ArchiveFilepath $archive
    }

    Write-Host "`nTesting process completed on $numArchivesToTest archives!" -ForegroundColor Green
}

<#
.SYNOPSIS
Returns the filepath of the 7zip program installed on this system.

.DESCRIPTION
This cmdlet returns the filepath of the 7zip program installed on this system. It looks in both the
"Program Files" and "Program Files (x86)" folders, which 7zip is installed to by default. If 7z.exe 
is not found in any of these locations, an exception is thrown.

#>
function Get-SevenZipProgramFilepath {
    [CmdletBinding()]
    param ()
    $thisFn = $MyInvocation.MyCommand.Name
    
    $szDefaultFilepath = "$env:ProgramFiles\7-Zip\7z.exe"
    $szDefaultFilepath86 = "${env:ProgramFiles(x86)}\7-Zip\7z.exe"

    if (Test-Path -Path $szDefaultFilepath) {
        $szProgramFilepath = $szDefaultFilepath

    } elseif (Test-Path -Path $szDefaultFilepath86) {
        $szProgramFilepath = $szDefaultFilepath86

    } else {
        throw "7zip not found in either expected install dirs: $zipExecutablePath; $zipExecutablePath86"
    }

    Write-Verbose "$($thisFn): Found 7z program filepath at $($szProgramFilepath)"
    return $szProgramFilepath
}

<#
.SYNOPSIS
Returns the current timestamp, formatted for use as a filename "tag". Returns time in the format:
"[YYYY-MM-DD HH_mm_SS]""

#>
function Get-FilenameCurrentTimestampTag {
    $currentDt = Get-Date
    return "[$($currentDt.ToString('yyyy-MM-dd HH_mm_ss'))]"
}


Export-ModuleMember -Function Compress-Directory, Test-Archive, Compress-AllDirectoriesInRootDirectory, Test-AllArchivesInRootDirectory






