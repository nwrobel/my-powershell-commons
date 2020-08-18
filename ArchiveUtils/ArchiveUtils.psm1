# TDOD: add inline help docs to make these cmdlets usable by others on Github - make sure that
# "get-help" returns correct info for each cmdlet

function Archive-AllDirsInParentFolder {
    param (
        [string]$ParentFolder
    )

    $foldersToZip = (Get-ChildItem -Path $ParentFolder -Directory | Select-Object -ExpandProperty FullName)
    $numFoldersToZip = @($foldersToZip).Count

    Write-Host "$numFoldersToZip folders were found within the parent folder, ready to archive:" -ForegroundColor Yellow
    $foldersToZip

    do {
        $continue = Read-Host -Prompt "Proceed with archiving these folders? [Y/N]"

        if ($continue -eq 'n') {
            throw "Exiting due to choice of user"
        
        } elseif ($continue -ne 'y') {
            Write-Host "Incorrect response: Enter Y or N" -ForegroundColor Red
        }
    
    } while ($continue -ne 'y')

    Write-Host "`nStarting compression process on $numFoldersToZip folders..."

    foreach ($folder in $foldersToZip) {
        Write-Host "`n---------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Starting compression, directory: $folder" -ForegroundColor Cyan

        Compress-DirectoryToArchive -SourceDir $folder -DestinationDir $ParentFolder

        Write-Host "`nArchive creation complete for this directory!" -ForegroundColor Cyan
    }

    Write-Host "`nCompression process completed on $numFoldersToZip folders!" -ForegroundColor Green
}

# Compress the repo down into an archive using 7zip (high compression settings for best ratio)
# Type: 7z file
# Max compression (level 9), slowest but smallest file possible
# Preserve timestamps: modified, accessed, created
# Preserve file attributes
function Archive-Directory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourceDir,
        [Parameter(Mandatory=$true)]
        [string]$DestinationDir
    )
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

    Write-Verbose "$($thisFn): Found 7z program filepath: $($szProgramFilepath)"

    $folderName = (Split-Path $SourceDir -Leaf)
    $filenameTimestamp = (Get-FilenameCurrentTimestampTag)
    $targetArchiveFilename = "$($filenameTimestamp) $($folderName).archive.7z"
    $targetArchiveFilepath = (Join-Path $DestinationDir -ChildPath $targetArchiveFilename)

    Write-Verbose "$($thisFn): The following archive is ready to be created:`n  Source Dir: $($SourceDir)`n  Target Archive: $($targetArchiveFilepath)"

    $args = @(
        'a', 
        '-mx=9', 
        '-t7z', 
        '-mtm=on',
        '-mtc=on',
        '-mta=on',
        '-mtr=on',
        $targetArchiveFilepath,
        $SourceDir
    )
    $formattedArgs = $args | ForEach-Object { return "`"" + $_ + "`"" }

    Write-Verbose "$($thisFn): Launching compression operation with archive-level (highest) settings" 
    Start-Process -FilePath $szProgramFilepath -ArgumentList $formattedArgs -Wait -NoNewWindow

    Write-Verbose "$($thisFn): Archive of directory created successfully"
}

function Get-FilenameCurrentTimestampTag {
    $currentDt = Get-Date
    return "[$($currentDt.ToString('yyyy-MM-dd'))]"
}

function Test-AllArchivesInParentFolder {
    param (
        [string]$ParentFolder
    )

    $zipExecutablePath = "$env:ProgramFiles\7-Zip\7z.exe"
    $zipExecutablePath86 = "${env:ProgramFiles(x86)}\7-Zip\7z.exe"

    if (Test-Path -Path $zipExecutablePath) {
        Set-Alias -Name sz -Value $zipExecutablePath

    } elseif (Test-Path -Path $zipExecutablePath86) {
        Set-Alias -Name sz -Value $zipExecutablePath86

    } else {
        throw "7zip not found in either expected install dirs: $zipExecutablePath; $zipExecutablePath86"
    }

    $archivesToTest = (Get-ChildItem -Path $ParentFolder -Filter '*.7z' | Select-Object -ExpandProperty FullName)
    $numArchivesToTest = @($archivesToTest).Count

    Write-Host "$numArchivesToTest 7z archives were found within the parent folder, ready to test:" -ForegroundColor Yellow
    $archivesToTest

    do {
        $continue = Read-Host -Prompt "Proceed with testing these archives? [Y/N]"

        if ($continue -eq 'n') {
            throw "Exiting due to choice of user"
        
        } elseif ($continue -ne 'y') {
            Write-Host "Incorrect response: Enter Y or N" -ForegroundColor Red
        }
    
    } while ($continue -ne 'y')

    Write-Host "`nStarting testing process on $numArchivesToTest archives..."

    foreach ($archive in $archivesToTest) {
        Write-Host "`n---------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Starting test: $archive" -ForegroundColor Cyan

        sz t $archive * -r

        Write-Host "`nTest complete for this archive!" -ForegroundColor Cyan
    }

    Write-Host "`nTesting process completed on $numArchivesToTest archives!" -ForegroundColor Green
    
}

Export-ModuleMember -Function Archive-AllDirsInParentFolder, Archive-Directory, Test-AllArchivesInParentFolder






