param(
    [string]$ExcelFile = ".\PART_MODEL_with_results_PDF.xlsx",
    [string]$PathColumn = "Datasheet Link",
    [string]$FolderColumn = "A",
    [string]$StatusColumn = "",
    [string]$RequiredStatus = "",
    [string]$OutDir = ".\Downloads from DSL",
    [string]$LogFile = ".\download_log_legacy.csv",
    [ValidateSet("Skip", "Rename", "Overwrite")]
    [string]$IfExists = "Rename",
    [string]$DatasheetsLibraryRoot = "C:\Users\H588238\OneDrive - Honeywell\TSI Team Site (GPS Sofia) - General\3 DATASHEETS\Datasheets Library",
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    [ValidateRange(1, 60)]
    [int]$RetryDelaySeconds = 2,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:logInitialized = $false
$script:resolvedLocalPathCache = @{}

function Ensure-ModuleInstalled {
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "$Name not found. Installing for current user..."
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }

    try {
        Import-Module $Name -Force -DisableNameChecking
    }
    catch {
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
        Import-Module $Name -Force -DisableNameChecking
    }
}

function Get-UniqueTargetPath {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][string]$FileName
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $candidate = Join-Path $Directory $FileName
    $index = 2

    while (Test-Path $candidate) {
        $candidate = Join-Path $Directory ("{0}_{1}{2}" -f $baseName, $index, $extension)
        $index++
    }

    return $candidate
}

function Write-LogRow {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $true)][string]$Path
    )

    if (-not $script:logInitialized) {
        $Row | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        $script:logInitialized = $true
        return
    }

    $Row | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
}

function Resolve-LocalSourcePath {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$LibraryRoot
    )

    if ($script:resolvedLocalPathCache.ContainsKey($InputPath)) {
        return $script:resolvedLocalPathCache[$InputPath]
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add($InputPath)

    try {
        $decodedInput = [System.Uri]::UnescapeDataString($InputPath)
        if ($decodedInput -ne $InputPath) {
            $candidates.Add($decodedInput)
        }
    }
    catch {
        # Keep raw candidate only.
    }

    if ($LibraryRoot -and (Test-Path $LibraryRoot)) {
        $fileName = [System.IO.Path]::GetFileName($InputPath)
        if ($fileName) {
            $decodedFileName = $fileName
            try {
                $decodedFileName = [System.Uri]::UnescapeDataString($fileName)
            }
            catch {
                # Keep original filename.
            }

            foreach ($nameCandidate in @($fileName, $decodedFileName) | Select-Object -Unique) {
                if (-not $nameCandidate) {
                    continue
                }

                $directCandidate = Join-Path $LibraryRoot $nameCandidate
                $candidates.Add($directCandidate)
            }
        }
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if ($candidate -and (Test-Path $candidate)) {
            $script:resolvedLocalPathCache[$InputPath] = $candidate
            return $candidate
        }
    }

    if ($LibraryRoot -and (Test-Path $LibraryRoot)) {
        $fileName = [System.IO.Path]::GetFileName($InputPath)
        if ($fileName) {
            $decodedFileName = $fileName
            try {
                $decodedFileName = [System.Uri]::UnescapeDataString($fileName)
            }
            catch {
                # Keep original filename.
            }

            foreach ($nameCandidate in @($fileName, $decodedFileName) | Select-Object -Unique) {
                if (-not $nameCandidate) {
                    continue
                }

                $found = Get-ChildItem -Path $LibraryRoot -Filter $nameCandidate -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) {
                    $script:resolvedLocalPathCache[$InputPath] = $found.FullName
                    return $found.FullName
                }
            }
        }
    }

    $script:resolvedLocalPathCache[$InputPath] = $null
    return $null
}

function Resolve-UrlColumn {
    param(
        [string[]]$AvailableColumns,
        [string]$PreferredColumn
    )

    if ($AvailableColumns -contains $PreferredColumn) {
        return $PreferredColumn
    }

    foreach ($candidate in @("Datasheet Link", "MappedUrl", "Result", "Column3", "Url", "URL", "Link")) {
        if ($AvailableColumns -contains $candidate) {
            return $candidate
        }
    }

    throw "URL column '$PreferredColumn' not found. Available columns: $($AvailableColumns -join ', ')"
}

function Resolve-FolderColumn {
    param(
        [string[]]$AvailableColumns,
        [string]$PreferredColumn
    )

    # "A" means use the first column in the sheet.
    if ($PreferredColumn -eq "A") {
        return $AvailableColumns[0]
    }

    if ($AvailableColumns -contains $PreferredColumn) {
        return $PreferredColumn
    }

    return $AvailableColumns[0]
}

function Sanitize-RelativeFolder {
    param([string]$Value)

    if (-not $Value) {
        return ""
    }

    $folder = $Value.Trim()
    if (-not $folder) {
        return ""
    }

    $folder = [Regex]::Replace($folder, '[<>:"/\\|?*\x00-\x1F]', '_')

    $folder = $folder -replace '^[\\/]+', ''
    $folder = $folder -replace '\\.\\.', '_'
    return $folder
}

function Resolve-SourceFromCell {
    param([string]$Value)

    if (-not $Value) {
        return $null
    }

    $text = $Value.Trim()
    if (-not $text) {
        return $null
    }

    if ($text -match '^(?:[A-Za-z]:\\|\\\\)') {
        return [PSCustomObject]@{
            SourceType = "LOCAL"
            SourceValue = $text
        }
    }

    return $null
}

function Get-DownloadItemsFromExcel {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$UrlColumn,
        [string]$FolderCol,
        [string]$FilterStatusColumn,
        [string]$FilterStatusValue
    )

    if (-not (Test-Path $Path)) {
        return @()
    }

    Ensure-ModuleInstalled -Name "ImportExcel"
    $rows = Import-Excel -Path $Path
    if (-not $rows -or $rows.Count -eq 0) {
        return @()
    }

    $firstRow = $rows | Select-Object -First 1
    $availableColumns = @($firstRow.PSObject.Properties.Name)
    if ($availableColumns.Count -eq 0) {
        return @()
    }

    $resolvedUrlColumn = Resolve-UrlColumn -AvailableColumns $availableColumns -PreferredColumn $UrlColumn
    $resolvedFolderColumn = Resolve-FolderColumn -AvailableColumns $availableColumns -PreferredColumn $FolderCol

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        if ($FilterStatusColumn -and $FilterStatusValue -and ($row.PSObject.Properties.Name -contains $FilterStatusColumn)) {
            $statusValue = [string]$row.$FilterStatusColumn
            if ($statusValue.Trim().ToUpperInvariant() -ne $FilterStatusValue.ToUpperInvariant()) {
                continue
            }
        }

        $sourceRaw = [string]$row.$resolvedUrlColumn
        $resolvedSource = Resolve-SourceFromCell -Value $sourceRaw
        if ($resolvedSource) {
            $relativeFolderRaw = [string]$row.$resolvedFolderColumn
            $relativeFolder = Sanitize-RelativeFolder -Value $relativeFolderRaw
            $items.Add([PSCustomObject]@{
            SourceType = $resolvedSource.SourceType
            SourceValue = $resolvedSource.SourceValue
                    RelativeFolder = $relativeFolder
                }) | Out-Null
        }
    }

    return $items.ToArray()
}

if (-not (Test-Path $DatasheetsLibraryRoot)) {
    throw "Datasheets library root not found: $DatasheetsLibraryRoot"
}

$inputItems = @(Get-DownloadItemsFromExcel -Path $ExcelFile -UrlColumn $PathColumn -FolderCol $FolderColumn -FilterStatusColumn $StatusColumn -FilterStatusValue $RequiredStatus)
if (@($inputItems).Count -eq 0) {
    throw "No local datasheet source paths found in Excel file '$ExcelFile'."
}

Write-Host ("Using Excel file: {0}" -f (Resolve-Path $ExcelFile))

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$logParent = Split-Path -Parent $LogFile
if ($logParent) {
    New-Item -ItemType Directory -Path $logParent -Force | Out-Null
}
if (Test-Path $LogFile) {
    Remove-Item -Path $LogFile -Force
}

$parsedRows = @($inputItems | ForEach-Object {
        $targetDirectory = $OutDir
        if ($_.RelativeFolder) {
            $targetDirectory = Join-Path $OutDir $_.RelativeFolder
        }

        $localSourcePath = $_.SourceValue
        [PSCustomObject]@{
            SourceType = "LOCAL"
            OriginalUrl = ""
            Uri         = $null
            HostSite    = "__LOCAL__"
            ServerPath  = ""
            LocalSourcePath = $localSourcePath
            FileName    = [System.IO.Path]::GetFileName($localSourcePath)
            TargetDir   = $targetDirectory
            RelativeFolder = $_.RelativeFolder
        }
    })

if ($DryRun) {
    Write-Host ("DryRun: parsed {0} input item(s)." -f $parsedRows.Count)
    $localCount = @($parsedRows | Where-Object { $_.SourceType -eq "LOCAL" }).Count
    Write-Host ("- Local-path items: {0}" -f $localCount)
    $parsedRows | Group-Object HostSite | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("- {0} => {1} file(s)" -f $_.Name, $_.Count)
    }
    return
}

$success = 0
$failure = 0
$skipped = 0

$parsedRows | Group-Object HostSite | ForEach-Object {
    $siteUrl = $_.Name

    foreach ($item in $_.Group) {
        try {
            New-Item -ItemType Directory -Path $item.TargetDir -Force | Out-Null

            $baseTargetPath = Join-Path $item.TargetDir $item.FileName
            $targetPath = $baseTargetPath
            if (Test-Path $baseTargetPath) {
                if ($IfExists -eq "Skip") {
                    $skipped++
                    Write-LogRow -Path $LogFile -Row ([PSCustomObject]@{
                            SourceType = $item.SourceType
                            Url        = $item.OriginalUrl
                            SourcePath = $item.LocalSourcePath
                            Site       = $siteUrl
                            Status     = "SKIP"
                            Folder     = $item.RelativeFolder
                            SavedAs    = $item.FileName
                            Message    = "Destination file already exists"
                        })
                    Write-Host ("SKIP {0}" -f $item.FileName)
                    continue
                }

                if ($IfExists -eq "Rename") {
                    $targetPath = Get-UniqueTargetPath -Directory $item.TargetDir -FileName $item.FileName
                }
            }

            $targetName = [System.IO.Path]::GetFileName($targetPath)

            if ($item.SourceType -eq "LOCAL") {
                $effectiveSourcePath = Resolve-LocalSourcePath -InputPath $item.LocalSourcePath -LibraryRoot $DatasheetsLibraryRoot
                if ($effectiveSourcePath -and (Test-Path $effectiveSourcePath)) {
                    Copy-Item -Path $effectiveSourcePath -Destination $targetPath -Force
                }
                else {
                    throw "Local source file not found: $($item.LocalSourcePath)"
                }
            }
            else {
                throw "Unsupported source type in local-only mode: $($item.SourceType)"
            }

            $success++
            Write-LogRow -Path $LogFile -Row ([PSCustomObject]@{
                    SourceType = $item.SourceType
                    Url        = $item.OriginalUrl
                    SourcePath = $item.LocalSourcePath
                    Site       = $siteUrl
                    Status     = "OK"
                    Folder     = $item.RelativeFolder
                    SavedAs    = $targetName
                    Message    = ""
                })
            Write-Host ("OK   {0}" -f $targetName)
        }
        catch {
            $failure++
            Write-LogRow -Path $LogFile -Row ([PSCustomObject]@{
                    SourceType = $item.SourceType
                    Url        = $item.OriginalUrl
                    SourcePath = $item.LocalSourcePath
                    Site       = $siteUrl
                    Status     = "FAIL"
                    Folder     = $item.RelativeFolder
                    SavedAs    = ""
                    Message    = $_.Exception.Message
                })
            Write-Host ("FAIL {0}" -f $item.FileName)
        }
    }
}

Write-Host ""
Write-Host ("Completed. Success: {0}  Failed: {1}  Skipped: {2}" -f $success, $failure, $skipped)
Write-Host ("Log: {0}" -f (Resolve-Path $LogFile))
Write-Host ("Output folder: {0}" -f (Resolve-Path $OutDir))
