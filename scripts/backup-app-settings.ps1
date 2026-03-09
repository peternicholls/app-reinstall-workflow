[CmdletBinding()]
param(
    [string]$OutputRoot,

    [switch]$SkipColorProfiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-BackupBasePath {
    param(
        [AllowNull()]
        [string]$PreferredPath
    )

    if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
        return [System.IO.Path]::GetFullPath($PreferredPath)
    }

    $candidates = @(
        [Environment]::GetFolderPath('MyDocuments'),
        (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents'),
        $env:TEMP
    )

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    throw 'Unable to determine a valid settings backup folder.'
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    return $Path
}

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$documents = [Environment]::GetFolderPath('MyDocuments')
$localDocuments = Join-Path -Path $env:USERPROFILE -ChildPath 'Documents'
$backupRoot = Ensure-Directory -Path (Join-Path -Path (Resolve-BackupBasePath -PreferredPath $OutputRoot) -ChildPath "AppSettingsBackup_$timestamp")
$logFile = Join-Path -Path $backupRoot -ChildPath 'backup-log.txt'
$manifestPath = Join-Path -Path $backupRoot -ChildPath 'backup-manifest.json'

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestampedMessage = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $timestampedMessage | Tee-Object -FilePath $logFile -Append | Out-Null
}

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory)]
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    return $null
}

function Copy-FolderIfExists {
    param (
        [Parameter(Mandatory)]
        [string[]]$SourceCandidates,

        [Parameter(Mandatory)]
        [string]$DestinationName,

        [Parameter(Mandatory)]
        [string]$Description
    )

    $source = Resolve-ExistingPath -Candidates $SourceCandidates
    if ($null -eq $source) {
        Write-Log "WARNING: Not found for $Description"
        return [PSCustomObject]@{
            Description = $Description
            Source = $null
            Destination = $null
            Copied = $false
        }
    }

    $destination = Join-Path -Path $backupRoot -ChildPath $DestinationName
    Write-Log "Copying $Description: $source -> $destination"
    Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force

    return [PSCustomObject]@{
        Description = $Description
        Source = $source
        Destination = $destination
        Copied = $true
    }
}

Write-Log "Backup started"
Write-Log "Backup folder: $backupRoot"

$results = New-Object 'System.Collections.Generic.List[object]'
$results.Add((Copy-FolderIfExists -SourceCandidates @((Join-Path -Path $env:APPDATA -ChildPath 'DisplayCAL')) -DestinationName 'DisplayCAL' -Description 'DisplayCAL settings'))

if (-not $SkipColorProfiles.IsPresent) {
    $results.Add((Copy-FolderIfExists -SourceCandidates @('C:\Windows\System32\spool\drivers\color') -DestinationName 'ColorProfiles' -Description 'Windows color profiles'))
}

$results.Add((Copy-FolderIfExists -SourceCandidates @(
    (Join-Path -Path $documents -ChildPath 'SpectraCal\Calman'),
    (Join-Path -Path $localDocuments -ChildPath 'SpectraCal\Calman')
) -DestinationName 'Calman' -Description 'Calman settings'))

$manifest = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    backupRoot = $backupRoot
    items = @($results)
}

$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Log 'Backup completed'
Write-Host "Backup complete. Saved to: $backupRoot"