$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$documents = [Environment]::GetFolderPath("MyDocuments")
$localDocuments = Join-Path $env:USERPROFILE "Documents"
$backupRoot = Join-Path $documents "AppSettingsBackup_$timestamp"
$logFile = Join-Path $backupRoot "backup-log.txt"

New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null

$displayCalPath = Join-Path $env:APPDATA "DisplayCAL"
$colorProfilesPath = "C:\Windows\System32\spool\drivers\color"

$possibleCalmanPaths = @(
    (Join-Path $documents "SpectraCal\Calman"),
    (Join-Path $localDocuments "SpectraCal\Calman")
)

function Log {
    param([string]$Message)
    $Message | Tee-Object -FilePath $logFile -Append
}

function Copy-FolderIfExists {
    param (
        [string]$Source,
        [string]$DestinationName
    )

    if (Test-Path $Source) {
        $destination = Join-Path $backupRoot $DestinationName
        Log "Copying: $Source -> $destination"
        Copy-Item -Path $Source -Destination $destination -Recurse -Force
        return $true
    } else {
        Log "WARNING: Not found: $Source"
        return $false
    }
}

Log "Backup started: $(Get-Date)"
Log "Backup folder: $backupRoot"

Copy-FolderIfExists -Source $displayCalPath -DestinationName "DisplayCAL" | Out-Null
Copy-FolderIfExists -Source $colorProfilesPath -DestinationName "ColorProfiles" | Out-Null

$calmanCopied = $false
foreach ($path in $possibleCalmanPaths) {
    if (Copy-FolderIfExists -Source $path -DestinationName "Calman") {
        $calmanCopied = $true
        break
    }
}

if (-not $calmanCopied) {
    Log "WARNING: No Calman folder found in expected locations."
}

Log "Backup completed: $(Get-Date)"
Write-Host "Backup complete. Saved to: $backupRoot"