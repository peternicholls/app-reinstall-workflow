[CmdletBinding()]
param(
    [string]$OutputRoot,

    [switch]$SkipWingetExport,

    [switch]$SkipStoreApps,

    [switch]$SkipFolderListings
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

function Resolve-BackupBasePath {
    param(
        [AllowNull()]
        [string]$PreferredPath
    )

    if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
        return Resolve-AppPath -Path $PreferredPath
    }

    $candidates = @(
        (Join-Path -Path $env:USERPROFILE -ChildPath 'OneDrive\Documents'),
        [Environment]::GetFolderPath('MyDocuments'),
        $env:TEMP
    )

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    throw 'Unable to determine a valid output folder for the backup pack.'
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

function Write-Section {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$FilePath
    )

    Add-Content -LiteralPath $FilePath -Value ''
    Add-Content -LiteralPath $FilePath -Value ('=' * 80)
    Add-Content -LiteralPath $FilePath -Value $Title
    Add-Content -LiteralPath $FilePath -Value ('=' * 80)
}

function Export-InstalledPrograms {
    param(
        [Parameter(Mandatory)]
        [string]$CatalogCsvPath,

        [Parameter(Mandatory)]
        [string]$DetailedCsvPath,

        [Parameter(Mandatory)]
        [string]$TextPath
    )

    $programs = @(Get-InstalledProgramInventory)
    $catalogRows = @(
        $programs |
            Select-Object DisplayName, DisplayVersion, Publisher
    )

    $catalogRows | Export-Csv -LiteralPath $CatalogCsvPath -NoTypeInformation -Encoding UTF8
    $programs | Export-Csv -LiteralPath $DetailedCsvPath -NoTypeInformation -Encoding UTF8
    $programs |
        Format-Table DisplayName, DisplayVersion, Publisher, Scope, Architecture -AutoSize |
        Out-String -Width 4096 |
        Set-Content -LiteralPath $TextPath -Encoding UTF8

    return [PSCustomObject]@{
        ProgramCount = $programs.Count
        Programs = $programs
    }
}

function Export-WingetPackageList {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Exported = $false
            Details = 'winget.exe not found'
        }
    }

    $output = @(
        & winget.exe export --output $OutputPath --include-versions --accept-source-agreements 2>&1
    )
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        return [PSCustomObject]@{
            Exported = $false
            Details = ([string]::Join(' ', ($output | Select-Object -Last 3))).Trim()
        }
    }

    return [PSCustomObject]@{
        Exported = $true
        Details = 'winget export completed'
    }
}

function Export-StartupItems {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $items = New-Object 'System.Collections.Generic.List[object]'
    $registryPaths = @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
    )

    foreach ($path in $registryPaths) {
        if (-not (Test-Path -LiteralPath $path)) {
            continue
        }

        $item = Get-ItemProperty -LiteralPath $path
        foreach ($property in $item.PSObject.Properties) {
            if ($property.Name -match '^PS') {
                continue
            }

            $items.Add([PSCustomObject]@{
                Source = $path
                Name = [string]$property.Name
                Value = [string]$property.Value
            })
        }
    }

    $startupFolders = @(
        (Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\Windows\Start Menu\Programs\Startup'),
        (Join-Path -Path $env:ProgramData -ChildPath 'Microsoft\Windows\Start Menu\Programs\StartUp')
    )

    foreach ($folder in $startupFolders) {
        if (-not (Test-Path -LiteralPath $folder)) {
            continue
        }

        foreach ($file in @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction SilentlyContinue)) {
            $items.Add([PSCustomObject]@{
                Source = $folder
                Name = [string]$file.Name
                Value = [string]$file.FullName
            })
        }
    }

    @($items) | Sort-Object Source, Name | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

function Export-ShortcutInventory {
    param(
        [Parameter(Mandatory)]
        [string[]]$Folders,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $results = New-Object 'System.Collections.Generic.List[object]'
    $shell = New-Object -ComObject WScript.Shell

    try {
        foreach ($folder in $Folders) {
            if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
                continue
            }

            foreach ($shortcutFile in @(Get-ChildItem -LiteralPath $folder -Recurse -Filter '*.lnk' -Force -ErrorAction SilentlyContinue)) {
                $targetPath = $null
                try {
                    $shortcut = $shell.CreateShortcut($shortcutFile.FullName)
                    $targetPath = [string]$shortcut.TargetPath
                }
                catch {
                    $targetPath = $null
                }

                $results.Add([PSCustomObject]@{
                    ShortcutName = [string]$shortcutFile.Name
                    ShortcutPath = [string]$shortcutFile.FullName
                    TargetPath = $targetPath
                    SourceFolder = [string]$folder
                })
            }
        }
    }
    finally {
        if ($null -ne $shell) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }
    }

    @($results) | Sort-Object SourceFolder, ShortcutName | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

function Export-TaskbarPins {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $taskbarPath = Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'
    Export-ShortcutInventory -Folders @($taskbarPath) -OutputPath $OutputPath
}

function Export-FolderListings {
    param(
        [Parameter(Mandatory)]
        [string[]]$Folders,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    foreach ($folder in $Folders) {
        Write-Section -Title $folder -FilePath $OutputPath
        if (-not (Test-Path -LiteralPath $folder)) {
            Add-Content -LiteralPath $OutputPath -Value 'Path not found.'
            continue
        }

        Get-ChildItem -LiteralPath $folder -Force -ErrorAction SilentlyContinue |
            Select-Object Name, FullName, Mode, LastWriteTime |
            Format-Table -AutoSize |
            Out-String -Width 4096 |
            Add-Content -LiteralPath $OutputPath
    }
}

function Export-StoreApps {
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $apps = @(
        Get-AppxPackage -ErrorAction SilentlyContinue |
            Select-Object Name, Publisher, Version, InstallLocation |
            Sort-Object Name
    )

    $apps | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

function Export-ReinstallChecklist {
    param(
        [Parameter(Mandatory)]
        [object[]]$Programs,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $Programs |
        Select-Object
            @{ Name = 'AppName'; Expression = { $_.DisplayName } },
            @{ Name = 'Version'; Expression = { $_.DisplayVersion } },
            @{ Name = 'Publisher'; Expression = { $_.Publisher } },
            @{ Name = 'Reinstall'; Expression = { '' } },
            @{ Name = 'Priority'; Expression = { '' } },
            @{ Name = 'LicenseOrLoginNeeded'; Expression = { '' } },
            @{ Name = 'Notes'; Expression = { '' } } |
        Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$basePath = Resolve-BackupBasePath -PreferredPath $OutputRoot
$backupFolder = Ensure-Directory -Path (Join-Path -Path $basePath -ChildPath ("Reinstall-App-Backup-$timestamp"))

Write-Host "Saving files to: $backupFolder"
Write-Host ''

$installedCsv = Join-Path -Path $backupFolder -ChildPath 'installed-programs.csv'
$installedDetailedCsv = Join-Path -Path $backupFolder -ChildPath 'installed-programs-detailed.csv'
$installedTxt = Join-Path -Path $backupFolder -ChildPath 'installed-programs.txt'
$wingetJson = Join-Path -Path $backupFolder -ChildPath 'winget-apps.json'
$startupCsv = Join-Path -Path $backupFolder -ChildPath 'startup-items.csv'
$desktopShortcutsCsv = Join-Path -Path $backupFolder -ChildPath 'desktop-shortcuts.csv'
$startMenuShortcutsCsv = Join-Path -Path $backupFolder -ChildPath 'startmenu-shortcuts.csv'
$taskbarPinsCsv = Join-Path -Path $backupFolder -ChildPath 'taskbar-pinned-items.csv'
$appFoldersTxt = Join-Path -Path $backupFolder -ChildPath 'common-app-folders.txt'
$storeAppsCsv = Join-Path -Path $backupFolder -ChildPath 'store-apps.csv'
$reinstallCsv = Join-Path -Path $backupFolder -ChildPath 'reinstall-checklist.csv'
$summaryTxt = Join-Path -Path $backupFolder -ChildPath 'README-summary.txt'

$programExport = Export-InstalledPrograms -CatalogCsvPath $installedCsv -DetailedCsvPath $installedDetailedCsv -TextPath $installedTxt
$wingetResult = if ($SkipWingetExport.IsPresent) {
    [PSCustomObject]@{
        Exported = $false
        Details = 'winget export skipped'
    }
}
else {
    Export-WingetPackageList -OutputPath $wingetJson
}

Export-StartupItems -OutputPath $startupCsv
Export-ShortcutInventory -Folders @([Environment]::GetFolderPath('Desktop')) -OutputPath $desktopShortcutsCsv
Export-ShortcutInventory -Folders @(
    (Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\Windows\Start Menu\Programs'),
    (Join-Path -Path $env:ProgramData -ChildPath 'Microsoft\Windows\Start Menu\Programs')
) -OutputPath $startMenuShortcutsCsv
Export-TaskbarPins -OutputPath $taskbarPinsCsv

if (-not $SkipFolderListings.IsPresent) {
    Export-FolderListings -Folders @(
        'C:\Program Files',
        'C:\Program Files (x86)',
        $env:LOCALAPPDATA,
        $env:APPDATA
    ) -OutputPath $appFoldersTxt
}

if (-not $SkipStoreApps.IsPresent) {
    Export-StoreApps -OutputPath $storeAppsCsv
}

Export-ReinstallChecklist -Programs $programExport.Programs -OutputPath $reinstallCsv

@"
Reinstall App Backup Summary
Generated: $(Get-Date)
Backup folder: $backupFolder

Key file:
- installed-programs.csv is the import-ready file for Initialize-AppCatalog.ps1

Files included:
- installed-programs.csv
- installed-programs-detailed.csv
- installed-programs.txt
- reinstall-checklist.csv
- startup-items.csv
- desktop-shortcuts.csv
- startmenu-shortcuts.csv
- taskbar-pinned-items.csv
- winget-apps.json ($($wingetResult.Details))
- store-apps.csv $(if ($SkipStoreApps.IsPresent) { '(skipped)' } else { '(created when available)' })
- common-app-folders.txt $(if ($SkipFolderListings.IsPresent) { '(skipped)' } else { '(created)' })

Program count: $($programExport.ProgramCount)

Notes:
- Replace the repository root installed-programs.csv with the one from this backup pack when preparing a reinstall catalog.
- Portable apps, browser extensions, plugins, and game libraries may still need separate backup steps.
- Review reinstall-checklist.csv and mark what you actually want back on the rebuilt machine.
"@ | Set-Content -LiteralPath $summaryTxt -Encoding UTF8

Write-Host 'Done.'
Write-Host 'Folder created:'
Write-Host $backupFolder
Write-Host ''
Write-Host 'Start with:'
Write-Host ' - installed-programs.csv'
Write-Host ' - reinstall-checklist.csv'
Write-Host ' - README-summary.txt'