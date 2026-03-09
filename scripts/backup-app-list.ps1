# backup-app-list.ps1
# Creates a reinstall reference pack for apps/utilities before formatting Windows

$ErrorActionPreference = "SilentlyContinue"

function Ensure-Folder {
    param([string]$Path)
    if (!(Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Write-Section {
    param(
        [string]$Title,
        [string]$File
    )
    Add-Content -Path $File -Value ""
    Add-Content -Path $File -Value ("=" * 80)
    Add-Content -Path $File -Value $Title
    Add-Content -Path $File -Value ("=" * 80)
}

function Get-SafeFolder {
    param([string]$PreferredPath)

    if (Test-Path $PreferredPath) {
        return $PreferredPath
    }

    $fallback = Join-Path $env:USERPROFILE "Documents"
    if (Test-Path $fallback) {
        return $fallback
    }

    return $env:TEMP
}

function Export-InstalledPrograms {
    param([string]$OutCsv, [string]$OutTxt)

    $programs = Get-ItemProperty `
        HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
        HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, `
        HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Where-Object { $_.DisplayName } |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation, UninstallString |
        Sort-Object DisplayName -Unique

    $programs | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
    $programs | Format-Table -AutoSize | Out-String -Width 4096 | Set-Content -Path $OutTxt -Encoding UTF8

    return $programs
}

function Export-Winget {
    param([string]$OutJson)

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        winget export -o $OutJson --include-versions --accept-source-agreements
        return $true
    }
    else {
        return $false
    }
}

function Export-StartupApps {
    param([string]$OutCsv)

    $startupItems = @()

    $registryPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
    )

    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            $item = Get-ItemProperty -Path $path
            foreach ($prop in $item.PSObject.Properties) {
                if ($prop.Name -notmatch "^PS") {
                    $startupItems += [PSCustomObject]@{
                        Source = $path
                        Name   = $prop.Name
                        Value  = $prop.Value
                    }
                }
            }
        }
    }

    $startupFolders = @(
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
    )

    foreach ($folder in $startupFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Force | ForEach-Object {
                $startupItems += [PSCustomObject]@{
                    Source = $folder
                    Name   = $_.Name
                    Value  = $_.FullName
                }
            }
        }
    }

    $startupItems | Sort-Object Source, Name | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

function Export-Shortcuts {
    param(
        [string[]]$Folders,
        [string]$OutCsv
    )

    $wsh = New-Object -ComObject WScript.Shell
    $results = @()

    foreach ($folder in $Folders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Filter *.lnk -Force | ForEach-Object {
                $targetPath = $null
                try {
                    $shortcut = $wsh.CreateShortcut($_.FullName)
                    $targetPath = $shortcut.TargetPath
                } catch {}

                $results += [PSCustomObject]@{
                    ShortcutName = $_.Name
                    ShortcutPath = $_.FullName
                    TargetPath   = $targetPath
                    SourceFolder = $folder
                }
            }
        }
    }

    $results | Sort-Object ShortcutName | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

function Export-TaskbarPins {
    param([string]$OutCsv)

    $taskbarPath = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

    $wsh = New-Object -ComObject WScript.Shell
    $results = @()

    if (Test-Path $taskbarPath) {
        Get-ChildItem -Path $taskbarPath -Filter *.lnk -Force | ForEach-Object {
            $targetPath = $null
            try {
                $shortcut = $wsh.CreateShortcut($_.FullName)
                $targetPath = $shortcut.TargetPath
            } catch {}

            $results += [PSCustomObject]@{
                ShortcutName = $_.Name
                ShortcutPath = $_.FullName
                TargetPath   = $targetPath
            }
        }
    }

    $results | Sort-Object ShortcutName | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

function Export-FolderListings {
    param(
        [string[]]$Folders,
        [string]$OutTxt
    )

    if (Test-Path $OutTxt) {
        Remove-Item $OutTxt -Force
    }

    foreach ($folder in $Folders) {
        Write-Section -Title $folder -File $OutTxt
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Force |
                Select-Object Name, FullName, Mode, LastWriteTime |
                Format-Table -AutoSize |
                Out-String -Width 4096 |
                Add-Content -Path $OutTxt
        }
        else {
            Add-Content -Path $OutTxt -Value "Path not found."
        }
    }
}

function Export-StoreApps {
    param([string]$OutCsv)

    $apps = Get-AppxPackage |
        Select-Object Name, Publisher, Version, InstallLocation |
        Sort-Object Name

    $apps | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

function Create-ReinstallChecklist {
    param(
        [object[]]$Programs,
        [string]$OutCsv
    )

    $checklist = $Programs | Select-Object `
        @{Name="AppName";Expression={$_.DisplayName}},
        @{Name="Version";Expression={$_.DisplayVersion}},
        @{Name="Publisher";Expression={$_.Publisher}},
        @{Name="Reinstall";Expression={""}},
        @{Name="Priority";Expression={""}},
        @{Name="LicenseOrLoginNeeded";Expression={""}},
        @{Name="Notes";Expression={""}}

    $checklist | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

# Decide base output folder
$preferredBase = Join-Path $env:USERPROFILE "OneDrive\Documents"
$baseFolder = Get-SafeFolder -PreferredPath $preferredBase

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$outFolder = Join-Path $baseFolder "Reinstall-App-Backup-$timestamp"
Ensure-Folder -Path $outFolder

Write-Host "Saving files to: $outFolder"
Write-Host ""

# Output file paths
$installedCsv       = Join-Path $outFolder "installed-programs.csv"
$installedTxt       = Join-Path $outFolder "installed-programs.txt"
$wingetJson         = Join-Path $outFolder "winget-apps.json"
$startupCsv         = Join-Path $outFolder "startup-items.csv"
$desktopShortcuts   = Join-Path $outFolder "desktop-shortcuts.csv"
$startMenuShortcuts = Join-Path $outFolder "startmenu-shortcuts.csv"
$taskbarPinsCsv     = Join-Path $outFolder "taskbar-pinned-items.csv"
$appFoldersTxt      = Join-Path $outFolder "common-app-folders.txt"
$storeAppsCsv       = Join-Path $outFolder "store-apps.csv"
$reinstallCsv       = Join-Path $outFolder "reinstall-checklist.csv"
$summaryTxt         = Join-Path $outFolder "README-summary.txt"

# Export data
$programs = Export-InstalledPrograms -OutCsv $installedCsv -OutTxt $installedTxt
$wingetOk = Export-Winget -OutJson $wingetJson

Export-StartupApps -OutCsv $startupCsv

Export-Shortcuts `
    -Folders @(
        [Environment]::GetFolderPath("Desktop")
    ) `
    -OutCsv $desktopShortcuts

Export-Shortcuts `
    -Folders @(
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    ) `
    -OutCsv $startMenuShortcuts

Export-TaskbarPins -OutCsv $taskbarPinsCsv

Export-FolderListings `
    -Folders @(
        "C:\Program Files",
        "C:\Program Files (x86)",
        "$env:LOCALAPPDATA",
        "$env:APPDATA"
    ) `
    -OutTxt $appFoldersTxt

Export-StoreApps -OutCsv $storeAppsCsv
Create-ReinstallChecklist -Programs $programs -OutCsv $reinstallCsv

# Summary
@"
Reinstall App Backup Summary
Generated: $(Get-Date)

Files included:
- installed-programs.csv
- installed-programs.txt
- reinstall-checklist.csv
- startup-items.csv
- desktop-shortcuts.csv
- startmenu-shortcuts.csv
- taskbar-pinned-items.csv
- store-apps.csv
- common-app-folders.txt
- winget-apps.json $(if ($wingetOk) { "(created)" } else { "(winget not available)" })

Notes:
- Portable apps may still need checking manually.
- Browser extensions are not included here.
- Game libraries and plugins may need separate backup/checking.
- Review reinstall-checklist.csv and mark what you actually want back.
"@ | Set-Content -Path $summaryTxt -Encoding UTF8

Write-Host "Done."
Write-Host "Folder created:"
Write-Host $outFolder
Write-Host ""
Write-Host "Open this folder and start with:"
Write-Host " - reinstall-checklist.csv"
Write-Host " - README-summary.txt"