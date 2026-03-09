# Windows App Reinstall Catalog

This repository contains a PowerShell workflow for capturing a Windows software inventory, converting it into a structured catalog, enriching that catalog with installer metadata, and preparing a reinstall plan.

It is intended for scenarios such as:

- rebuilding a workstation after a clean install
- documenting the applications on a machine before migration or replacement
- tracking which apps can be restored automatically and which still require manual download steps

The workflow centers on `catalog/apps.json`, a machine-readable catalog that keeps installation state, package candidates, local installer matches, classification data, and manual follow-up notes in one place.

## What The Repository Does

The scripts in this repository help you:

- export the currently installed Windows programs to CSV
- initialize a JSON catalog from that snapshot
- re-scan the machine and mark apps as installed, missing, or ignored
- classify entries so runtimes and system components can be separated from user applications
- look up likely `winget` package IDs and latest available versions
- discover matching installer files in local folders
- stage installers into a single working directory
- generate a manual-source queue for apps that cannot be resolved automatically
- optionally fetch installers from curated manual reference URLs when direct downloads are available
- build an install plan and optionally execute it

## Requirements

- Windows
- PowerShell 7 or Windows PowerShell with script execution enabled for local scripts
- `winget` for package lookup, version refresh, and optional download/install flows

Run commands from the repository root:

```powershell
pwsh -File .\scripts\Sync-AppCatalog.ps1 -View Summary
```

## Repository Layout

| Path | Purpose |
| --- | --- |
| `scripts/` | Entry-point scripts for inventory, catalog maintenance, staging, and installation |
| `scripts/lib/AppCatalog.psm1` | Shared helper functions used by all scripts |
| `catalog/apps.json` | Primary working catalog |
| `installed-programs.csv` | Source inventory snapshot |
| `output/` | Generated reports, queues, logs, and staged installers |

`output/` is ignored by Git except for a placeholder, so generated reports and staged installers do not need to be committed.

## Quick Start

### 1. Export the current software inventory

```powershell
pwsh -File .\scripts\Get-InstalledPrograms.ps1 -Format Csv -OutputPath .\installed-programs.csv
```

### 2. Build the catalog

```powershell
pwsh -File .\scripts\Initialize-AppCatalog.ps1
```

This creates `catalog/apps.json` from the CSV snapshot.

### 3. Sync the catalog against the current machine

```powershell
pwsh -File .\scripts\Sync-AppCatalog.ps1 -View Summary
pwsh -File .\scripts\Sync-AppCatalog.ps1 -View Missing
```

### 4. Classify entries

```powershell
pwsh -File .\scripts\Classify-AppCatalog.ps1 -ApplyIgnoreRecommendations
```

This helps remove obvious runtimes, SDK fragments, and system-managed components from the reinstall target list.

### 5. Resolve package IDs and latest versions

```powershell
pwsh -File .\scripts\Resolve-WingetPackages.ps1 -MaxApps 20
pwsh -File .\scripts\Update-LatestAppVersions.ps1 -MaxApps 20
```

### 6. Search local folders for installers

```powershell
pwsh -File .\scripts\Find-AppInstallers.ps1 -SearchRoot C:\Installers,$env:USERPROFILE\Downloads
```

### 7. Stage installers into a working folder

```powershell
pwsh -File .\scripts\Prepare-AppInstallers.ps1 -DownloadWithWinget
```

### 8. Review apps that still need manual acquisition

```powershell
pwsh -File .\scripts\Get-ManualSourceQueue.ps1 -Format Table -UpdateCatalog
```

### 9. Optionally fetch installers from manual reference URLs

```powershell
pwsh -File .\scripts\Download-ManualReferenceInstallers.ps1 -MaxApps 5
```

### 10. Build an install plan

```powershell
pwsh -File .\scripts\Install-PreparedApps.ps1 -Mode Plan
```

Review the generated plan before any execution step.

### 11. Execute when ready

```powershell
pwsh -File .\scripts\Install-PreparedApps.ps1 -Mode Execute
```

## Script Reference

| Script | Purpose |
| --- | --- |
| `Get-InstalledPrograms.ps1` | Reads installed-program data from the local machine and outputs table, JSON, or CSV |
| `Initialize-AppCatalog.ps1` | Creates `catalog/apps.json` from `installed-programs.csv` |
| `Sync-AppCatalog.ps1` | Re-checks the machine and updates each app status |
| `Classify-AppCatalog.ps1` | Assigns buckets such as application, runtime, sdk, driver, browser, or developer-tool |
| `Resolve-WingetPackages.ps1` | Finds likely `winget` package candidates and stores the best match |
| `Update-LatestAppVersions.ps1` | Refreshes `latest.version` for apps that already have a `wingetId` |
| `Find-AppInstallers.ps1` | Searches local folders for likely installer files |
| `Prepare-AppInstallers.ps1` | Copies or downloads installers into `output/staged-installers` and writes `output/install-queue.json` |
| `Get-ManualSourceQueue.ps1` | Produces a queue of unresolved apps that still need vendor, Microsoft, OEM, or archive lookup |
| `Download-ManualReferenceInstallers.ps1` | Attempts to download installers from `installer.manualReferenceUrl` and stage them for later execution |
| `Set-AppStatus.ps1` | Manually overrides one catalog entry |
| `Install-PreparedApps.ps1` | Generates an install plan and can run supported installers |

## Catalog Model

Each app entry in `catalog/apps.json` includes fields such as:

- `name`, `publisher`, `expectedVersion`
- `desired` and `status`
- `detection.*` for last-seen metadata and match names
- `classification.*` for bucket, recommendation, and rationale
- `latest.*` for current package version information
- `installer.*` for local candidates, selected installer path, `winget` metadata, manual-source hints, staging path, and readiness
- `notes` for any manual annotations

This structure is designed to keep all reinstall-related decisions close to the application record they affect.

## Reports And Generated Files

The scripts can generate artifacts such as:

- `output/winget-report.json`
- `output/classification-report.json`
- `output/manual-source-queue.json`
- `output/manual-download-queue.json`
- `output/install-queue.json`
- `output/install-log.json`
- `output/staged-installers/`

These files are intended as working data, not long-term source files.

## Safety Notes

- Prefer `Install-PreparedApps.ps1 -Mode Plan` before any execution.
- EXE installers may require explicit silent arguments in `installer.installArgs` before they are safe to run unattended.
- Review `winget` matches and local installer candidates before trusting them as final.
- Device drivers, OEM utilities, and licensed software often need vendor-specific handling even when they appear in the catalog.

## Privacy And Publishing Notes

This workflow operates on machine-derived data. Files such as `installed-programs.csv`, `catalog/apps.json`, generated reports, and staged installer paths can reveal:

- installed software inventory
- usernames and local file paths
- internal tooling names
- vendor portals or licensed software usage

If you use this repository publicly, treat those files as environment-specific data and sanitize or exclude them before publishing.

## Limitations

- The repository is Windows-focused.
- `winget` matching is heuristic and may need manual correction.
- Automatic execution currently covers `msi` packages, `winget install`, and `exe` installers when silent arguments are known.
- Some applications will always remain manual because they are licensed, OEM-specific, legacy, or not available from a trusted package source.