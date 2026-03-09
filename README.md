# App Reinstall Workflow

This workspace now treats `installed-programs.csv` as a source snapshot and keeps the working inventory in `catalog/apps.json`.

The JSON catalog is easier to maintain because each app can carry status, installer metadata, aliases, and staging information in one place.

## Scripts

Run these from the workspace root with `pwsh.exe` on Windows.

`pwsh.exe -File .\scripts\Initialize-AppCatalog.ps1`

Creates `catalog/apps.json` from `installed-programs.csv`.

`pwsh.exe -File .\scripts\Sync-AppCatalog.ps1 -View Summary`

Checks the machine registry and marks apps in the catalog as `installed`, `missing`, or `ignored`.

`pwsh.exe -File .\scripts\Sync-AppCatalog.ps1 -View Missing`

Shows only the apps that are currently missing.

`pwsh.exe -File .\scripts\Set-AppStatus.ps1 -Name "Google Chrome" -Status missing -Exact`

Manually overrides one catalog entry.

`pwsh.exe -File .\scripts\Find-AppInstallers.ps1 -SearchRoot C:\Users\engli\Downloads,C:\Installers -ReportPath .\output\installer-report.json`

Searches for installer files and records the best local matches in the catalog.

`pwsh.exe -File .\scripts\Resolve-WingetPackages.ps1 -MaxApps 20`

Searches winget for likely package IDs and stores the best candidates in the catalog.

`pwsh.exe -File .\scripts\Update-LatestAppVersions.ps1 -MaxApps 20`

Refreshes `latest.version` for apps that already have a `wingetId`, so the catalog keeps a current available version separate from the CSV snapshot version.

`pwsh.exe -File .\scripts\Classify-AppCatalog.ps1 -ApplyIgnoreRecommendations`

Classifies apps into application, runtime, sdk, system-component, driver, browser, or developer-tool buckets and can auto-ignore obvious non-reinstall targets.

`pwsh.exe -File .\scripts\Prepare-AppInstallers.ps1 -DownloadWithWinget`

Stages local installers into `output/staged-installers` and optionally downloads installers via `winget` for apps where `installer.wingetId` is set.

`pwsh.exe -File .\scripts\Install-PreparedApps.ps1 -Mode Plan`

Builds an install plan from the staged queue. Use `-Mode Execute` to actually run ready installers.

`pwsh.exe -File .\scripts\Get-ManualSourceQueue.ps1 -Format Table`

Builds a manual-acquisition queue for missing apps that have neither a safe `wingetId` nor a local installer candidate, with source hints for vendor portals, Microsoft components, OEM drivers, or legacy archives.

`pwsh.exe -File .\scripts\Get-ManualSourceQueue.ps1 -Format Table -UpdateCatalog`

Builds the same manual-acquisition queue and also writes the current manual-source recommendation back into `catalog/apps.json` under `installer.manualAcquisitionType`, `installer.manualSourceHint`, `installer.manualReferenceUrl`, `installer.manualReason`, and `installer.manualUpdatedAt`.

## Catalog Shape

Each app entry includes:

- `name`, `publisher`, `expectedVersion`
- `latest.version`, `latest.source`, `latest.checkedAt`, `latest.packageId`
- `desired` and `status`
- `classification.bucket`, `classification.recommendedAction`, `classification.reason`
- `detection.matchNames` for alternate display names
- `installer.localPath`, `installer.localCandidates`, `installer.wingetCandidates`, `installer.wingetId`, `installer.downloadedPath`, `installer.installArgs`, `installer.ready`
- `installer.manualAcquisitionType`, `installer.manualSourceHint`, `installer.manualReferenceUrl`, `installer.manualReason`, `installer.manualUpdatedAt` for unresolved apps that need vendor or Microsoft download pages

## Suggested Workflow

1. Initialize the catalog once from the CSV.
2. Run the sync script whenever you want a fresh installed vs missing report.
3. Run the classification script to separate real apps from runtimes, SDK fragments, and system components.
4. Run the winget resolver to fill in likely package IDs for missing apps.
5. Run the latest-version refresh script so current available versions are tracked independently from `expectedVersion`.
6. Add alternate display names or adjust `installer.wingetId` values in `catalog/apps.json` where matching needs manual cleanup.
7. Run installer discovery against your downloads or software archive folders.
8. Run the prepare script to copy or download installers into a staging folder.
9. Run the manual-source queue script for the apps that still have no package source or local installer.
10. Run the install planner, then execute only when the commands look correct.