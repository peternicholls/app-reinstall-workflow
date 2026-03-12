# Windows App Reinstall Catalog: Code Evaluation

## Summary

This repository is a practical PowerShell workflow for rebuilding a Windows app setup from an installed-program inventory. It is more than a loose script collection, but it is still best described as a technical-user workflow rather than a polished end-user product.

The current checkout implements a catalog-driven reinstall planner with:

- inventory capture from Windows uninstall data
- catalog initialization and schema backfill
- current-machine sync and status reporting
- rule-based classification and ignore recommendations
- `winget` candidate lookup and latest-version refresh
- local installer discovery and staging
- manual-source guidance and optional manual-reference downloading
- install-plan generation and guarded execution

## What The Code Currently Implements

### Repository state

This checkout contains source scripts, a sample `installed-programs.csv`, placeholder working directories, a supporting evaluation document, and Pester tests under `tests/`.

Generated catalog state and working outputs are not committed:

- `catalog/apps.json` is ignored by Git
- `output/` is a working-data directory
- `staged-installers/` is a local staging area

### Main architecture

The workflow is PowerShell-first and centers on `scripts/lib/AppCatalog.psm1`, which contains most of the domain logic:

- catalog IO and schema validation helpers
- installed-program inventory
- text normalization and app matching
- classification rules
- local installer discovery
- `winget` search parsing and package scoring
- manual reference resolution and download staging

The entry-point scripts stay relatively thin and mostly orchestrate the shared module, which is structurally sound. The tradeoff is that the module is already large enough to be a maintenance hotspot.

### Primary user-facing workflow

`scripts/AppReinstall.ps1` is the intended default entry point. In the current code it exposes these actions:

- `Doctor`
- `Validate`
- `Capture`
- `Initialize`
- `Status`
- `Prepare`
- `Plan`
- `Execute`

That wrapper now adds meaningful workflow value beyond simple script chaining:

- preflight checks and readiness gates
- report writing to `output/`
- automatic catalog initialization during prepare when valid input exists
- optional checklist-based app selection before execution
- support for `winget install` planning and guarded `.exe` handling

### Lower-level scripts

The underlying scripts form a coherent pipeline:

- `Get-InstalledPrograms.ps1` exports installed-program data as table, JSON, or CSV
- `backup-app-list.ps1` creates a timestamped backup pack with inventory and environment references
- `backup-app-settings.ps1` backs up a narrow allowlist of settings folders and color profiles
- `Initialize-AppCatalog.ps1` builds `catalog/apps.json` from `installed-programs.csv`
- `Sync-AppCatalog.ps1` refreshes `installed`, `missing`, and `ignored` state
- `Classify-AppCatalog.ps1` assigns buckets and can auto-apply ignore recommendations
- `Resolve-WingetPackages.ps1` stores likely `winget` candidates and IDs
- `Update-LatestAppVersions.ps1` refreshes `latest.version` for apps with a `wingetId`
- `Find-AppInstallers.ps1` searches local folders for probable installer matches
- `Prepare-AppInstallers.ps1` stages local installers, `winget` downloads, or manual-reference downloads
- `Get-ManualSourceQueue.ps1` writes manual acquisition guidance back to the catalog and to a report
- `Download-ManualReferenceInstallers.ps1` attempts direct downloads from cataloged manual reference URLs
- `Install-PreparedApps.ps1` builds a plan or executes supported installs
- `Set-AppStatus.ps1` provides manual status overrides

### Execution behavior

The install step is narrower and more realistic than a full recovery system:

- automated execution supports `winget install`, `.msi`, and `.exe`
- `.exe` installers require `installer.installArgs` unless the caller supplies `-AllowExeWithoutArgs`
- unsupported or unresolved items are still represented in the plan as skipped work
- `-InteractiveChecklist` allows app-by-app selection before execution

## Evaluation

### What is strong

- The project has a clear working model. `catalog/apps.json` is the right central abstraction for reinstall decisions.
- The wrapper script improves usability. `Doctor`, `Validate`, `Prepare`, `Plan`, and `Execute` form a credible default flow.
- The code is honest about partial automation. It separates ready installers from unresolved or manual work instead of pretending everything is recoverable.
- Generated reports are useful. The workflow emits preflight, sync, classification, resolution, installer, queue, plan, and validation artifacts.
- There is now automated verification. The repository includes Pester tests for catalog validation and install-plan behavior.
- The install path has one important safeguard: unattended `.exe` execution is blocked by default when silent arguments are unknown.

### What is weak

- `scripts/lib/AppCatalog.psm1` is still too large and concentrates most risk in one file.
- Test coverage exists but remains narrow. The current suite does not yet cover classification rules, installer discovery, manual-source logic, or the wrapper flow.
- Matching and classification remain heuristic-heavy. False positives and false negatives are still possible in `winget` resolution and installer selection.
- Manual-source recommendations are still tailored to a specific environment in several places, especially for niche tooling and OEM software.
- Settings backup scope is intentionally small and not yet manifest-driven.
- Safety is suitable for a personal workstation workflow, not for unattended fleet recovery. There is no checksum policy, rollback strategy, or trust verification layer for staged installers.

## Product Fit

The most credible positioning for the current codebase is:

> A Windows app reinstall workflow for technical users who are willing to review a catalog, inspect heuristic matches, and handle unresolved applications manually.

That is a narrower claim than "automatic Windows rebuild," but it is well supported by the implementation that exists today.

## Highest-Value Next Steps

1. Expand Pester coverage around classification, `winget` resolution, manual-source metadata, and queue generation.
2. Split `AppCatalog.psm1` into smaller modules such as inventory, classification, package resolution, discovery, and staging.
3. Move manual-source recommendation rules into a data file or manifest instead of hard-coded branches.
4. Add stronger installer trust checks such as hashes or explicit approval metadata for staged files.
5. Replace the hard-coded settings backup allowlist with a user-editable manifest.

## Verification Limits

This evaluation is based on reading the repository code in the current checkout and the included tests. I did not execute the Windows-specific workflow end to end here because this environment does not provide live Windows registry access or `winget.exe`.
