# Windows App Reinstall Catalog: Code Evaluation

## Summary

This repository is a practical PowerShell workflow for rebuilding a Windows app setup from an installed-program inventory. It is more than a script collection, but it is not yet a broadly reusable product. Based on the code that is actually present, the project is best described as a catalog-driven reinstall planner for a technical Windows user.

The core design is sound:

- capture installed programs from Windows uninstall registry keys
- normalize them into `catalog/apps.json`
- classify entries and mark non-targets
- try to resolve `winget` packages and local installers
- generate manual-source recommendations for unresolved apps
- stage installers and build an install queue
- optionally execute supported installs

## What The Code Actually Implements

### Repository state

This checkout contains source scripts, a sample `installed-programs.csv`, and placeholder directories for `catalog/`, `output/`, and `staged-installers/`. The earlier report assumed generated catalog data and staged artifacts were committed; that is not true in the current repository.

### Main architecture

The workflow is PowerShell-first and centers on one large shared module: `scripts/lib/AppCatalog.psm1`. That module is about 1,400 lines and contains most of the real logic:

- catalog IO and schema backfill
- installed-program inventory and matching
- app classification
- local installer discovery
- `winget` search parsing and package scoring
- manual download-page scraping and staging helpers

The entry-point scripts are thin wrappers around that module, which is good structurally, although the shared module is already large enough to be a maintenance risk.

### Workflow scripts

The scripts in the repo implement a coherent pipeline:

- `Get-InstalledPrograms.ps1` reads uninstall registry keys and exports table, JSON, or CSV
- `backup-app-list.ps1` creates a backup folder with app inventory, startup items, shortcuts, taskbar pins, optional Store app export, folder listings, and a reinstall checklist
- `backup-app-settings.ps1` backs up only a very small allowlist: DisplayCAL, Calman, and optional Windows color profiles
- `Initialize-AppCatalog.ps1` groups CSV rows by normalized display name and publisher, then builds `catalog/apps.json`
- `Sync-AppCatalog.ps1` re-detects installed apps and sets `installed`, `missing`, or `ignored`
- `Classify-AppCatalog.ps1` applies rule-based buckets and can auto-ignore recommended entries
- `Resolve-WingetPackages.ps1` searches `winget` and stores scored candidates
- `Update-LatestAppVersions.ps1` refreshes versions for apps that already have a `winget` ID
- `Find-AppInstallers.ps1` scans common folders and scores installer filename matches
- `Get-ManualSourceQueue.ps1` produces manual acquisition guidance and can write it back into the catalog
- `Prepare-AppInstallers.ps1` stages local installers, `winget` downloads, or manual-reference downloads into `staged-installers/`
- `Install-PreparedApps.ps1` builds a plan or executes supported installs
- `Set-AppStatus.ps1` provides a manual status override

### Execution behavior

The install step is narrower than the original report implied:

- automated execution supports `winget install`, `.msi`, and `.exe`
- `.exe` requires `installer.installArgs` unless the caller uses `-AllowExeWithoutArgs`
- other installer types may be staged, but they are not executed automatically
- there is an interactive checklist UI for selecting apps before execution

## Evaluation

### What is strong

- The pipeline is clear and consistent. The project has a real model instead of a loose set of scripts.
- `catalog/apps.json` is the right abstraction. Detection state, classification, package metadata, and installer metadata live together.
- The repo is honest about partial automation. It supports `winget`, local installers, and manual follow-up instead of pretending everything can be restored automatically.
- The backup-list script is useful. It captures more context than just installed apps.
- The install script is cautious in one important area: unattended `.exe` installs are blocked unless arguments are known or explicitly overridden.

### What is weak

- The shared module is too large. Roughly all domain logic is concentrated in one file, which makes change risk higher.
- There is no test suite. I found no Pester tests or other automated verification in the repo.
- Matching and classification are heuristic-heavy. Name normalization, token scoring, `winget` ranking, and installer matching all rely on embedded rules.
- Manual-source recommendations are partly hard-coded to a niche software set, including automotive, calibration, printer, and Microsoft developer-tool cases. That makes the project feel tailored to one user environment.
- Settings backup scope is narrow. The current code does not back up common items like browser profiles, Git config, SSH config, VS Code, Windows Terminal, or PowerShell profiles.
- Safety is decent for a personal tool but not strong enough for a wider audience. There is no checksum tracking, trust verification, elevation planning, rollback strategy, or preflight validation layer.

## Product Fit

This project does not currently support the claim "works for any Windows user." The code supports a narrower and more credible claim:

> A Windows app reinstall planner for technical users who are willing to review a catalog, inspect heuristic matches, and handle unresolved apps manually.

That is still a worthwhile project. The implementation already fits that scope reasonably well.

## Recommended Rewrite Of The Project Positioning

This repository is a PowerShell-based Windows reinstall workflow that captures installed programs, builds a catalog of desired apps, resolves likely package or installer sources, stages supported installers, and generates a reviewable install plan.

It is best suited to personal workstation rebuilds, not unattended recovery for arbitrary Windows environments.

## Highest-Value Next Steps

1. Add Pester tests for catalog initialization, name matching, classification, `winget` parsing, and queue generation.
2. Split `AppCatalog.psm1` into smaller modules such as inventory, classification, package resolution, installer discovery, and staging.
3. Replace hard-coded manual-source rules with a data file or manifest.
4. Add preflight validation for required tools, admin needs, missing files, and unsupported installer types.
5. Expand settings backup through an explicit allowlist manifest instead of app-specific code.

## Verification Limits

This evaluation is based on reading the repository code in the current checkout. I did not execute the workflow here because the environment is not a Windows PowerShell runtime with registry access and `winget.exe` available.
