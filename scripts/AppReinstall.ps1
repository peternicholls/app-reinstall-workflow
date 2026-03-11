<#
.SYNOPSIS
Guided entry point for the Windows app reinstall workflow.

.DESCRIPTION
Wraps the lower-level scripts in this repository behind a smaller set of user-facing actions.
Use this script for the default workflow and keep the individual scripts for advanced or
troubleshooting scenarios.

.EXAMPLE
pwsh -File .\scripts\AppReinstall.ps1 -Action Doctor

.EXAMPLE
pwsh -File .\scripts\AppReinstall.ps1 -Action Prepare

.EXAMPLE
pwsh -File .\scripts\AppReinstall.ps1 -Action Plan

.EXAMPLE
pwsh -File .\scripts\AppReinstall.ps1 -Action Execute -InteractiveChecklist
#>
[CmdletBinding()]
param(
    [ValidateSet('Help', 'Doctor', 'Capture', 'Initialize', 'Status', 'Prepare', 'Plan', 'Execute')]
    [string]$Action = 'Help',

    [string]$CsvPath,

    [string]$CatalogPath,

    [string]$QueuePath,

    [string]$StageDirectory,

    [string]$PlanPath,

    [string]$LogPath,

    [string]$DoctorReportPath,

    [ValidateSet('Inventory', 'BackupPack')]
    [string]$CaptureMode = 'Inventory',

    [string]$OutputRoot,

    [string[]]$SearchRoot,

    [int]$MaxApps = 0,

    [int]$CandidatesPerApp = 5,

    [switch]$SkipWingetDownload,

    [switch]$DownloadFromManualReferences,

    [switch]$SkipIgnoreRecommendations,

    [switch]$IncludeInstalled,

    [switch]$Force,

    [switch]$InteractiveChecklist,

    [switch]$UseWingetWhenAvailable,

    [switch]$AllowExeWithoutArgs,

    [switch]$IncludeSettingsBackup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

function Resolve-WorkflowPath {
    param(
        [AllowNull()]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$DefaultPath
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $DefaultPath
    }

    return Resolve-AppPath -Path $Path
}

function Invoke-WorkflowScript {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptName,

        [Parameter(Mandatory)]
        [string]$StepName,

        [hashtable]$Parameters
    )

    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath $ScriptName
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "Workflow script not found: $scriptPath"
    }

    Write-Host ''
    Write-Host ('=> {0}' -f $StepName)

    if ($null -eq $Parameters -or $Parameters.Count -eq 0) {
        return & $scriptPath
    }

    return & $scriptPath @Parameters
}

function New-DoctorCheck {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('pass', 'warning', 'fail')]
        [string]$Status,

        [Parameter(Mandatory)]
        [string]$Details
    )

    return [PSCustomObject]@{
        name = $Name
        status = $Status
        details = $Details
    }
}

function Test-IsProcessElevated {
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        return $false
    }

    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Test-PathWriteAccess {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$DirectoryPath
    )

    $targetDirectory = if ($DirectoryPath.IsPresent) {
        $Path
    }
    else {
        Split-Path -Path $Path -Parent
    }

    if ([string]::IsNullOrWhiteSpace($targetDirectory)) {
        return [PSCustomObject]@{
            canWrite = $false
            details = 'No writable directory could be determined'
        }
    }

    try {
        if (-not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        $probePath = Join-Path -Path $targetDirectory -ChildPath ([System.IO.Path]::GetRandomFileName())
        Set-Content -LiteralPath $probePath -Value 'write-test' -Encoding UTF8
        Remove-Item -LiteralPath $probePath -Force

        return [PSCustomObject]@{
            canWrite = $true
            details = $targetDirectory
        }
    }
    catch {
        return [PSCustomObject]@{
            canWrite = $false
            details = $_.Exception.Message
        }
    }
}

function Get-EffectiveMaxApps {
    param(
        [Parameter(Mandatory)]
        [string]$ResolvedCatalogPath,

        [Parameter(Mandatory)]
        [int]$RequestedMaxApps,

        [switch]$ForLatestVersionRefresh,

        [switch]$IncludeInstalledApps
    )

    if ($RequestedMaxApps -gt 0) {
        return $RequestedMaxApps
    }

    $catalog = Read-AppCatalog -CatalogPath $ResolvedCatalogPath
    if ($ForLatestVersionRefresh.IsPresent) {
        return @(
            $catalog.apps |
                Where-Object {
                    (-not [string]::IsNullOrWhiteSpace([string]$_.installer.wingetId)) -and
                    ($IncludeInstalledApps.IsPresent -or $_.status -ne 'installed')
                }
        ).Count
    }

    return @(
        $catalog.apps |
            Where-Object {
                ($IncludeInstalledApps.IsPresent -or $_.status -ne 'installed') -and
                $_.desired -ne $false
            }
    ).Count
}

function Test-WorkflowPrerequisites {
    param(
        [Parameter(Mandatory)]
        [string]$ResolvedCsvPath,

        [Parameter(Mandatory)]
        [string]$ResolvedCatalogPath,

        [Parameter(Mandatory)]
        [string]$ResolvedQueuePath,

        [Parameter(Mandatory)]
        [string]$ResolvedStageDirectory,

        [Parameter(Mandatory)]
        [string]$ResolvedPlanPath,

        [Parameter(Mandatory)]
        [string]$ResolvedExecutionLogPath
    )

    $checks = New-Object 'System.Collections.Generic.List[object]'
    $catalogExists = Test-Path -LiteralPath $ResolvedCatalogPath
    $catalogValid = $false
    $catalogSummary = $null
    $csvExists = Test-Path -LiteralPath $ResolvedCsvPath
    $queueExists = Test-Path -LiteralPath $ResolvedQueuePath
    $queueValid = $false
    $queueItemCount = 0
    $readyQueueItemCount = 0
    $missingStagedFileCount = 0
    $unreadyQueueItemCount = 0
    $stageDirectoryExists = Test-Path -LiteralPath $ResolvedStageDirectory
    $planPathWriteAccess = Test-PathWriteAccess -Path $ResolvedPlanPath
    $executionLogWriteAccess = Test-PathWriteAccess -Path $ResolvedExecutionLogPath
    $stageDirectoryWriteAccess = Test-PathWriteAccess -Path $ResolvedStageDirectory -DirectoryPath

    $isWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
    if ($isWindows) {
        $checks.Add((New-DoctorCheck -Name 'Operating system' -Status 'pass' -Details 'Windows host detected'))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Operating system' -Status 'fail' -Details 'This workflow requires a Windows host'))
    }

    $edition = [string]$PSVersionTable.PSEdition
    $version = [string]$PSVersionTable.PSVersion
    $supportedPowerShell = (
        ($edition -eq 'Desktop' -and $PSVersionTable.PSVersion.Major -ge 5) -or
        ($edition -eq 'Core' -and $PSVersionTable.PSVersion.Major -ge 7)
    )

    if ($supportedPowerShell) {
        $checks.Add((New-DoctorCheck -Name 'PowerShell version' -Status 'pass' -Details ("{0} {1}" -f $edition, $version)))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'PowerShell version' -Status 'fail' -Details ("Unsupported PowerShell runtime: {0} {1}" -f $edition, $version)))
    }

    if ($isWindows) {
        if (Test-IsProcessElevated) {
            $checks.Add((New-DoctorCheck -Name 'Administrator privileges' -Status 'pass' -Details 'Elevated PowerShell session detected'))
        }
        else {
            $checks.Add((New-DoctorCheck -Name 'Administrator privileges' -Status 'warning' -Details 'Not elevated; some installers may fail or prompt during execution'))
        }
    }

    $wingetCommand = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($null -ne $wingetCommand) {
        $checks.Add((New-DoctorCheck -Name 'winget availability' -Status 'pass' -Details $wingetCommand.Source))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'winget availability' -Status 'fail' -Details 'winget.exe was not found in PATH'))
    }

    if ($csvExists) {
        $checks.Add((New-DoctorCheck -Name 'Installed program export' -Status 'pass' -Details $ResolvedCsvPath))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Installed program export' -Status 'warning' -Details 'installed-programs.csv is missing; prepare can only continue if a catalog already exists'))
    }

    if ($catalogExists) {
        try {
            $catalog = Read-AppCatalog -CatalogPath $ResolvedCatalogPath
            $catalogSummary = [PSCustomObject]@{
                schemaVersion = $catalog.schemaVersion
                appCount = @($catalog.apps).Count
            }
            $catalogValid = $true
            $checks.Add((New-DoctorCheck -Name 'Catalog file' -Status 'pass' -Details ("schemaVersion={0}; apps={1}" -f [string]$catalog.schemaVersion, @($catalog.apps).Count)))
        }
        catch {
            $checks.Add((New-DoctorCheck -Name 'Catalog file' -Status 'fail' -Details $_.Exception.Message))
        }
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Catalog file' -Status 'warning' -Details 'catalog/apps.json does not exist yet'))
    }

    if ($queueExists) {
        try {
            $queueDocument = Get-Content -LiteralPath $ResolvedQueuePath -Raw | ConvertFrom-Json
            $queueItems = @($queueDocument.items)
            $queueItemCount = $queueItems.Count
            $readyQueueItemCount = @($queueItems | Where-Object { $_.ready }).Count
            $unreadyQueueItemCount = @($queueItems | Where-Object { -not $_.ready }).Count
            $missingStagedFileCount = @(
                $queueItems |
                    Where-Object {
                        (-not [string]::IsNullOrWhiteSpace([string]$_.stagedPath)) -and
                        (-not (Test-Path -LiteralPath ([string]$_.stagedPath)))
                    }
            ).Count
            $queueValid = $true
            $checks.Add((New-DoctorCheck -Name 'Install queue' -Status 'pass' -Details ("items={0}; ready={1}; unready={2}; path={3}" -f $queueItemCount, $readyQueueItemCount, $unreadyQueueItemCount, $ResolvedQueuePath)))

            if ($missingStagedFileCount -eq 0) {
                $checks.Add((New-DoctorCheck -Name 'Staged installer paths' -Status 'pass' -Details 'All staged installer paths referenced by the queue exist'))
            }
            else {
                $checks.Add((New-DoctorCheck -Name 'Staged installer paths' -Status 'warning' -Details ("{0} queue items reference missing staged files" -f $missingStagedFileCount)))
            }
        }
        catch {
            $checks.Add((New-DoctorCheck -Name 'Install queue' -Status 'fail' -Details $_.Exception.Message))
        }
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Install queue' -Status 'warning' -Details 'output/install-queue.json does not exist yet'))
    }

    if ($stageDirectoryExists) {
        $checks.Add((New-DoctorCheck -Name 'Stage directory' -Status 'pass' -Details $ResolvedStageDirectory))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Stage directory' -Status 'warning' -Details 'staged-installers does not exist yet; prepare will create it'))
    }

    if ($planPathWriteAccess.canWrite) {
        $checks.Add((New-DoctorCheck -Name 'Plan output path' -Status 'pass' -Details $planPathWriteAccess.details))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Plan output path' -Status 'fail' -Details $planPathWriteAccess.details))
    }

    if ($executionLogWriteAccess.canWrite) {
        $checks.Add((New-DoctorCheck -Name 'Execution log path' -Status 'pass' -Details $executionLogWriteAccess.details))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Execution log path' -Status 'fail' -Details $executionLogWriteAccess.details))
    }

    if ($stageDirectoryWriteAccess.canWrite) {
        $checks.Add((New-DoctorCheck -Name 'Stage directory write access' -Status 'pass' -Details $stageDirectoryWriteAccess.details))
    }
    else {
        $checks.Add((New-DoctorCheck -Name 'Stage directory write access' -Status 'fail' -Details $stageDirectoryWriteAccess.details))
    }

    $failedCheckCount = @($checks | Where-Object { $_.status -eq 'fail' }).Count
    $prepareReady = $isWindows -and $supportedPowerShell -and ($null -ne $wingetCommand) -and (($catalogValid) -or $csvExists) -and $stageDirectoryWriteAccess.canWrite
    $planReady = $isWindows -and $supportedPowerShell -and $catalogValid -and $queueValid -and $planPathWriteAccess.canWrite
    $executeReady = $planReady -and $executionLogWriteAccess.canWrite -and ($readyQueueItemCount -gt 0) -and ($missingStagedFileCount -eq 0)

    return [PSCustomObject]@{
        generatedAt = (Get-Date).ToString('o')
        checks = @($checks)
        failedChecks = $failedCheckCount
        paths = [PSCustomObject]@{
            csvPath = $ResolvedCsvPath
            catalogPath = $ResolvedCatalogPath
            queuePath = $ResolvedQueuePath
            stageDirectory = $ResolvedStageDirectory
            planPath = $ResolvedPlanPath
            executionLogPath = $ResolvedExecutionLogPath
        }
        catalog = [PSCustomObject]@{
            exists = $catalogExists
            valid = $catalogValid
            summary = $catalogSummary
        }
        queue = [PSCustomObject]@{
            exists = $queueExists
            valid = $queueValid
            itemCount = $queueItemCount
            readyItemCount = $readyQueueItemCount
            unreadyItemCount = $unreadyQueueItemCount
            missingStagedFileCount = $missingStagedFileCount
        }
        readiness = [PSCustomObject]@{
            prepare = $prepareReady
            plan = $planReady
            execute = $executeReady
        }
    }
}

function Write-DoctorReport {
    param(
        [Parameter(Mandatory)]
        $DoctorResult,

        [Parameter(Mandatory)]
        [string]$ResolvedReportPath
    )

    $reportDirectory = Split-Path -Path $ResolvedReportPath -Parent
    if (-not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
    }

    $DoctorResult | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ResolvedReportPath -Encoding UTF8
}

function Show-WorkflowHelp {
    $helpText = @'
Windows App Reinstall Workflow

Actions:
  Doctor      Run prerequisite and workspace checks.
  Capture     Export installed programs to CSV or create a backup pack.
  Initialize  Build catalog/apps.json from installed-programs.csv.
  Status      Refresh and show catalog status counts.
  Prepare     Sync, classify, resolve packages, search installers, stage installers, and build manual-source guidance.
  Plan        Generate an install plan from the current install queue.
  Execute     Execute the current install queue.

Recommended workflow:
  pwsh -File .\scripts\AppReinstall.ps1 -Action Doctor
  pwsh -File .\scripts\AppReinstall.ps1 -Action Prepare
  pwsh -File .\scripts\AppReinstall.ps1 -Action Plan
  pwsh -File .\scripts\AppReinstall.ps1 -Action Execute -InteractiveChecklist

Examples:
  pwsh -File .\scripts\AppReinstall.ps1 -Action Capture
  pwsh -File .\scripts\AppReinstall.ps1 -Action Prepare -SearchRoot C:\Installers,$env:USERPROFILE\Downloads
  pwsh -File .\scripts\AppReinstall.ps1 -Action Prepare -SkipWingetDownload
    pwsh -File .\scripts\AppReinstall.ps1 -Action Plan -PlanPath .\output\install-plan.json
  pwsh -File .\scripts\AppReinstall.ps1 -Action Doctor | Format-List
'@

    Write-Host $helpText
}

$resolvedCsvPath = Resolve-WorkflowPath -Path $CsvPath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'installed-programs.csv')
$resolvedCatalogPath = Resolve-WorkflowPath -Path $CatalogPath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json')
$resolvedQueuePath = Resolve-WorkflowPath -Path $QueuePath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'output\install-queue.json')
$resolvedStageDirectory = Resolve-WorkflowPath -Path $StageDirectory -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'staged-installers')
$resolvedPlanPath = Resolve-WorkflowPath -Path $PlanPath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'output\install-plan.json')
$resolvedLogPath = Resolve-WorkflowPath -Path $LogPath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'output\install-log.json')
$resolvedDoctorReportPath = Resolve-WorkflowPath -Path $DoctorReportPath -DefaultPath (Join-Path -Path $projectRoot -ChildPath 'output\preflight-report.json')
$downloadWithWinget = -not $SkipWingetDownload.IsPresent

switch ($Action) {
    'Help' {
        Show-WorkflowHelp
    }

    'Doctor' {
        $doctor = Test-WorkflowPrerequisites -ResolvedCsvPath $resolvedCsvPath -ResolvedCatalogPath $resolvedCatalogPath -ResolvedQueuePath $resolvedQueuePath -ResolvedStageDirectory $resolvedStageDirectory -ResolvedPlanPath $resolvedPlanPath -ResolvedExecutionLogPath $resolvedLogPath
        Write-DoctorReport -DoctorResult $doctor -ResolvedReportPath $resolvedDoctorReportPath
        $doctor
    }

    'Capture' {
        if ($CaptureMode -eq 'BackupPack') {
            $backupListParams = @{}
            if (-not [string]::IsNullOrWhiteSpace($OutputRoot)) {
                $backupListParams.OutputRoot = (Resolve-AppPath -Path $OutputRoot)
            }

            $backupListResult = Invoke-WorkflowScript -ScriptName 'backup-app-list.ps1' -StepName 'Create backup pack' -Parameters $backupListParams
            $settingsBackupResult = $null

            if ($IncludeSettingsBackup.IsPresent) {
                $settingsParams = @{}
                if (-not [string]::IsNullOrWhiteSpace($OutputRoot)) {
                    $settingsParams.OutputRoot = (Resolve-AppPath -Path $OutputRoot)
                }

                $settingsBackupResult = Invoke-WorkflowScript -ScriptName 'backup-app-settings.ps1' -StepName 'Back up selected app settings' -Parameters $settingsParams
            }

            [PSCustomObject]@{
                action = 'capture'
                captureMode = $CaptureMode
                appListBackup = $backupListResult
                settingsBackup = $settingsBackupResult
            }
            break
        }

        $captureParams = @{
            Format = 'Csv'
            OutputPath = $resolvedCsvPath
        }

        $captureResult = @(Invoke-WorkflowScript -ScriptName 'Get-InstalledPrograms.ps1' -StepName 'Export installed programs to CSV' -Parameters $captureParams)
        [PSCustomObject]@{
            action = 'capture'
            captureMode = $CaptureMode
            csvPath = $resolvedCsvPath
            programCount = $captureResult.Count
        }
    }

    'Initialize' {
        if (-not (Test-Path -LiteralPath $resolvedCsvPath)) {
            throw "Installed program export not found: $resolvedCsvPath"
        }

        $initializeParams = @{
            CsvPath = $resolvedCsvPath
            CatalogPath = $resolvedCatalogPath
        }
        if ($Force.IsPresent) {
            $initializeParams.Force = $true
        }

        Invoke-WorkflowScript -ScriptName 'Initialize-AppCatalog.ps1' -StepName 'Initialize catalog' -Parameters $initializeParams
    }

    'Status' {
        if (-not (Test-Path -LiteralPath $resolvedCatalogPath)) {
            throw "Catalog not found: $resolvedCatalogPath"
        }

        $statusParams = @{
            CatalogPath = $resolvedCatalogPath
            View = 'Summary'
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\sync-report.json')
        }

        Invoke-WorkflowScript -ScriptName 'Sync-AppCatalog.ps1' -StepName 'Refresh catalog status summary' -Parameters $statusParams
    }

    'Prepare' {
        $doctor = Test-WorkflowPrerequisites -ResolvedCsvPath $resolvedCsvPath -ResolvedCatalogPath $resolvedCatalogPath -ResolvedQueuePath $resolvedQueuePath -ResolvedStageDirectory $resolvedStageDirectory -ResolvedPlanPath $resolvedPlanPath -ResolvedExecutionLogPath $resolvedLogPath
        Write-DoctorReport -DoctorResult $doctor -ResolvedReportPath $resolvedDoctorReportPath

        if (-not $doctor.readiness.prepare) {
            throw 'Prepare prerequisites are not satisfied. Run -Action Doctor and review output/preflight-report.json.'
        }

        if (-not $doctor.catalog.valid) {
            $initializeParams = @{
                CsvPath = $resolvedCsvPath
                CatalogPath = $resolvedCatalogPath
            }
            if ($Force.IsPresent) {
                $initializeParams.Force = $true
            }

            $null = Invoke-WorkflowScript -ScriptName 'Initialize-AppCatalog.ps1' -StepName 'Initialize catalog' -Parameters $initializeParams
        }

        $syncResult = Invoke-WorkflowScript -ScriptName 'Sync-AppCatalog.ps1' -StepName 'Sync catalog against current machine' -Parameters @{
            CatalogPath = $resolvedCatalogPath
            View = 'Summary'
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\sync-report.json')
        }

        $classifyParams = @{
            CatalogPath = $resolvedCatalogPath
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\classification-report.json')
        }
        if (-not $SkipIgnoreRecommendations.IsPresent) {
            $classifyParams.ApplyIgnoreRecommendations = $true
        }
        $classificationResult = Invoke-WorkflowScript -ScriptName 'Classify-AppCatalog.ps1' -StepName 'Classify apps and apply ignore recommendations' -Parameters $classifyParams

        $effectiveResolveMaxApps = Get-EffectiveMaxApps -ResolvedCatalogPath $resolvedCatalogPath -RequestedMaxApps $MaxApps -IncludeInstalledApps:$IncludeInstalled.IsPresent
        $resolveParams = @{
            CatalogPath = $resolvedCatalogPath
            MaxApps = $effectiveResolveMaxApps
            CandidatesPerApp = $CandidatesPerApp
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\winget-report.json')
        }
        if ($IncludeInstalled.IsPresent) {
            $resolveParams.IncludeInstalled = $true
        }
        if ($Force.IsPresent) {
            $resolveParams.Force = $true
        }
        $wingetResult = Invoke-WorkflowScript -ScriptName 'Resolve-WingetPackages.ps1' -StepName 'Resolve winget packages' -Parameters $resolveParams

        $effectiveLatestMaxApps = Get-EffectiveMaxApps -ResolvedCatalogPath $resolvedCatalogPath -RequestedMaxApps $MaxApps -ForLatestVersionRefresh -IncludeInstalledApps:$IncludeInstalled.IsPresent
        $latestParams = @{
            CatalogPath = $resolvedCatalogPath
            MaxApps = $effectiveLatestMaxApps
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\latest-version-report.json')
        }
        if ($IncludeInstalled.IsPresent) {
            $latestParams.IncludeInstalled = $true
        }
        if ($Force.IsPresent) {
            $latestParams.Force = $true
        }
        $latestVersionResult = Invoke-WorkflowScript -ScriptName 'Update-LatestAppVersions.ps1' -StepName 'Refresh latest available versions' -Parameters $latestParams

        $findInstallerParams = @{
            CatalogPath = $resolvedCatalogPath
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\installer-report.json')
        }
        if ($IncludeInstalled.IsPresent) {
            $findInstallerParams.IncludeInstalled = $true
        }
        if ($null -ne $SearchRoot -and $SearchRoot.Count -gt 0) {
            $findInstallerParams.SearchRoot = $SearchRoot
        }
        $installerSearchResult = Invoke-WorkflowScript -ScriptName 'Find-AppInstallers.ps1' -StepName 'Search for local installers' -Parameters $findInstallerParams

        $prepareInstallerParams = @{
            CatalogPath = $resolvedCatalogPath
            StageDirectory = $resolvedStageDirectory
            QueuePath = $resolvedQueuePath
        }
        if ($downloadWithWinget) {
            $prepareInstallerParams.DownloadWithWinget = $true
        }
        if ($DownloadFromManualReferences.IsPresent) {
            $prepareInstallerParams.DownloadFromManualReferences = $true
        }
        if ($IncludeInstalled.IsPresent) {
            $prepareInstallerParams.IncludeInstalled = $true
        }
        if ($Force.IsPresent) {
            $prepareInstallerParams.Force = $true
        }
        $queueResult = Invoke-WorkflowScript -ScriptName 'Prepare-AppInstallers.ps1' -StepName 'Stage installers and build install queue' -Parameters $prepareInstallerParams

        $manualSourceResult = Invoke-WorkflowScript -ScriptName 'Get-ManualSourceQueue.ps1' -StepName 'Build manual-source queue' -Parameters @{
            CatalogPath = $resolvedCatalogPath
            ReportPath = (Join-Path -Path $projectRoot -ChildPath 'output\manual-source-queue.json')
            UpdateCatalog = $true
            Format = 'Json'
        }

        [PSCustomObject]@{
            action = 'prepare'
            catalogPath = $resolvedCatalogPath
            queuePath = $resolvedQueuePath
            stageDirectory = $resolvedStageDirectory
            downloadWithWinget = $downloadWithWinget
            manualReferenceDownloads = $DownloadFromManualReferences.IsPresent
            summary = [PSCustomObject]@{
                sync = $syncResult
                classification = [PSCustomObject]@{
                    totalApps = $classificationResult.totalApps
                    ignoredApps = $classificationResult.ignoredApps
                }
                wingetResolved = @($wingetResult.resolved).Count
                latestVersionRefreshCount = @($latestVersionResult.refreshed).Count
                installerSearch = [PSCustomObject]@{
                    searchedRoots = $installerSearchResult.searchedRoots
                    appsWithInstallers = $installerSearchResult.appsWithInstallers
                }
                installQueueItems = @($queueResult.items).Count
                readyInstallers = @($queueResult.items | Where-Object { $_.ready }).Count
                manualApps = $manualSourceResult.totalManualApps
            }
            nextAction = 'Run AppReinstall.ps1 -Action Plan after reviewing the generated reports.'
        }
    }

    'Plan' {
        $doctor = Test-WorkflowPrerequisites -ResolvedCsvPath $resolvedCsvPath -ResolvedCatalogPath $resolvedCatalogPath -ResolvedQueuePath $resolvedQueuePath -ResolvedStageDirectory $resolvedStageDirectory -ResolvedPlanPath $resolvedPlanPath -ResolvedExecutionLogPath $resolvedLogPath
        Write-DoctorReport -DoctorResult $doctor -ResolvedReportPath $resolvedDoctorReportPath

        if (-not $doctor.readiness.plan) {
            throw 'Plan prerequisites are not satisfied. Run -Action Doctor and review output/preflight-report.json.'
        }

        $planParams = @{
            CatalogPath = $resolvedCatalogPath
            QueuePath = $resolvedQueuePath
            Mode = 'Plan'
            LogPath = $resolvedPlanPath
        }
        if ($IncludeInstalled.IsPresent) {
            $planParams.IncludeInstalled = $true
        }
        if ($UseWingetWhenAvailable.IsPresent) {
            $planParams.UseWingetWhenAvailable = $true
        }
        if ($AllowExeWithoutArgs.IsPresent) {
            $planParams.AllowExeWithoutArgs = $true
        }

        $planResult = Invoke-WorkflowScript -ScriptName 'Install-PreparedApps.ps1' -StepName 'Build install plan' -Parameters $planParams
        [PSCustomObject]@{
            action = 'plan'
            planPath = $resolvedPlanPath
            plannedItems = @($planResult.items).Count
            readyItems = @($planResult.items | Where-Object { $_.ready }).Count
            skippedItems = @($planResult.items | Where-Object { $_.executionState -eq 'skipped' }).Count
        }
    }

    'Execute' {
        $doctor = Test-WorkflowPrerequisites -ResolvedCsvPath $resolvedCsvPath -ResolvedCatalogPath $resolvedCatalogPath -ResolvedQueuePath $resolvedQueuePath -ResolvedStageDirectory $resolvedStageDirectory -ResolvedPlanPath $resolvedPlanPath -ResolvedExecutionLogPath $resolvedLogPath
        Write-DoctorReport -DoctorResult $doctor -ResolvedReportPath $resolvedDoctorReportPath

        if (-not $doctor.readiness.execute) {
            throw 'Execute prerequisites are not satisfied. Run -Action Doctor and review output/preflight-report.json.'
        }

        $executeParams = @{
            CatalogPath = $resolvedCatalogPath
            QueuePath = $resolvedQueuePath
            Mode = 'Execute'
            LogPath = $resolvedLogPath
        }
        if ($IncludeInstalled.IsPresent) {
            $executeParams.IncludeInstalled = $true
        }
        if ($UseWingetWhenAvailable.IsPresent) {
            $executeParams.UseWingetWhenAvailable = $true
        }
        if ($AllowExeWithoutArgs.IsPresent) {
            $executeParams.AllowExeWithoutArgs = $true
        }
        if ($InteractiveChecklist.IsPresent) {
            $executeParams.InteractiveChecklist = $true
        }

        $executeResult = Invoke-WorkflowScript -ScriptName 'Install-PreparedApps.ps1' -StepName 'Execute install queue' -Parameters $executeParams
        [PSCustomObject]@{
            action = 'execute'
            logPath = $resolvedLogPath
            succeeded = @($executeResult.items | Where-Object { $_.executionState -eq 'succeeded' }).Count
            failed = @($executeResult.items | Where-Object { $_.executionState -eq 'failed' }).Count
            skipped = @($executeResult.items | Where-Object { $_.executionState -eq 'skipped' }).Count
        }
    }
}