[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string]$QueuePath,

    [string[]]$Name,

    [ValidateSet('Plan', 'Execute')]
    [string]$Mode = 'Plan',

    [switch]$IncludeInstalled,

    [switch]$UseWingetWhenAvailable,

    [switch]$AllowExeWithoutArgs,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($QueuePath)) {
    $QueuePath = Join-Path -Path $projectRoot -ChildPath 'output\install-queue.json'
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path $projectRoot -ChildPath 'output\install-log.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$queueDocument = Get-Content -LiteralPath (Resolve-AppPath -Path $QueuePath) -Raw | ConvertFrom-Json
$queueItems = @($queueDocument.items)
$normalizedNames = @{}
foreach ($itemName in @($Name)) {
    if (-not [string]::IsNullOrWhiteSpace($itemName)) {
        $normalizedNames[(Normalize-AppText -Value $itemName)] = $true
    }
}

$workItems = foreach ($app in @($catalog.apps)) {
    if ($normalizedNames.Count -gt 0) {
        $candidateName = Normalize-AppText -Value ([string]$app.name)
        if (-not $normalizedNames.ContainsKey($candidateName)) {
            continue
        }
    }

    if ($app.desired -eq $false -and -not $IncludeInstalled.IsPresent) {
        continue
    }

    if ($app.status -eq 'installed' -and -not $IncludeInstalled.IsPresent) {
        continue
    }

    $queueItem = @($queueItems | Where-Object { [string]$_.name -eq [string]$app.name }) | Select-Object -First 1
    if ($null -eq $queueItem) {
        continue
    }

    $plannedCommand = $null
    $plannedArgs = @()
    $method = 'unresolved'
    $ready = $false
    $details = $null

    if ($UseWingetWhenAvailable.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$app.installer.wingetId)) {
        $method = 'winget-install'
        $ready = $true
        $plannedCommand = 'winget.exe'
        $plannedArgs = @(
            'install',
            '--id', [string]$app.installer.wingetId,
            '--exact',
            '--source', 'winget',
            '--accept-package-agreements',
            '--accept-source-agreements',
            '--disable-interactivity'
        )
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$queueItem.stagedPath) -and (Test-Path -LiteralPath $queueItem.stagedPath)) {
        $extension = [System.IO.Path]::GetExtension([string]$queueItem.stagedPath).ToLowerInvariant()
        switch ($extension) {
            '.msi' {
                $method = 'msi'
                $ready = $true
                $plannedCommand = 'msiexec.exe'
                $plannedArgs = @('/i', [string]$queueItem.stagedPath, '/qn', '/norestart')
            }
            '.exe' {
                $method = 'exe'
                $plannedCommand = [string]$queueItem.stagedPath
                if (-not [string]::IsNullOrWhiteSpace([string]$app.installer.installArgs)) {
                    $ready = $true
                    $plannedArgs = @([string]$app.installer.installArgs)
                }
                elseif ($AllowExeWithoutArgs.IsPresent) {
                    $ready = $true
                    $plannedArgs = @()
                    $details = 'executing exe without explicit silent arguments'
                }
                else {
                    $details = 'exe installer requires installer.installArgs or -AllowExeWithoutArgs'
                }
            }
            default {
                $method = $extension.TrimStart('.')
                $details = 'unsupported installer extension for automated execution'
            }
        }
    }
    else {
        $details = [string]$queueItem.details
    }

    $executionState = 'planned'
    $exitCode = $null

    if ($Mode -eq 'Execute' -and $ready) {
        & $plannedCommand @plannedArgs
        $exitCode = $LASTEXITCODE
        if ($null -eq $exitCode) {
            $exitCode = 0
        }

        if ($exitCode -eq 0) {
            $executionState = 'succeeded'
            $app.status = 'installed'
            $app.detection.lastCheckedAt = (Get-Date).ToString('o')
            $app.detection.lastSeenName = [string]$app.name
            $installedVersion = Convert-EmptyToNull -Value ([string]$app.latest.version)
            if ($null -eq $installedVersion) {
                $installedVersion = Convert-EmptyToNull -Value ([string]$app.expectedVersion)
            }
            $app.detection.lastSeenVersion = $installedVersion
        }
        else {
            $executionState = 'failed'
        }
    }
    elseif (-not $ready) {
        $executionState = 'skipped'
    }

    [PSCustomObject]@{
        name = [string]$app.name
        method = $method
        ready = $ready
        mode = $Mode
        executionState = $executionState
        command = $plannedCommand
        arguments = @($plannedArgs)
        exitCode = $exitCode
        details = $details
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$logDocument = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    mode = $Mode
    items = @($workItems)
}

$resolvedLogPath = Resolve-AppPath -Path $LogPath
$logDirectory = Split-Path -Path $resolvedLogPath -Parent
if (-not (Test-Path -LiteralPath $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

$logDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedLogPath -Encoding UTF8
$logDocument