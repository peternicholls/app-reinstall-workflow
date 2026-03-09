[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string]$StageDirectory,

    [string]$ReportPath,

    [string[]]$Name,

    [int]$MaxApps = 0,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($StageDirectory)) {
    $StageDirectory = Join-Path -Path $projectRoot -ChildPath 'output\staged-installers'
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path -Path $projectRoot -ChildPath 'output\manual-download-queue.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)

$resolvedStageDirectory = Resolve-AppPath -Path $StageDirectory
$resolvedReportPath = Resolve-AppPath -Path $ReportPath
if (-not (Test-Path -LiteralPath $resolvedStageDirectory)) {
    New-Item -Path $resolvedStageDirectory -ItemType Directory -Force | Out-Null
}

$selectedApps = @(
    $apps |
        Where-Object {
            $_.desired -eq $true -and
            $_.status -eq 'missing' -and
            -not [string]::IsNullOrWhiteSpace([string]$_.installer.manualReferenceUrl)
        }
)

$nameFilters = @(
    $Name |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
)

if ($nameFilters.Count -gt 0) {
    $selectedApps = @(
        $selectedApps |
            Where-Object {
                $appName = [string]$_.name
                @($nameFilters | Where-Object { $appName -like $_ }).Count -gt 0
            }
    )
}

if ($MaxApps -gt 0) {
    $selectedApps = @($selectedApps | Select-Object -First $MaxApps)
}

$reportItems = foreach ($app in $selectedApps) {
    $result = $null
    try {
        $result = Stage-ManualReferenceInstaller -App $app -StageDirectory $resolvedStageDirectory -Force:$Force.IsPresent
    }
    catch {
        $result = [PSCustomObject]@{
            status = 'error'
            stagedPath = $null
            downloadUrl = $null
            detection = $null
            details = $_.Exception.Message
        }
    }

    [PSCustomObject]@{
        name = [string]$app.name
        manualReferenceUrl = [string]$app.installer.manualReferenceUrl
        downloadUrl = $result.downloadUrl
        detection = $result.detection
        status = $result.status
        stagedPath = $result.stagedPath
        details = $result.details
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$reportDirectory = Split-Path -Path $resolvedReportPath -Parent
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    stageDirectory = $resolvedStageDirectory
    items = @($reportItems)
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
$report