[CmdletBinding()]
param(
    [string]$CatalogPath,

    [int]$MaxApps = 25,

    [int]$CandidatesPerApp = 5,

    [switch]$IncludeInstalled,

    [switch]$Force,

    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path -Path $projectRoot -ChildPath 'output\winget-report.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)
$targetApps = if ($IncludeInstalled.IsPresent) {
    $apps
}
else {
    @($apps | Where-Object { $_.status -ne 'installed' -and $_.desired -ne $false })
}

if (-not $Force.IsPresent) {
    $targetApps = @(
        $targetApps |
            Where-Object {
                [string]::IsNullOrWhiteSpace([string]$_.installer.wingetId) -and
                [string]::IsNullOrWhiteSpace([string]$_.installer.wingetCheckedAt)
            }
    )
}

$targetApps = @(
    $targetApps |
        Sort-Object @{ Expression = {
            if ([string]::IsNullOrWhiteSpace([string]$_.installer.wingetCheckedAt)) {
                return [datetime]::MinValue
            }

            return [datetime][string]$_.installer.wingetCheckedAt
        } }, name |
        Select-Object -First $MaxApps
)
$resolved = foreach ($app in $targetApps) {
    $candidates = @(Get-WingetPackageCandidates -App $app -Count $CandidatesPerApp)
    $checkedAt = (Get-Date).ToString('o')
    $app.installer.wingetCandidates = @($candidates)
    $app.installer.wingetCheckedAt = $checkedAt
    $app.latest.checkedAt = $checkedAt
    $app.latest.source = 'winget'

    if ($candidates.Count -gt 0) {
        $app.installer.wingetId = [string]$candidates[0].id
        $app.latest.version = [string]$candidates[0].version
        $app.latest.packageId = [string]$candidates[0].id
    }
    elseif ($Force.IsPresent) {
        $app.installer.wingetId = $null
        $app.latest.version = $null
        $app.latest.packageId = $null
    }
    else {
        $app.latest.version = $null
        $app.latest.packageId = $null
    }

    [PSCustomObject]@{
        name = [string]$app.name
        status = [string]$app.status
        expectedVersion = [string]$app.expectedVersion
        latestVersion = [string]$app.latest.version
        wingetId = [string]$app.installer.wingetId
        candidateCount = @($candidates).Count
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$report = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    maxApps = $MaxApps
    candidatesPerApp = $CandidatesPerApp
    resolved = @($resolved)
}

$resolvedReportPath = Resolve-AppPath -Path $ReportPath
$reportDirectory = Split-Path -Path $resolvedReportPath -Parent
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
$report