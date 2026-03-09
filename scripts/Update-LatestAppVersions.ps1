[CmdletBinding()]
param(
    [string]$CatalogPath,

    [int]$MaxApps = 25,

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
    $ReportPath = Join-Path -Path $projectRoot -ChildPath 'output\latest-version-report.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)
$targetApps = @(
    $apps |
        Where-Object {
            (-not [string]::IsNullOrWhiteSpace([string]$_.installer.wingetId)) -and
            ($IncludeInstalled.IsPresent -or $_.status -ne 'installed')
        }
)

if (-not $Force.IsPresent) {
    $targetApps = @(
        $targetApps |
            Where-Object { [string]::IsNullOrWhiteSpace([string]$_.latest.checkedAt) }
    )
}

$targetApps = @($targetApps | Select-Object -First $MaxApps)

$results = foreach ($app in $targetApps) {
    $checkedAt = (Get-Date).ToString('o')
    $package = Get-WingetPackageById -Id ([string]$app.installer.wingetId)

    $app.latest.checkedAt = $checkedAt
    $app.latest.source = 'winget'
    $app.latest.packageId = [string]$app.installer.wingetId

    if ($null -ne $package) {
        $app.latest.version = [string]$package.Version
    }
    elseif ($Force.IsPresent) {
        $app.latest.version = $null
    }

    [PSCustomObject]@{
        name = [string]$app.name
        expectedVersion = [string]$app.expectedVersion
        latestVersion = [string]$app.latest.version
        wingetId = [string]$app.installer.wingetId
        checkedAt = $checkedAt
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$report = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    maxApps = $MaxApps
    refreshed = @($results)
}

$resolvedReportPath = Resolve-AppPath -Path $ReportPath
$reportDirectory = Split-Path -Path $resolvedReportPath -Parent
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
$report