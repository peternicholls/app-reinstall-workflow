[CmdletBinding()]
param(
    [string]$CatalogPath,

    [switch]$ApplyIgnoreRecommendations,

    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path -Path $projectRoot -ChildPath 'output\classification-report.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)

foreach ($app in $apps) {
    Initialize-AppEntry -App $app
    $classification = Get-AppClassification -App $app
    $app.classification.bucket = [string]$classification.bucket
    $app.classification.recommendedAction = [string]$classification.recommendedAction
    $app.classification.reason = [string]$classification.reason

    if ($ApplyIgnoreRecommendations.IsPresent -and [string]$classification.recommendedAction -eq 'ignore') {
        $app.desired = $false
        if ($app.status -ne 'installed') {
            $app.status = 'ignored'
        }
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$bucketSummary = @(
    $apps |
        Group-Object { [string]$_.classification.bucket } |
        Sort-Object Name |
        ForEach-Object {
            [PSCustomObject]@{
                bucket = $_.Name
                count = $_.Count
            }
        }
)

$report = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    applyIgnoreRecommendations = $ApplyIgnoreRecommendations.IsPresent
    totalApps = $apps.Count
    ignoredApps = @($apps | Where-Object { $_.desired -eq $false }).Count
    bucketSummary = $bucketSummary
    apps = @(
        $apps |
            Select-Object name, status, desired, @{ Name = 'bucket'; Expression = { $_.classification.bucket } }, @{ Name = 'recommendedAction'; Expression = { $_.classification.recommendedAction } }, @{ Name = 'reason'; Expression = { $_.classification.reason } }
    )
}

$resolvedReportPath = Resolve-AppPath -Path $ReportPath
$reportDirectory = Split-Path -Path $resolvedReportPath -Parent
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
$report