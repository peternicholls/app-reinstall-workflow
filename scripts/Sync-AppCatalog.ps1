[CmdletBinding()]
param(
    [string]$CatalogPath,

    [ValidateSet('Summary', 'All', 'Installed', 'Missing', 'Ignored')]
    [string]$View = 'Summary',

    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$installedPrograms = Get-InstalledProgramInventory
$catalog = Update-AppCatalogStatuses -Catalog $catalog -InstalledPrograms $installedPrograms
$resolvedCatalogPath = Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath

$apps = @($catalog.apps)
$summary = [PSCustomObject]@{
    catalogPath = $resolvedCatalogPath
    checkedAt = (Get-Date).ToString('o')
    total = $apps.Count
    installed = @($apps | Where-Object { $_.status -eq 'installed' }).Count
    missing = @($apps | Where-Object { $_.status -eq 'missing' }).Count
    ignored = @($apps | Where-Object { $_.status -eq 'ignored' }).Count
}

if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $report = [PSCustomObject]@{
        summary = $summary
        apps = $apps
    }
    $resolvedReportPath = Resolve-AppPath -Path $ReportPath
    $reportDirectory = Split-Path -Path $resolvedReportPath -Parent
    if (-not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
    }

    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
}

switch ($View) {
    'All' { $apps }
    'Installed' { $apps | Where-Object { $_.status -eq 'installed' } }
    'Missing' { $apps | Where-Object { $_.status -eq 'missing' } }
    'Ignored' { $apps | Where-Object { $_.status -eq 'ignored' } }
    default { $summary }
}