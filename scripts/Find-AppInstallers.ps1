[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string[]]$SearchRoot,

    [int]$Limit = 5,

    [switch]$IncludeInstalled,

    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ($null -eq $SearchRoot -or $SearchRoot.Count -eq 0) {
    $SearchRoot = @(
        (Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads'),
        (Join-Path -Path $env:USERPROFILE -ChildPath 'Desktop'),
        (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents'),
        'C:\Installers',
        'C:\Software'
    )
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)
$targetApps = if ($IncludeInstalled.IsPresent) {
    $apps
}
else {
    @($apps | Where-Object { $_.status -ne 'installed' })
}

$installerFiles = @(Get-InstallerFileInventory -SearchRoots $SearchRoot)
$discoveredAt = (Get-Date).ToString('o')

foreach ($app in $targetApps) {
    $candidates = @(Get-AppInstallerCandidates -App $app -InstallerFiles $installerFiles -Limit $Limit)
    $app.installer.localCandidates = @($candidates)
    $app.installer.discoveredAt = $discoveredAt

    if ($candidates.Count -gt 0) {
        $app.installer.localPath = [string]$candidates[0].path
        $app.installer.localType = [string]$candidates[0].extension
    }
    else {
        $app.installer.localPath = $null
        $app.installer.localType = $null
    }

    $app.installer.ready = [bool](
        (-not [string]::IsNullOrWhiteSpace([string]$app.installer.localPath) -and (Test-Path -LiteralPath $app.installer.localPath)) -or
        (-not [string]::IsNullOrWhiteSpace([string]$app.installer.downloadedPath) -and (Test-Path -LiteralPath $app.installer.downloadedPath))
    )
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$report = [PSCustomObject]@{
    searchedRoots = $SearchRoot
    scannedInstallerFiles = $installerFiles.Count
    appsChecked = $targetApps.Count
    appsWithInstallers = @($targetApps | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.installer.localPath) }).Count
    apps = @(
        $targetApps |
            Select-Object name, status, @{ Name = 'localPath'; Expression = { $_.installer.localPath } }, @{ Name = 'ready'; Expression = { $_.installer.ready } }
    )
}

if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $resolvedReportPath = Resolve-AppPath -Path $ReportPath
    $reportDirectory = Split-Path -Path $resolvedReportPath -Parent
    if (-not (Test-Path -LiteralPath $reportDirectory)) {
        New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
    }

    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8
}

$report