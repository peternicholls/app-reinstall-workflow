[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Name,

    [ValidateSet('unknown', 'installed', 'missing', 'ignored')]
    [string]$Status,

    [string]$CatalogPath,

    [switch]$Exact,

    [switch]$ClearDetection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$normalizedTarget = Normalize-AppText -Value $Name
$apps = @($catalog.apps)

$matchingApps = @(
    $apps |
        Where-Object {
            $candidate = Normalize-AppText -Value ([string]$_.name)
            if ($Exact.IsPresent) {
                return $candidate -eq $normalizedTarget
            }

            return $candidate -eq $normalizedTarget -or $candidate.Contains($normalizedTarget) -or $normalizedTarget.Contains($candidate)
        }
)

if ($matchingApps.Count -eq 0) {
    throw "No app found matching '$Name'."
}

if ($matchingApps.Count -gt 1) {
    $candidateNames = $matchingApps | Select-Object -ExpandProperty name
    throw "More than one app matched '$Name': $($candidateNames -join ', '). Use -Exact with the full name."
}

$app = $matchingApps[0]
$app.status = $Status
$app.detection.lastCheckedAt = (Get-Date).ToString('o')

if ($Status -eq 'installed') {
    $app.detection.lastSeenName = [string]$app.name
    $app.detection.lastSeenVersion = Convert-EmptyToNull -Value ([string]$app.expectedVersion)
}
elseif ($ClearDetection.IsPresent) {
    $app.detection.lastSeenName = $null
    $app.detection.lastSeenVersion = $null
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null
$app