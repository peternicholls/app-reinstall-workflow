[CmdletBinding()]
param(
    [string]$CsvPath,
    [string]$CatalogPath,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $CsvPath = Join-Path -Path $projectRoot -ChildPath 'installed-programs.csv'
}

if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$resolvedCsvPath = Resolve-AppPath -Path $CsvPath
$resolvedCatalogPath = Resolve-AppPath -Path $CatalogPath

if ((Test-Path -LiteralPath $resolvedCatalogPath) -and -not $Force.IsPresent) {
    throw "Catalog already exists: $resolvedCatalogPath. Use -Force to overwrite it."
}

$rows = @(
    Import-Csv -LiteralPath $resolvedCsvPath |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.DisplayName) }
)

$apps = foreach ($group in ($rows | Group-Object { '{0}|{1}' -f (Normalize-AppText -Value $_.DisplayName), (Normalize-AppText -Value $_.Publisher) })) {
    $preferredRow = @(
        $group.Group |
            Sort-Object DisplayVersion -Descending
    )[0]

    [PSCustomObject]@{
        name = [string]$preferredRow.DisplayName
        publisher = Convert-EmptyToNull -Value ([string]$preferredRow.Publisher)
        expectedVersion = Convert-EmptyToNull -Value ([string]$preferredRow.DisplayVersion)
        desired = $true
        status = 'unknown'
        detection = [PSCustomObject]@{
            matchNames = @(
                $group.Group.DisplayName |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                    Sort-Object -Unique
            )
            lastCheckedAt = $null
            lastSeenName = $null
            lastSeenVersion = $null
        }
        latest = [PSCustomObject]@{
            version = $null
            source = $null
            checkedAt = $null
            packageId = $null
        }
        classification = [PSCustomObject]@{
            bucket = $null
            recommendedAction = $null
            reason = $null
        }
        installer = [PSCustomObject]@{
            localPath = $null
            localType = $null
            localCandidates = @()
            discoveredAt = $null
            wingetId = $null
            wingetCheckedAt = $null
            wingetCandidates = @()
            downloadedPath = $null
            installArgs = $null
            preferredSource = $null
            manualAcquisitionType = $null
            manualSourceHint = $null
            manualReferenceUrl = $null
            manualReason = $null
            manualUpdatedAt = $null
            ready = $false
        }
        notes = $null
    }
}

$catalog = [PSCustomObject]@{
    schemaVersion = 1
    generatedAt = (Get-Date).ToString('o')
    source = [PSCustomObject]@{
        csvPath = $resolvedCsvPath
    }
    apps = @(
        $apps |
            Sort-Object name, publisher
    )
}

$savedPath = Save-AppCatalog -Catalog $catalog -CatalogPath $resolvedCatalogPath

[PSCustomObject]@{
    catalogPath = $savedPath
    appCount = @($catalog.apps).Count
    generatedAt = $catalog.generatedAt
}