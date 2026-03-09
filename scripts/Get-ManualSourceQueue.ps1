[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string]$ReportPath,

    [switch]$UpdateCatalog,

    [ValidateSet('Table', 'Json')]
    [string]$Format = 'Table'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path -Path $projectRoot -ChildPath 'output\manual-source-queue.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

function Get-ManualAcquisitionRecommendation {
    param(
        [Parameter(Mandatory)]
        $App
    )

    $name = Normalize-AppText -Value ([string]$App.name)
    $publisher = Normalize-AppText -Value ([string]$App.publisher)

    if ($name.Contains('adobe svg viewer')) {
        return [PSCustomObject]@{
            acquisitionType = 'legacy-archive'
            sourceHint = 'No current Adobe download; use retained media or internal archive only'
            reason = 'Adobe SVG Viewer is discontinued and Adobe no longer provides a current download'
            referenceUrl = 'https://community.adobe.com/t5/download-install-discussions/download-adobe-svg-viewer-3/m-p/4573449'
        }
    }

    if ($name.Contains('application verifier')) {
        return [PSCustomObject]@{
            acquisitionType = 'microsoft-sdk-component'
            sourceHint = 'Windows SDK or Visual Studio installer components'
            reason = 'Developer verification component rather than a standalone desktop app'
            referenceUrl = 'https://learn.microsoft.com/en-us/windows-hardware/drivers/devtest/application-verifier'
        }
    }

    if ($name.Contains('iis 10 0 express') -or $name.Contains('iis express')) {
        return [PSCustomObject]@{
            acquisitionType = 'microsoft-dev-component'
            sourceHint = 'Official Microsoft IIS Express download or Visual Studio installer'
            reason = 'Microsoft developer component not currently available through winget in this environment'
            referenceUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=48264'
        }
    }

    if ($name.Contains('sql server 2019 localdb') -or $name.Contains('localdb')) {
        return [PSCustomObject]@{
            acquisitionType = 'sql-server-installer'
            sourceHint = 'Official SQL Server Express download with LocalDB included'
            reason = 'LocalDB is typically distributed as part of SQL Server tooling rather than a separate winget package'
            referenceUrl = 'https://www.microsoft.com/en-us/download/details.aspx?id=101064'
        }
    }

    if ($name.Contains('realtek') -or $publisher.Contains('realtek')) {
        return [PSCustomObject]@{
            acquisitionType = 'oem-driver-site'
            sourceHint = 'Device OEM support page first, then Realtek driver listings if needed'
            reason = 'Hardware driver package should come from the OEM or chipset vendor'
            referenceUrl = 'https://www.realtek.com/Download/List?cate_id=590&menu_id=405'
        }
    }

    if ($name.Contains('messenger') -or $publisher.Contains('facebook')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-web'
            sourceHint = 'Meta Messenger web or Meta support page'
            reason = 'No safe desktop package source was confirmed; current acquisition may be web-first'
            referenceUrl = 'https://www.meta.com/en-gb/messenger/'
        }
    }

    if ($name.Contains('canon') -or $publisher.Contains('canon')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-support'
            sourceHint = 'Canon support and driver downloads'
            reason = 'Printer utility and driver packages are model-specific and should be sourced from the vendor'
            referenceUrl = 'https://www.canon-europe.com/supportproduct/tabcontent/?productTcmUri=tcm:13-1019644&type=drivers&language=en&os=all'
        }
    }

    if ($name.Contains('webasto') -or $publisher.Contains('webasto')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-support'
            sourceHint = 'Webasto service or support downloads'
            reason = 'Specialized OEM tooling with no safe package match found'
            referenceUrl = 'https://www.techwebasto.com/updates/tech-news/51-thermo-test-software-overview-v3-8.html'
        }
    }

    if ($name.Contains('forscan')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-site'
            sourceHint = 'Official FORScan downloads'
            reason = 'Niche automotive tool with no current winget package found'
            referenceUrl = 'https://forscan.org/download.html'
        }
    }

    if ($name.Contains('ptstroubleshooter') -or $publisher.Contains('ford')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-portal'
            sourceHint = 'Ford Tech Service dealer download portal'
            reason = 'Specialized vendor support tool outside normal public package feeds'
            referenceUrl = 'https://www.fordtechservice.dealerconnection.com/Rotunda/MCSIDSDownloadSoftware'
        }
    }

    if ($name.Contains('obdwiz')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-site'
            sourceHint = 'Official OBDLink or ScanTool downloads'
            reason = 'Vendor-specific automotive application with no package feed match'
            referenceUrl = 'https://www.obdlink.com/software/'
        }
    }

    if ($name.Contains('final draft')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-site'
            sourceHint = 'Official Final Draft downloads and installer guides'
            reason = 'Commercial software typically distributed through vendor-managed downloads'
            referenceUrl = 'https://www.finaldraft.com/download/'
        }
    }

    if ($name.Contains('calman') -or $publisher.Contains('portrait displays')) {
        return [PSCustomObject]@{
            acquisitionType = 'vendor-site'
            sourceHint = 'Portrait Displays product and download pages'
            reason = 'Licensed calibration software with no safe package feed match'
            referenceUrl = 'https://www.portrait.com/trial-downloads/'
        }
    }

    return [PSCustomObject]@{
        acquisitionType = 'manual-review'
        sourceHint = 'Manual vendor or internal archive lookup required'
        reason = 'No package source or local installer candidate found'
        referenceUrl = $null
    }
}

function Clear-ManualAcquisitionMetadata {
    param(
        [Parameter(Mandatory)]
        $App
    )

    $App.installer.manualAcquisitionType = $null
    $App.installer.manualSourceHint = $null
    $App.installer.manualReferenceUrl = $null
    $App.installer.manualReason = $null
    $App.installer.manualUpdatedAt = $null
}

function Set-ManualAcquisitionMetadata {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        $Recommendation,

        [Parameter(Mandatory)]
        [string]$UpdatedAt
    )

    $App.installer.manualAcquisitionType = [string]$Recommendation.acquisitionType
    $App.installer.manualSourceHint = [string]$Recommendation.sourceHint
    $App.installer.manualReferenceUrl = $Recommendation.referenceUrl
    $App.installer.manualReason = [string]$Recommendation.reason
    $App.installer.manualUpdatedAt = $UpdatedAt
}

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)

$generatedAt = (Get-Date).ToString('o')

$items = @(
    $apps |
        ForEach-Object {
            $app = $_
            $needsManualSource = (
                $app.desired -eq $true -and
                $app.status -eq 'missing' -and
                [string]::IsNullOrWhiteSpace([string]$app.installer.wingetId) -and
                [string]::IsNullOrWhiteSpace([string]$app.installer.localPath)
            )

            if (-not $needsManualSource) {
                if ($UpdateCatalog) {
                    Clear-ManualAcquisitionMetadata -App $app
                }

                return
            }

            $recommendation = Get-ManualAcquisitionRecommendation -App $app
            if ($UpdateCatalog) {
                Set-ManualAcquisitionMetadata -App $app -Recommendation $recommendation -UpdatedAt $generatedAt
            }

            [PSCustomObject]@{
                name = [string]$app.name
                publisher = [string]$app.publisher
                classification = [string]$app.classification.bucket
                acquisitionType = [string]$recommendation.acquisitionType
                sourceHint = [string]$recommendation.sourceHint
                reason = [string]$recommendation.reason
                referenceUrl = $recommendation.referenceUrl
            }
        } |
        Sort-Object acquisitionType, name
)

$report = [PSCustomObject]@{
    generatedAt = $generatedAt
    totalManualApps = $items.Count
    items = $items
}

if ($UpdateCatalog) {
    $catalog.generatedAt = $generatedAt
    Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null
}

$resolvedReportPath = Resolve-AppPath -Path $ReportPath
$reportDirectory = Split-Path -Path $resolvedReportPath -Parent
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedReportPath -Encoding UTF8

if ($Format -eq 'Json') {
    $report
}
else {
    $items | Format-Table -AutoSize
}