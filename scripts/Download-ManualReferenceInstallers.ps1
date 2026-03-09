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

$supportedExtensions = @('.exe', '.msi', '.msix', '.msixbundle', '.appx', '.appxbundle', '.zip')
$requestTimeoutSeconds = 120

function Get-LinkHrefValue {
    param(
        [Parameter(Mandatory)]
        $Link
    )

    foreach ($propertyName in 'href', 'Href') {
        if ($Link.PSObject.Properties[$propertyName]) {
            return [string]$Link.$propertyName
        }
    }

    return ''
}

function Get-LinkTextValue {
    param(
        [Parameter(Mandatory)]
        $Link
    )

    foreach ($propertyName in 'innerText', 'InnerText', 'outerText', 'OuterText') {
        if ($Link.PSObject.Properties[$propertyName]) {
            return [string]$Link.$propertyName
        }
    }

    return ''
}

function Get-DirectInstallerUrl {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        [string]$ReferenceUrl
    )

    $referenceUri = [Uri]$ReferenceUrl
    $referenceLower = $ReferenceUrl.ToLowerInvariant()
    if (@($supportedExtensions | Where-Object { $referenceLower.Contains($_) }).Count -gt 0) {
        return [PSCustomObject]@{
            downloadUrl = $ReferenceUrl
            sourcePage = $ReferenceUrl
            detection = 'direct-url'
        }
    }

    $response = Invoke-WebRequest -Uri $ReferenceUrl -MaximumRedirection 5 -TimeoutSec $requestTimeoutSeconds
    $appTokens = @(
        (Get-VersionlessAppText -Value ([string]$App.name)).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) |
            Where-Object { $_.Length -ge 3 }
    )

    $candidates = foreach ($link in @($response.Links)) {
        $href = Get-LinkHrefValue -Link $link
        if ([string]::IsNullOrWhiteSpace($href)) {
            continue
        }

        try {
            $candidateUri = [Uri]::new($referenceUri, $href)
        }
        catch {
            continue
        }

        $candidateUrl = $candidateUri.AbsoluteUri
        $candidateLower = $candidateUrl.ToLowerInvariant()
        $extension = $null
        foreach ($supportedExtension in $supportedExtensions) {
            if ($candidateLower -match [regex]::Escape($supportedExtension) + '([?#].*)?$') {
                $extension = $supportedExtension
                break
            }
        }

        if ($null -eq $extension) {
            continue
        }

        $score = 0
        if ($candidateUri.Host -eq $referenceUri.Host) {
            $score += 50
        }
        elseif ($candidateUri.Host.EndsWith('.' + $referenceUri.Host)) {
            $score += 25
        }

        if ($candidateUri.Scheme -eq 'https') {
            $score += 10
        }

        if ($candidateLower.Contains('download')) {
            $score += 20
        }

        foreach ($token in $appTokens) {
            if ($candidateLower.Contains($token)) {
                $score += 15
            }
        }

        $linkText = Normalize-AppText -Value (Get-LinkTextValue -Link $link)
        if ($linkText.Contains('download')) {
            $score += 10
        }

        [PSCustomObject]@{
            url = $candidateUrl
            score = $score
        }
    }

    $bestCandidate = @(
        $candidates |
            Sort-Object -Property @(
                @{ Expression = 'score'; Descending = $true },
                'url'
            ) |
            Select-Object -First 1
    )
    if ($bestCandidate.Count -eq 0) {
        return $null
    }

    return [PSCustomObject]@{
        downloadUrl = [string]$bestCandidate[0].url
        sourcePage = $ReferenceUrl
        detection = 'page-link'
    }
}

function Get-TargetFileName {
    param(
        [Parameter(Mandatory)]
        [string]$DownloadUrl,

        [Parameter(Mandatory)]
        $App
    )

    $uri = [Uri]$DownloadUrl
    $fileName = [System.IO.Path]::GetFileName($uri.AbsolutePath)
    if (-not [string]::IsNullOrWhiteSpace($fileName)) {
        return $fileName
    }

    $safePrefix = (Normalize-AppText -Value ([string]$App.name)) -replace ' ', '-'
    if ([string]::IsNullOrWhiteSpace($safePrefix)) {
        $safePrefix = 'app'
    }

    return $safePrefix + '.bin'
}

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
    $stagedPath = $null
    $status = 'unresolved'
    $details = $null
    $downloadUrl = $null
    $detection = $null

    try {
        $resolution = Get-DirectInstallerUrl -App $app -ReferenceUrl ([string]$app.installer.manualReferenceUrl)
        if ($null -eq $resolution) {
            $status = 'no-direct-link'
            $details = 'reference page did not expose a direct installer link'
        }
        else {
            $downloadUrl = [string]$resolution.downloadUrl
            $detection = [string]$resolution.detection
            $fileName = Get-TargetFileName -DownloadUrl $downloadUrl -App $app

            $appStageDirectory = Join-Path -Path $resolvedStageDirectory -ChildPath ((Normalize-AppText -Value ([string]$app.name)) -replace ' ', '-')
            if (-not (Test-Path -LiteralPath $appStageDirectory)) {
                New-Item -Path $appStageDirectory -ItemType Directory -Force | Out-Null
            }

            $destinationPath = Join-Path -Path $appStageDirectory -ChildPath $fileName
            if ($Force.IsPresent -or -not (Test-Path -LiteralPath $destinationPath)) {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath -MaximumRedirection 5 -TimeoutSec $requestTimeoutSeconds
            }

            if (Test-Path -LiteralPath $destinationPath) {
                $status = 'downloaded'
                $stagedPath = $destinationPath
                $app.installer.downloadedPath = $destinationPath
                $app.installer.preferredSource = 'manual-url'
                $app.installer.ready = $true
            }
            else {
                $status = 'download-failed'
                $details = 'download completed without a staged file'
            }
        }
    }
    catch {
        $status = 'error'
        $details = $_.Exception.Message
    }

    [PSCustomObject]@{
        name = [string]$app.name
        manualReferenceUrl = [string]$app.installer.manualReferenceUrl
        downloadUrl = $downloadUrl
        detection = $detection
        status = $status
        stagedPath = $stagedPath
        details = $details
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