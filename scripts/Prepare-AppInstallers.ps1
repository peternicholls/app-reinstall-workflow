[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string]$StageDirectory,

    [string]$QueuePath,

    [string[]]$Name,

    [switch]$DownloadWithWinget,

    [switch]$DownloadFromManualReferences,

    [switch]$IncludeInstalled,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($StageDirectory)) {
    $StageDirectory = Join-Path -Path $projectRoot -ChildPath 'staged-installers'
}

if ([string]::IsNullOrWhiteSpace($QueuePath)) {
    $QueuePath = Join-Path -Path $projectRoot -ChildPath 'output\install-queue.json'
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$supportedInstallerExtensions = Get-DefaultInstallerExtensions

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$apps = @($catalog.apps)
$targetApps = if ($IncludeInstalled.IsPresent) {
    $apps
}
else {
    @($apps | Where-Object { $_.status -ne 'installed' -and $_.desired -ne $false })
}

$nameFilters = @(
    $Name |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
)

if ($nameFilters.Count -gt 0) {
    $targetApps = @(
        $targetApps |
            Where-Object {
                $appName = [string]$_.name
                @($nameFilters | Where-Object { $appName -like $_ }).Count -gt 0
            }
    )
}

$resolvedStageDirectory = Resolve-AppPath -Path $StageDirectory
$resolvedQueuePath = Resolve-AppPath -Path $QueuePath

if (-not (Test-Path -LiteralPath $resolvedStageDirectory)) {
    New-Item -Path $resolvedStageDirectory -ItemType Directory -Force | Out-Null
}

$queue = foreach ($app in $targetApps) {
    $stagedPath = $null
    $source = 'unresolved'
    $ready = $false
    $details = $null

    if (-not [string]::IsNullOrWhiteSpace([string]$app.installer.localPath) -and (Test-Path -LiteralPath $app.installer.localPath)) {
        $fileName = Split-Path -Path $app.installer.localPath -Leaf
        $safePrefix = (Normalize-AppText -Value ([string]$app.name)) -replace ' ', '-'
        if ([string]::IsNullOrWhiteSpace($safePrefix)) {
            $safePrefix = 'app'
        }

        $destinationPath = Join-Path -Path $resolvedStageDirectory -ChildPath ($safePrefix + '-' + $fileName)
        if ($Force.IsPresent -or -not (Test-Path -LiteralPath $destinationPath)) {
            Copy-Item -LiteralPath $app.installer.localPath -Destination $destinationPath -Force:$Force.IsPresent
        }

        $stagedPath = $destinationPath
        $app.installer.downloadedPath = $destinationPath
        $source = 'local'
        $ready = $true
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$app.installer.downloadedPath) -and (Test-Path -LiteralPath $app.installer.downloadedPath)) {
        $stagedPath = [string]$app.installer.downloadedPath
        $source = if ([string]::IsNullOrWhiteSpace([string]$app.installer.preferredSource)) { 'downloaded' } else { [string]$app.installer.preferredSource }
        $ready = $true
    }
    elseif ($DownloadWithWinget.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$app.installer.wingetId)) {
        $appStageDirectory = Join-Path -Path $resolvedStageDirectory -ChildPath ((Normalize-AppText -Value ([string]$app.name)) -replace ' ', '-')
        if (-not (Test-Path -LiteralPath $appStageDirectory)) {
            New-Item -Path $appStageDirectory -ItemType Directory -Force | Out-Null
        }

        $downloadedFile = @(
            Get-ChildItem -LiteralPath $appStageDirectory -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $supportedInstallerExtensions -contains $_.Extension.ToLowerInvariant() } |
                Sort-Object LastWriteTime -Descending
        ) | Select-Object -First 1

        if ($null -ne $downloadedFile -and -not $Force.IsPresent) {
            $stagedPath = $downloadedFile.FullName
            $app.installer.downloadedPath = $downloadedFile.FullName
            $source = 'winget'
            $ready = $true
        }
        else {
            $wingetArgs = @(
                'download',
                '--id', [string]$app.installer.wingetId,
                '--exact',
                '--source', 'winget',
                '--download-directory', $appStageDirectory,
                '--accept-package-agreements',
                '--accept-source-agreements',
                '--disable-interactivity'
            )

            $wingetOutput = @(& winget.exe @wingetArgs 2>&1)
            $wingetExitCode = $LASTEXITCODE

            $downloadedFile = @(
                Get-ChildItem -LiteralPath $appStageDirectory -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { $supportedInstallerExtensions -contains $_.Extension.ToLowerInvariant() } |
                    Sort-Object LastWriteTime -Descending
            ) | Select-Object -First 1

            if ($null -ne $downloadedFile) {
                $stagedPath = $downloadedFile.FullName
                $app.installer.downloadedPath = $downloadedFile.FullName
                $source = 'winget'
                $ready = $true
            }
            else {
                if ($wingetExitCode -ne 0) {
                    $details = "winget download failed: $(([string]::Join(' ', ($wingetOutput | Select-Object -Last 3))).Trim())"
                }
                else {
                    $details = "winget download completed without a detectable installer file: $(([string]::Join(' ', ($wingetOutput | Select-Object -Last 3))).Trim())"
                }

                $remainingItems = @(Get-ChildItem -LiteralPath $appStageDirectory -Force -ErrorAction SilentlyContinue)
                if ($remainingItems.Count -eq 0) {
                    Remove-Item -LiteralPath $appStageDirectory -Force -ErrorAction SilentlyContinue
                }

                if ($DownloadFromManualReferences.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$app.installer.manualReferenceUrl)) {
                    $manualResult = Stage-ManualReferenceInstaller -App $app -StageDirectory $resolvedStageDirectory -Force:$Force.IsPresent
                    if ($manualResult.status -eq 'downloaded') {
                        $stagedPath = $manualResult.stagedPath
                        $source = 'manual-url'
                        $ready = $true
                        $details = $null
                    }
                    elseif ([string]::IsNullOrWhiteSpace($details)) {
                        $details = [string]$manualResult.details
                    }
                }
            }
        }
    }
    elseif ($DownloadFromManualReferences.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$app.installer.manualReferenceUrl)) {
        $manualResult = Stage-ManualReferenceInstaller -App $app -StageDirectory $resolvedStageDirectory -Force:$Force.IsPresent
        $stagedPath = $manualResult.stagedPath
        $source = if ($manualResult.status -eq 'downloaded') { 'manual-url' } else { 'unresolved' }
        $ready = ($manualResult.status -eq 'downloaded')
        $details = [string]$manualResult.details
    }
    else {
        if ([string]::IsNullOrWhiteSpace([string]$app.installer.localPath) -and [string]::IsNullOrWhiteSpace([string]$app.installer.wingetId)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$app.installer.manualReferenceUrl)) {
                $details = 'manual reference URL exists but -DownloadFromManualReferences was not supplied'
            }
            else {
                $details = 'no local installer or winget id recorded'
            }
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$app.installer.localPath)) {
            $details = 'local installer path is recorded but the file is missing'
        }
        else {
            if (-not [string]::IsNullOrWhiteSpace([string]$app.installer.manualReferenceUrl)) {
                $details = 'winget id exists but -DownloadWithWinget was not supplied; manual reference URL also available'
            }
            else {
                $details = 'winget id exists but -DownloadWithWinget was not supplied'
            }
        }
    }

    $app.installer.ready = $ready

    [PSCustomObject]@{
        name = [string]$app.name
        status = [string]$app.status
        source = $source
        ready = $ready
        stagedPath = $stagedPath
        details = $details
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$queueDirectory = Split-Path -Path $resolvedQueuePath -Parent
if (-not (Test-Path -LiteralPath $queueDirectory)) {
    New-Item -Path $queueDirectory -ItemType Directory -Force | Out-Null
}

$queueDocument = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    stageDirectory = $resolvedStageDirectory
    items = @($queue)
}

$queueDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedQueuePath -Encoding UTF8
$queueDocument