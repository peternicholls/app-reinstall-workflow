[CmdletBinding()]
param(
    [string]$CatalogPath,

    [string]$QueuePath,

    [string[]]$Name,

    [ValidateSet('Plan', 'Execute')]
    [string]$Mode = 'Plan',

    [switch]$IncludeInstalled,

    [switch]$UseWingetWhenAvailable,

    [switch]$AllowExeWithoutArgs,

    [switch]$InteractiveChecklist,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
    $CatalogPath = Join-Path -Path $projectRoot -ChildPath 'catalog\apps.json'
}

if ([string]::IsNullOrWhiteSpace($QueuePath)) {
    $QueuePath = Join-Path -Path $projectRoot -ChildPath 'output\install-queue.json'
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $defaultOutputName = if ($Mode -eq 'Plan') { 'output\install-plan.json' } else { 'output\install-log.json' }
    $LogPath = Join-Path -Path $projectRoot -ChildPath $defaultOutputName
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

function Invoke-InstallChecklistSelection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Apps
    )

    $orderedApps = @($Apps | Sort-Object -Property name)
    if ($orderedApps.Count -eq 0) {
        return @()
    }

    $selectedByName = @{}
    foreach ($entry in $orderedApps) {
        $selectedByName[[string]$entry.name] = $true
    }

    $cursorIndex = 0
    $topIndex = 0

    while ($true) {
        $windowHeight = 25
        try {
            if ($null -ne $Host.UI.RawUI -and $null -ne $Host.UI.RawUI.WindowSize) {
                $windowHeight = [int]$Host.UI.RawUI.WindowSize.Height
            }
        }
        catch {
            $windowHeight = 25
        }

        $visibleRows = [Math]::Max(5, $windowHeight - 8)
        if ($cursorIndex -lt $topIndex) {
            $topIndex = $cursorIndex
        }
        if ($cursorIndex -ge ($topIndex + $visibleRows)) {
            $topIndex = $cursorIndex - $visibleRows + 1
        }

        Clear-Host
        Write-Host 'Install checklist (all selected by default)'
        Write-Host 'Use Up/Down to move, Space to toggle, A=all, N=none, D/Enter=done, C/Esc=cancel.'
        Write-Host ''

        $endIndex = [Math]::Min($orderedApps.Count - 1, $topIndex + $visibleRows - 1)
        for ($index = $topIndex; $index -le $endIndex; $index++) {
            $app = $orderedApps[$index]
            $marker = if ($selectedByName[[string]$app.name]) { 'x' } else { ' ' }
            $pointer = if ($index -eq $cursorIndex) { '>' } else { ' ' }
            $position = $index + 1
            Write-Host ("{0} [{1}] {2,2}. {3}" -f $pointer, $marker, $position, [string]$app.name)
        }

        if ($topIndex -gt 0) {
            Write-Host '... (more above)'
        }
        if ($endIndex -lt ($orderedApps.Count - 1)) {
            Write-Host '... (more below)'
        }

        $selectedCount = @($orderedApps | Where-Object { $selectedByName[[string]$_.name] }).Count
        Write-Host ''
        Write-Host ("Selected: {0}/{1}" -f $selectedCount, $orderedApps.Count)

        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        $virtualKeyCode = [int]$key.VirtualKeyCode
        $charCode = [int][char]$key.Character
        $isDone = $false
        $isCancelled = $false

        switch ($virtualKeyCode) {
            38 {
                if ($cursorIndex -gt 0) {
                    $cursorIndex--
                }
            }
            40 {
                if ($cursorIndex -lt ($orderedApps.Count - 1)) {
                    $cursorIndex++
                }
            }
            32 {
                $currentName = [string]$orderedApps[$cursorIndex].name
                $selectedByName[$currentName] = -not $selectedByName[$currentName]
            }
            13 {
                $isDone = $true
            }
            27 {
                $isCancelled = $true
            }
            65 {
                foreach ($entry in $orderedApps) {
                    $selectedByName[[string]$entry.name] = $true
                }
            }
            78 {
                foreach ($entry in $orderedApps) {
                    $selectedByName[[string]$entry.name] = $false
                }
            }
            68 {
                $isDone = $true
            }
            67 {
                $isCancelled = $true
            }
            default {
                if ($charCode -ne 0) {
                    switch (([char]$charCode).ToString().ToLowerInvariant()) {
                        'a' {
                            foreach ($entry in $orderedApps) {
                                $selectedByName[[string]$entry.name] = $true
                            }
                        }
                        'n' {
                            foreach ($entry in $orderedApps) {
                                $selectedByName[[string]$entry.name] = $false
                            }
                        }
                        'd' {
                            $isDone = $true
                        }
                        'c' {
                            $isCancelled = $true
                        }
                    }
                }
            }
        }

        if ($isCancelled) {
            return @()
        }

        if ($isDone) {
            break
        }
    }

    return @(
        $orderedApps |
            Where-Object { $selectedByName[[string]$_.name] } |
            ForEach-Object { [string]$_.name }
    )
}

$catalog = Read-AppCatalog -CatalogPath $CatalogPath
$queueDocument = Get-Content -LiteralPath (Resolve-AppPath -Path $QueuePath) -Raw | ConvertFrom-Json
$queueItems = @($queueDocument.items)
$normalizedNames = @{}
foreach ($itemName in @($Name)) {
    if (-not [string]::IsNullOrWhiteSpace($itemName)) {
        $normalizedNames[(Normalize-AppText -Value $itemName)] = $true
    }
}

$candidateApps = foreach ($app in @($catalog.apps)) {
    if ($normalizedNames.Count -gt 0) {
        $candidateName = Normalize-AppText -Value ([string]$app.name)
        if (-not $normalizedNames.ContainsKey($candidateName)) {
            continue
        }
    }

    if ($app.desired -eq $false -and -not $IncludeInstalled.IsPresent) {
        continue
    }

    if ($app.status -eq 'installed' -and -not $IncludeInstalled.IsPresent) {
        continue
    }

    $app
}

if ($InteractiveChecklist.IsPresent) {
    $selectedNames = @(Invoke-InstallChecklistSelection -Apps $candidateApps)
    if ($selectedNames.Count -eq 0) {
        Write-Warning 'No apps selected in the checklist. Nothing will be processed.'
        $candidateApps = @()
    }
    else {
        $selectedNormalizedNames = @{}
        foreach ($selectedName in $selectedNames) {
            $selectedNormalizedNames[(Normalize-AppText -Value $selectedName)] = $true
        }

        $candidateApps = @(
            $candidateApps |
                Where-Object {
                    $selectedNormalizedNames.ContainsKey(
                        (Normalize-AppText -Value ([string]$_.name))
                    )
                }
        )
    }
}

$workItems = foreach ($app in @($candidateApps)) {

    $queueItem = @($queueItems | Where-Object { [string]$_.name -eq [string]$app.name }) | Select-Object -First 1
    if ($null -eq $queueItem) {
        continue
    }

    $plannedCommand = $null
    $plannedArgs = @()
    $method = 'unresolved'
    $ready = $false
    $details = $null

    if ($UseWingetWhenAvailable.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$app.installer.wingetId)) {
        $method = 'winget-install'
        $ready = $true
        $plannedCommand = 'winget.exe'
        $plannedArgs = @(
            'install',
            '--id', [string]$app.installer.wingetId,
            '--exact',
            '--source', 'winget',
            '--accept-package-agreements',
            '--accept-source-agreements',
            '--disable-interactivity'
        )
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$queueItem.stagedPath) -and (Test-Path -LiteralPath $queueItem.stagedPath)) {
        $extension = [System.IO.Path]::GetExtension([string]$queueItem.stagedPath).ToLowerInvariant()
        switch ($extension) {
            '.msi' {
                $method = 'msi'
                $ready = $true
                $plannedCommand = 'msiexec.exe'
                $plannedArgs = @('/i', [string]$queueItem.stagedPath, '/qn', '/norestart')
            }
            '.exe' {
                $method = 'exe'
                $plannedCommand = [string]$queueItem.stagedPath
                if (-not [string]::IsNullOrWhiteSpace([string]$app.installer.installArgs)) {
                    $ready = $true
                    $plannedArgs = @([string]$app.installer.installArgs)
                }
                elseif ($AllowExeWithoutArgs.IsPresent) {
                    $ready = $true
                    $plannedArgs = @()
                    $details = 'executing exe without explicit silent arguments'
                }
                else {
                    $details = 'exe installer requires installer.installArgs or -AllowExeWithoutArgs'
                }
            }
            default {
                $method = $extension.TrimStart('.')
                $details = 'unsupported installer extension for automated execution'
            }
        }
    }
    else {
        $details = [string]$queueItem.details
    }

    $executionState = 'planned'
    $exitCode = $null

    if ($Mode -eq 'Execute' -and $ready) {
        if ($plannedArgs.Count -gt 0) {
            $process = Start-Process -FilePath $plannedCommand -ArgumentList $plannedArgs -Wait -PassThru
        }
        else {
            $process = Start-Process -FilePath $plannedCommand -Wait -PassThru
        }

        $exitCode = [int]$process.ExitCode

        if ($exitCode -eq 0) {
            $executionState = 'succeeded'
            $app.status = 'installed'
            $app.detection.lastCheckedAt = (Get-Date).ToString('o')
            $app.detection.lastSeenName = [string]$app.name
            $installedVersion = Convert-EmptyToNull -Value ([string]$app.latest.version)
            if ($null -eq $installedVersion) {
                $installedVersion = Convert-EmptyToNull -Value ([string]$app.expectedVersion)
            }
            $app.detection.lastSeenVersion = $installedVersion
        }
        else {
            $executionState = 'failed'
        }
    }
    elseif (-not $ready) {
        $executionState = 'skipped'
    }

    [PSCustomObject]@{
        name = [string]$app.name
        source = [string]$queueItem.source
        stagedPath = [string]$queueItem.stagedPath
        method = $method
        ready = $ready
        mode = $Mode
        executionState = $executionState
        command = $plannedCommand
        arguments = @($plannedArgs)
        exitCode = $exitCode
        details = $details
    }
}

Save-AppCatalog -Catalog $catalog -CatalogPath $CatalogPath | Out-Null

$documentType = if ($Mode -eq 'Plan') { 'install-plan' } else { 'install-log' }
$outputDocument = [PSCustomObject]@{
    generatedAt = (Get-Date).ToString('o')
    documentType = $documentType
    mode = $Mode
    catalogPath = (Resolve-AppPath -Path $CatalogPath)
    queuePath = (Resolve-AppPath -Path $QueuePath)
    summary = [PSCustomObject]@{
        totalItems = @($workItems).Count
        readyItems = @($workItems | Where-Object { $_.ready }).Count
        plannedItems = @($workItems | Where-Object { $_.executionState -eq 'planned' }).Count
        succeededItems = @($workItems | Where-Object { $_.executionState -eq 'succeeded' }).Count
        failedItems = @($workItems | Where-Object { $_.executionState -eq 'failed' }).Count
        skippedItems = @($workItems | Where-Object { $_.executionState -eq 'skipped' }).Count
    }
    items = @($workItems)
}

$resolvedLogPath = Resolve-AppPath -Path $LogPath
$logDirectory = Split-Path -Path $resolvedLogPath -Parent
if (-not (Test-Path -LiteralPath $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

$outputDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedLogPath -Encoding UTF8
$outputDocument