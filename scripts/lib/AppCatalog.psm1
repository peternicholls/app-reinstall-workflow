Set-StrictMode -Version Latest

function Resolve-AppPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location) -ChildPath $Path))
}

function Convert-EmptyToNull {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    return $Value.Trim()
}

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) {
        return ''
    }

    return [string]$property.Value
}

function Test-HasProperty {
    param(
        [Parameter(Mandatory)]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    return $null -ne $InputObject.PSObject.Properties[$PropertyName]
}

function Normalize-AppText {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $normalized = $Value.ToLowerInvariant()
    $normalized = $normalized -replace '[^a-z0-9]+', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

function Get-AppClassification {
    param(
        [Parameter(Mandatory)]
        $App
    )

    $name = [string]$App.name
    $publisher = [string]$App.publisher
    $normalizedName = Normalize-AppText -Value $name
    $normalizedPublisher = Normalize-AppText -Value $publisher

    $result = [ordered]@{
        bucket = 'application'
        recommendedAction = 'keep'
        reason = 'default application classification'
    }

    $runtimePatterns = @(
        'visual c',
        'redistributable',
        'minimum runtime',
        'additional runtime',
        'net host',
        'net runtime',
        'host fx resolver',
        'dotnet host',
        'dotnet runtime',
        'dotnet host fx resolver',
        'asp net core',
        'desktop runtime',
        'standard targeting pack',
        'coreruntime',
        'webview2 runtime'
    )

    $sdkPatterns = @(
        'windows sdk',
        'software development kit',
        'universal crt',
        'winrt intellisense',
        'net native sdk',
        'extension sdk',
        'sdk arm',
        'sdk addon',
        'app certification kit',
        'application compatibility database',
        'kits configuration installer',
        'winappdeploy',
        'msi development tools',
        'setup wmi provider',
        'supportedapilist',
        'coreeditorfonts'
    )

    $systemPatterns = @(
        'msxml',
        'update health tools',
        'windows subsystem for linux update',
        'click to run licensing component',
        'click to run extensibility component',
        'bing service',
        'kb5001716'
    )

    $driverPatterns = @(
        'driver',
        'graphics',
        'card reader',
        'realtek',
        'intel processor graphics'
    )

    $developerPatterns = @(
        'visual studio code',
        'powershell',
        'phpstorm',
        'iis express',
        'web deploy',
        'openjdk',
        'localdb',
        'sql server',
        'application verifier'
    )

    if ($normalizedName -match 'chrome|edge|firefox|earth pro') {
        $result.bucket = 'browser'
        $result.reason = 'browser-like package name'
    }
    elseif (@($driverPatterns | Where-Object { $normalizedName.Contains($_) }).Count -gt 0) {
        $result.bucket = 'driver'
        $result.reason = 'driver-related package name'
    }
    elseif ($normalizedName.Contains('application compatibility database') -or $normalizedName.Contains('setup wmi provider')) {
        $result.bucket = 'sdk'
        $result.recommendedAction = 'ignore'
        $result.reason = 'developer support component'
    }
    elseif ($normalizedName -eq 'microsoft 365 en us' -or ($normalizedName.Contains('microsoft 365') -and $normalizedName.Contains('en us'))) {
        $result.bucket = 'system-component'
        $result.recommendedAction = 'ignore'
        $result.reason = 'language or bundled office component'
    }
    elseif (@($developerPatterns | Where-Object { $normalizedName.Contains($_) }).Count -gt 0) {
        $result.bucket = 'developer-tool'
        $result.reason = 'developer-tool package name'
    }
    elseif (@($runtimePatterns | Where-Object { $normalizedName.Contains($_) }).Count -gt 0) {
        $result.bucket = 'runtime'
        $result.recommendedAction = 'ignore'
        $result.reason = 'runtime or redistributable component'
    }
    elseif (@($sdkPatterns | Where-Object { $normalizedName.Contains($_) }).Count -gt 0) {
        $result.bucket = 'sdk'
        $result.recommendedAction = 'ignore'
        $result.reason = 'sdk or developer support component'
    }
    elseif (@($systemPatterns | Where-Object { $normalizedName.Contains($_) }).Count -gt 0) {
        $result.bucket = 'system-component'
        $result.recommendedAction = 'ignore'
        $result.reason = 'system-managed or bundled component'
    }
    elseif ($normalizedPublisher -eq 'microsoft corporation' -and $normalizedName -match 'windows|office 16|microsoft 365') {
        $result.bucket = 'system-component'
        $result.recommendedAction = 'review'
        $result.reason = 'microsoft system-related package'
    }

    return [PSCustomObject]$result
}

function Get-VersionlessAppText {
    param(
        [AllowNull()]
        [string]$Value
    )

    $normalized = Normalize-AppText -Value $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $tokens = @(
        $normalized.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) |
            Where-Object {
                $_ -notmatch '^\d+([._-]\d+)+([a-z]+)?$' -and
                $_ -notmatch '^\d+[a-z]+$' -and
                $_ -notin @('x64', 'x86', 'arm', 'arm64')
            }
    )

    return ($tokens -join ' ').Trim()
}

function Read-AppCatalog {
    param(
        [Parameter(Mandatory)]
        [string]$CatalogPath
    )

    $resolvedPath = Resolve-AppPath -Path $CatalogPath
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "Catalog file not found: $resolvedPath"
    }

    $catalog = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
    foreach ($app in @($catalog.apps)) {
        Initialize-AppEntry -App $app
    }

    return $catalog
}

function Save-AppCatalog {
    param(
        [Parameter(Mandatory)]
        $Catalog,

        [Parameter(Mandatory)]
        [string]$CatalogPath
    )

    $resolvedPath = Resolve-AppPath -Path $CatalogPath
    $directoryPath = Split-Path -Path $resolvedPath -Parent
    if (-not (Test-Path -LiteralPath $directoryPath)) {
        New-Item -Path $directoryPath -ItemType Directory -Force | Out-Null
    }

    $Catalog | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $resolvedPath -Encoding UTF8
    return $resolvedPath
}

function New-ValidationIssue {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('error', 'warning')]
        [string]$Severity,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Message
    )

    return [PSCustomObject]@{
        severity = $Severity
        path = $Path
        message = $Message
    }
}

function Test-AppCatalogStructure {
    param(
        [Parameter(Mandatory)]
        [string]$CatalogPath
    )

    $resolvedPath = Resolve-AppPath -Path $CatalogPath
    $issues = New-Object 'System.Collections.Generic.List[object]'
    $catalog = $null
    $schemaVersion = $null
    $apps = @()

    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        $issues.Add((New-ValidationIssue -Severity 'error' -Path 'catalog' -Message "Catalog file not found: $resolvedPath"))
    }
    else {
        try {
            $catalog = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
        }
        catch {
            $issues.Add((New-ValidationIssue -Severity 'error' -Path 'catalog' -Message $_.Exception.Message))
        }
    }

    if ($null -ne $catalog) {
        if (Test-HasProperty -InputObject $catalog -PropertyName 'schemaVersion') {
            $schemaVersion = $catalog.schemaVersion
            if ([string]$schemaVersion -ne '1') {
                $issues.Add((New-ValidationIssue -Severity 'warning' -Path 'schemaVersion' -Message ("Unsupported schemaVersion value: {0}" -f [string]$schemaVersion)))
            }
        }
        else {
            $issues.Add((New-ValidationIssue -Severity 'error' -Path 'schemaVersion' -Message 'Catalog is missing schemaVersion'))
        }

        if (-not (Test-HasProperty -InputObject $catalog -PropertyName 'apps') -or $null -eq $catalog.apps) {
            $issues.Add((New-ValidationIssue -Severity 'error' -Path 'apps' -Message 'Catalog is missing apps'))
        }
        else {
            $apps = @($catalog.apps)
        }

        $allowedStatuses = @('unknown', 'installed', 'missing', 'ignored')
        for ($index = 0; $index -lt $apps.Count; $index++) {
            $app = $apps[$index]
            $appPath = "apps[$index]"

            if ([string]::IsNullOrWhiteSpace([string](Get-OptionalPropertyValue -InputObject $app -PropertyName 'name'))) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$appPath.name" -Message 'App entry is missing name'))
            }

            if (-not (Test-HasProperty -InputObject $app -PropertyName 'desired') -or -not ($app.desired -is [bool])) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$appPath.desired" -Message 'App entry is missing a boolean desired flag'))
            }

            if (-not (Test-HasProperty -InputObject $app -PropertyName 'status') -or [string]::IsNullOrWhiteSpace([string]$app.status)) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$appPath.status" -Message 'App entry is missing status'))
            }
            elseif ([string]$app.status -notin $allowedStatuses) {
                $issues.Add((New-ValidationIssue -Severity 'warning' -Path "$appPath.status" -Message ("Unexpected status value: {0}" -f [string]$app.status)))
            }

            foreach ($propertyName in 'detection', 'latest', 'classification', 'installer') {
                if (-not (Test-HasProperty -InputObject $app -PropertyName $propertyName) -or $null -eq $app.$propertyName) {
                    $issues.Add((New-ValidationIssue -Severity 'error' -Path ("{0}.{1}" -f $appPath, $propertyName) -Message ("App entry is missing {0}" -f $propertyName)))
                }
            }
        }
    }

    $errorCount = @($issues | Where-Object { $_.severity -eq 'error' }).Count
    $warningCount = @($issues | Where-Object { $_.severity -eq 'warning' }).Count

    return [PSCustomObject]@{
        type = 'catalog'
        path = $resolvedPath
        isValid = ($errorCount -eq 0)
        issueCount = $issues.Count
        errorCount = $errorCount
        warningCount = $warningCount
        issues = @($issues)
        summary = [PSCustomObject]@{
            schemaVersion = $schemaVersion
            appCount = $apps.Count
        }
    }
}

function Test-InstallQueueStructure {
    param(
        [Parameter(Mandatory)]
        [string]$QueuePath
    )

    $resolvedPath = Resolve-AppPath -Path $QueuePath
    $issues = New-Object 'System.Collections.Generic.List[object]'
    $queueDocument = $null
    $items = @()

    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        $issues.Add((New-ValidationIssue -Severity 'error' -Path 'queue' -Message "Install queue file not found: $resolvedPath"))
    }
    else {
        try {
            $queueDocument = Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json
        }
        catch {
            $issues.Add((New-ValidationIssue -Severity 'error' -Path 'queue' -Message $_.Exception.Message))
        }
    }

    if ($null -ne $queueDocument) {
        if (-not (Test-HasProperty -InputObject $queueDocument -PropertyName 'items') -or $null -eq $queueDocument.items) {
            $issues.Add((New-ValidationIssue -Severity 'error' -Path 'items' -Message 'Install queue is missing items'))
        }
        else {
            $items = @($queueDocument.items)
        }

        for ($index = 0; $index -lt $items.Count; $index++) {
            $item = $items[$index]
            $itemPath = "items[$index]"

            if ([string]::IsNullOrWhiteSpace([string](Get-OptionalPropertyValue -InputObject $item -PropertyName 'name'))) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$itemPath.name" -Message 'Queue item is missing name'))
            }

            if ([string]::IsNullOrWhiteSpace([string](Get-OptionalPropertyValue -InputObject $item -PropertyName 'source'))) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$itemPath.source" -Message 'Queue item is missing source'))
            }

            if (-not (Test-HasProperty -InputObject $item -PropertyName 'ready') -or -not ($item.ready -is [bool])) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$itemPath.ready" -Message 'Queue item is missing a boolean ready flag'))
            }

            if (-not (Test-HasProperty -InputObject $item -PropertyName 'stagedPath')) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$itemPath.stagedPath" -Message 'Queue item is missing stagedPath'))
            }
            elseif ($item.ready -is [bool] -and $item.ready -and [string]$item.source -ne 'winget' -and [string]::IsNullOrWhiteSpace([string]$item.stagedPath)) {
                $issues.Add((New-ValidationIssue -Severity 'error' -Path "$itemPath.stagedPath" -Message 'Ready queue item is missing stagedPath'))
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$item.stagedPath) -and -not (Test-Path -LiteralPath ([string]$item.stagedPath))) {
                $issues.Add((New-ValidationIssue -Severity 'warning' -Path "$itemPath.stagedPath" -Message ("Referenced staged installer does not exist: {0}" -f [string]$item.stagedPath)))
            }
        }
    }

    $readyItemCount = @($items | Where-Object { $_.ready -is [bool] -and $_.ready }).Count
    $errorCount = @($issues | Where-Object { $_.severity -eq 'error' }).Count
    $warningCount = @($issues | Where-Object { $_.severity -eq 'warning' }).Count

    return [PSCustomObject]@{
        type = 'install-queue'
        path = $resolvedPath
        isValid = ($errorCount -eq 0)
        issueCount = $issues.Count
        errorCount = $errorCount
        warningCount = $warningCount
        issues = @($issues)
        summary = [PSCustomObject]@{
            itemCount = $items.Count
            readyItemCount = $readyItemCount
        }
    }
}

function Initialize-AppEntry {
    param(
        [Parameter(Mandatory)]
        $App
    )

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'desired')) {
        $App | Add-Member -NotePropertyName desired -NotePropertyValue $true
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'status')) {
        $App | Add-Member -NotePropertyName status -NotePropertyValue 'unknown'
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'notes')) {
        $App | Add-Member -NotePropertyName notes -NotePropertyValue $null
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'latest') -or $null -eq $App.latest) {
        $App | Add-Member -NotePropertyName latest -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    foreach ($propertyName in 'version', 'source', 'checkedAt', 'packageId') {
        if (-not (Test-HasProperty -InputObject $App.latest -PropertyName $propertyName)) {
            $App.latest | Add-Member -NotePropertyName $propertyName -NotePropertyValue $null
        }
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'classification') -or $null -eq $App.classification) {
        $App | Add-Member -NotePropertyName classification -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    foreach ($propertyName in 'bucket', 'recommendedAction', 'reason') {
        if (-not (Test-HasProperty -InputObject $App.classification -PropertyName $propertyName)) {
            $App.classification | Add-Member -NotePropertyName $propertyName -NotePropertyValue $null
        }
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'detection') -or $null -eq $App.detection) {
        $App | Add-Member -NotePropertyName detection -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    if (-not (Test-HasProperty -InputObject $App.detection -PropertyName 'matchNames') -or $null -eq $App.detection.matchNames) {
        $matchNames = @()
        if (-not [string]::IsNullOrWhiteSpace([string]$App.name)) {
            $matchNames += [string]$App.name
        }
        $App.detection | Add-Member -NotePropertyName matchNames -NotePropertyValue @($matchNames | Sort-Object -Unique) -Force
    }

    foreach ($propertyName in 'lastCheckedAt', 'lastSeenName', 'lastSeenVersion') {
        if (-not (Test-HasProperty -InputObject $App.detection -PropertyName $propertyName)) {
            $App.detection | Add-Member -NotePropertyName $propertyName -NotePropertyValue $null
        }
    }

    if (-not (Test-HasProperty -InputObject $App -PropertyName 'installer') -or $null -eq $App.installer) {
        $App | Add-Member -NotePropertyName installer -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    foreach ($propertyName in 'localPath', 'localType', 'discoveredAt', 'wingetId', 'downloadedPath', 'wingetCheckedAt') {
        if (-not (Test-HasProperty -InputObject $App.installer -PropertyName $propertyName)) {
            $App.installer | Add-Member -NotePropertyName $propertyName -NotePropertyValue $null
        }
    }

    foreach ($propertyName in 'installArgs', 'preferredSource', 'manualAcquisitionType', 'manualSourceHint', 'manualReferenceUrl', 'manualReason', 'manualUpdatedAt') {
        if (-not (Test-HasProperty -InputObject $App.installer -PropertyName $propertyName)) {
            $App.installer | Add-Member -NotePropertyName $propertyName -NotePropertyValue $null
        }
    }

    if (-not (Test-HasProperty -InputObject $App.installer -PropertyName 'localCandidates') -or $null -eq $App.installer.localCandidates) {
        $App.installer | Add-Member -NotePropertyName localCandidates -NotePropertyValue @() -Force
    }

    if (-not (Test-HasProperty -InputObject $App.installer -PropertyName 'wingetCandidates') -or $null -eq $App.installer.wingetCandidates) {
        $App.installer | Add-Member -NotePropertyName wingetCandidates -NotePropertyValue @() -Force
    }

    if (-not (Test-HasProperty -InputObject $App.installer -PropertyName 'ready')) {
        $App.installer | Add-Member -NotePropertyName ready -NotePropertyValue $false
    }
}

function Get-InstalledProgramInventory {
    $registryTargets = @(
        [PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'machine'; Architecture = 'x64' },
        [PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'machine'; Architecture = 'x86' },
        [PSCustomObject]@{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'user'; Architecture = 'unknown' }
    )

    $results = foreach ($target in $registryTargets) {
        Get-ItemProperty -Path $target.Path -ErrorAction SilentlyContinue |
            Where-Object { -not [string]::IsNullOrWhiteSpace((Get-OptionalPropertyValue -InputObject $_ -PropertyName 'DisplayName')) } |
            ForEach-Object {
                [PSCustomObject]@{
                    DisplayName = Get-OptionalPropertyValue -InputObject $_ -PropertyName 'DisplayName'
                    DisplayVersion = Get-OptionalPropertyValue -InputObject $_ -PropertyName 'DisplayVersion'
                    Publisher = Get-OptionalPropertyValue -InputObject $_ -PropertyName 'Publisher'
                    InstallLocation = Get-OptionalPropertyValue -InputObject $_ -PropertyName 'InstallLocation'
                    UninstallString = Get-OptionalPropertyValue -InputObject $_ -PropertyName 'UninstallString'
                    Scope = $target.Scope
                    Architecture = $target.Architecture
                }
            }
    }

    return @(
        $results |
            Sort-Object DisplayName, DisplayVersion, Publisher -Unique
    )
}

function Get-AppNameCandidates {
    param(
        [Parameter(Mandatory)]
        $App
    )

    $names = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$App.name)) {
        $names += [string]$App.name
    }

    if ($null -ne $App.detection -and $null -ne $App.detection.matchNames) {
        $names += @(
            $App.detection.matchNames |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                ForEach-Object { [string]$_ }
        )
    }

    return @($names | Sort-Object -Unique)
}

function Find-InstalledProgramMatch {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        [object[]]$InstalledPrograms
    )

    $exactNames = @{}
    $coreNames = @{}
    $appTokens = @{}

    foreach ($name in Get-AppNameCandidates -App $App) {
        $normalizedName = Normalize-AppText -Value $name
        $coreName = Get-VersionlessAppText -Value $name

        if (-not [string]::IsNullOrWhiteSpace($normalizedName)) {
            $exactNames[$normalizedName] = $true
        }

        if (-not [string]::IsNullOrWhiteSpace($coreName)) {
            $coreNames[$coreName] = $true
            foreach ($token in (Get-SearchTokens -Value $coreName)) {
                $appTokens[$token] = $true
            }
        }
    }

    if ($exactNames.Count -eq 0 -and $coreNames.Count -eq 0) {
        return $null
    }

    $normalizedPublisher = Normalize-AppText -Value ([string]$App.publisher)
    $expectedVersion = Convert-EmptyToNull -Value ([string]$App.expectedVersion)

    $programMatches = @(
        foreach ($program in $InstalledPrograms) {
        $programName = Normalize-AppText -Value ([string]$program.DisplayName)
        $programCoreName = Get-VersionlessAppText -Value ([string]$program.DisplayName)
        $programPublisher = Normalize-AppText -Value ([string]$program.Publisher)
        $score = 0

        if ($exactNames.ContainsKey($programName)) {
            $score += 120
        }

        if (-not [string]::IsNullOrWhiteSpace($programCoreName) -and $coreNames.ContainsKey($programCoreName)) {
            $score += 100
        }
        elseif (-not [string]::IsNullOrWhiteSpace($programCoreName)) {
            foreach ($coreName in $coreNames.Keys) {
                if ($programCoreName.Contains($coreName) -or $coreName.Contains($programCoreName)) {
                    $score += 75
                    break
                }
            }
        }

        if ($appTokens.Count -gt 0) {
            $matchingTokens = 0
            foreach ($token in $appTokens.Keys) {
                if ($programCoreName.Contains($token) -or $programName.Contains($token)) {
                    $matchingTokens += 1
                }
            }

            if ($matchingTokens -gt 0) {
                $score += ($matchingTokens * 12)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($normalizedPublisher) -and $programPublisher -eq $normalizedPublisher) {
            $score += 20
        }

        if ($null -ne $expectedVersion -and ([string]$program.DisplayVersion) -eq $expectedVersion) {
            $score += 15
        }

        if ($score -lt 60) {
            continue
        }

        [PSCustomObject]@{
            Program = $program
            Score = $score
        }
        }
    )

    if ($programMatches.Count -eq 0) {
        return $null
    }

    return @(
        $programMatches |
            Sort-Object -Property Score -Descending
    )[0].Program
}

function Update-AppCatalogStatuses {
    param(
        [Parameter(Mandatory)]
        $Catalog,

        [Parameter(Mandatory)]
        [object[]]$InstalledPrograms
    )

    $checkedAt = (Get-Date).ToString('o')

    foreach ($app in @($Catalog.apps)) {
        Initialize-AppEntry -App $app
        $classification = Get-AppClassification -App $app
        $app.classification.bucket = [string]$classification.bucket
        $app.classification.recommendedAction = [string]$classification.recommendedAction
        $app.classification.reason = [string]$classification.reason
        $match = Find-InstalledProgramMatch -App $app -InstalledPrograms $InstalledPrograms

        $app.detection.lastCheckedAt = $checkedAt
        if ($null -ne $match) {
            $app.status = 'installed'
            $app.detection.lastSeenName = [string]$match.DisplayName
            $app.detection.lastSeenVersion = if ([string]::IsNullOrWhiteSpace([string]$match.DisplayVersion)) { $null } else { [string]$match.DisplayVersion }
        }
        elseif ($app.desired -eq $false) {
            $app.status = 'ignored'
            $app.detection.lastSeenName = $null
            $app.detection.lastSeenVersion = $null
        }
        else {
            $app.status = 'missing'
            $app.detection.lastSeenName = $null
            $app.detection.lastSeenVersion = $null
        }

        $app.installer.ready = [bool](
            (-not [string]::IsNullOrWhiteSpace([string]$app.installer.localPath) -and (Test-Path -LiteralPath $app.installer.localPath)) -or
            (-not [string]::IsNullOrWhiteSpace([string]$app.installer.downloadedPath) -and (Test-Path -LiteralPath $app.installer.downloadedPath))
        )
    }

    return $Catalog
}

function Get-DefaultInstallerExtensions {
    return @('.exe', '.msi', '.msix', '.msixbundle', '.appx', '.appxbundle', '.zip')
}

function Get-DefaultWebRequestHeaders {
    param(
        [AllowNull()]
        [string]$Referer,

        [switch]$ForBinaryDownload
    )

    $headers = @{
        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0 Safari/537.36'
        'Accept-Language' = 'en-GB,en-US;q=0.9,en;q=0.8'
        'Accept' = if ($ForBinaryDownload.IsPresent) { '*/*' } else { 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' }
    }

    if (-not [string]::IsNullOrWhiteSpace($Referer)) {
        $headers['Referer'] = $Referer
    }

    return $headers
}

function Invoke-WebContentRequest {
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [int]$TimeoutSeconds = 120,

        [AllowNull()]
        [string]$Referer
    )

    $curlCommand = Get-Command -Name 'curl.exe' -ErrorAction SilentlyContinue
    if ($null -ne $curlCommand) {
        $headers = Get-DefaultWebRequestHeaders -Referer $Referer
        $curlArgs = @(
            '--location',
            '--silent',
            '--show-error',
            '--max-time', [string]$TimeoutSeconds,
            '--user-agent', [string]$headers['User-Agent'],
            '--header', ('Accept-Language: ' + [string]$headers['Accept-Language']),
            '--header', ('Accept: ' + [string]$headers['Accept']),
            '--url', $Url
        )

        if (-not [string]::IsNullOrWhiteSpace($Referer)) {
            $curlArgs += @('--referer', $Referer)
        }

        try {
            $rawOutput = @(& $curlCommand.Source @curlArgs)
            if ($LASTEXITCODE -eq 0) {
                $content = [string]::Join([Environment]::NewLine, @($rawOutput | ForEach-Object { [string]$_ }))
                if (-not [string]::IsNullOrWhiteSpace($content)) {
                    return [PSCustomObject]@{
                        Content = $content
                        Links = @()
                    }
                }
            }
        }
        catch {
        }
    }

    try {
        $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -MaximumRedirection 5 -TimeoutSec $TimeoutSeconds -Headers (Get-DefaultWebRequestHeaders -Referer $Referer)
        if ($null -ne $response) {
            return $response
        }
    }
    catch {
    }

    return $null
}

function Get-WebLinkPropertyValue {
    param(
        [Parameter(Mandatory)]
        $Link,

        [Parameter(Mandatory)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        if ($Link.PSObject.Properties[$propertyName]) {
            return [string]$Link.$propertyName
        }
    }

    return ''
}

function Get-WebResponseCandidateUrls {
    param(
        [Parameter(Mandatory)]
        $Response,

        [Parameter(Mandatory)]
        [string]$BaseUrl
    )

    $baseUri = [Uri]$BaseUrl
    $candidates = [System.Collections.Generic.List[string]]::new()

    $linkCollection = @()
    if ((Test-HasProperty -InputObject $Response -PropertyName 'Links') -and $null -ne $Response.Links) {
        $linkCollection = @($Response.Links)
    }

    foreach ($link in $linkCollection) {
        $href = Get-WebLinkPropertyValue -Link $link -PropertyNames @('href', 'Href')
        if (-not [string]::IsNullOrWhiteSpace($href)) {
            $candidates.Add($href)
        }
    }

    $rawContent = ''
    if ((Test-HasProperty -InputObject $Response -PropertyName 'Content') -and $null -ne $Response.Content) {
        $rawContent = [string]$Response.Content
    }

    if (-not [string]::IsNullOrWhiteSpace($rawContent)) {
        $decodedContent = [System.Net.WebUtility]::HtmlDecode($rawContent)

        foreach ($match in [regex]::Matches($decodedContent, '(?i)href\s*=\s*["''][^"''#>]+["'']')) {
            $rawHref = [string]$match.Value
            $cleanHref = $rawHref -replace '^(?i)href\s*=\s*["'']', ''
            $cleanHref = $cleanHref -replace '["'']$', ''
            if (-not [string]::IsNullOrWhiteSpace($cleanHref)) {
                $candidates.Add($cleanHref)
            }
        }

        foreach ($match in [regex]::Matches($decodedContent, '(?i)https?://[^\s"''<>]+')) {
            $candidates.Add([string]$match.Value)
        }
    }

    $resolvedUrls = foreach ($candidate in $candidates) {
        $workingCandidate = [string]$candidate
        if ([string]::IsNullOrWhiteSpace($workingCandidate)) {
            continue
        }

        $workingCandidate = $workingCandidate.Trim()
        while ($workingCandidate.Length -gt 0 -and '.,;)]}>'.Contains($workingCandidate.Substring($workingCandidate.Length - 1, 1))) {
            $workingCandidate = $workingCandidate.Substring(0, $workingCandidate.Length - 1)
        }

        if ([string]::IsNullOrWhiteSpace($workingCandidate)) {
            continue
        }

        if ($workingCandidate.StartsWith('mailto:', [System.StringComparison]::OrdinalIgnoreCase) -or $workingCandidate.StartsWith('javascript:', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $workingCandidate = [System.Net.WebUtility]::HtmlDecode($workingCandidate)
        if ($workingCandidate.Contains('%')) {
            try {
                $workingCandidate = [Uri]::UnescapeDataString($workingCandidate)
            }
            catch {
            }
        }

        try {
            [Uri]::new($baseUri, $workingCandidate).AbsoluteUri
        }
        catch {
            continue
        }
    }

    return @($resolvedUrls | Sort-Object -Unique)
}

function Get-AdditionalReferencePageUrls {
    param(
        [Parameter(Mandatory)]
        [string]$PageUrl
    )

    $pageUri = [Uri]$PageUrl
    $pageLower = $PageUrl.ToLowerInvariant()
    $extraUrls = [System.Collections.Generic.List[string]]::new()

    if ($pageUri.Host.EndsWith('microsoft.com') -and $pageLower.Contains('/download/details.aspx')) {
        $idMatch = [regex]::Match([string]$pageUri.Query, '(?i)(?:^\?|&)id=([^&]+)')
        if ($idMatch.Success) {
            $extraUrls.Add('https://www.microsoft.com/en-us/download/confirmation.aspx?id=' + [Uri]::UnescapeDataString($idMatch.Groups[1].Value))
        }
    }

    if ($pageUri.Host.EndsWith('techwebasto.com') -and -not $pageLower.Contains('/tech-tools/software.html')) {
        $extraUrls.Add('https://www.techwebasto.com/tech-tools/software.html')
    }

    return @($extraUrls | Sort-Object -Unique)
}

function Test-DirectInstallerUrl {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    $urlLower = $Url.ToLowerInvariant()
    if ($urlLower -match '\.(exe|msi|msix|msixbundle|appx|appxbundle|zip)([?#].*)?$') {
        return $true
    }

    try {
        $uri = [Uri]$Url
        return $uri.Host -match '(^|\.)download\.microsoft\.com$'
    }
    catch {
        return $false
    }
}

function Get-InstallerCandidateScore {
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [Uri]$ReferenceUri,

        [Parameter(Mandatory)]
        [string[]]$AppTokens,

        [switch]$ForFollowUpPage
    )

    $candidateUri = [Uri]$Url
    $candidateLower = $Url.ToLowerInvariant()
    $score = 0

    if ($candidateUri.Host -eq $ReferenceUri.Host) {
        $score += 50
    }
    elseif ($candidateUri.Host.EndsWith('.' + $ReferenceUri.Host)) {
        $score += 25
    }

    if ($candidateUri.Scheme -eq 'https') {
        $score += 10
    }

    if ($candidateLower.Contains('download')) {
        $score += 20
    }

    if ($candidateLower.Contains('setup') -or $candidateLower.Contains('installer')) {
        $score += 10
    }

    if ($ForFollowUpPage) {
        foreach ($keyword in @('confirmation.aspx', 'details.aspx', 'software', 'support', 'drivers', 'detailid=', 'downloads')) {
            if ($candidateLower.Contains($keyword)) {
                $score += 10
            }
        }
    }

    foreach ($token in $AppTokens) {
        if ($candidateLower.Contains($token)) {
            $score += 15
        }
    }

    return $score
}

function Get-ContentContextTokenScore {
    param(
        [AllowNull()]
        [string]$Context,

        [Parameter(Mandatory)]
        [string[]]$AppTokens
    )

    $normalizedContext = Normalize-AppText -Value $Context
    if ([string]::IsNullOrWhiteSpace($normalizedContext)) {
        return 0
    }

    $score = 0
    foreach ($token in $AppTokens) {
        if ($normalizedContext.Contains($token)) {
            $score += 20
        }
    }

    return $score
}

function Get-MicrosoftDownloadCenterCandidates {
    param(
        [Parameter(Mandatory)]
        [string]$PageUrl,

        [AllowNull()]
        [string]$RawContent,

        [Parameter(Mandatory)]
        [string[]]$AppTokens
    )

    if ([string]::IsNullOrWhiteSpace($RawContent)) {
        return @()
    }

    $pageUri = [Uri]$PageUrl
    if (-not $pageUri.Host.EndsWith('microsoft.com') -or -not $PageUrl.ToLowerInvariant().Contains('/download/details.aspx')) {
        return @()
    }

    $detailsMatch = [regex]::Match($RawContent, 'window\.__DLCDetails__\s*=\s*(\{.+?\})\s*;', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $detailsMatch.Success) {
        return @()
    }

    try {
        $detailsRoot = $detailsMatch.Groups[1].Value | ConvertFrom-Json
    }
    catch {
        return @()
    }

    $detailsView = $detailsRoot.dlcDetailsView
    if ($null -eq $detailsView) {
        return @()
    }

    $referenceUri = [Uri]$PageUrl
    $detailContext = [string]::Join(' ', @(
        [string]$detailsView.downloadTitle,
        [string]$detailsView.downloadDescription,
        [string]$detailsView.detailsSection,
        [string]$detailsView.installInstructionSection,
        [string]$detailsView.additionalInformationSection
    ))

    return @(
        @($detailsView.downloadFile) |
            ForEach-Object {
                $url = [string]$_.url
                if ([string]::IsNullOrWhiteSpace($url)) {
                    return
                }

                $fileName = [string]$_.name
                $score = Get-InstallerCandidateScore -Url $url -ReferenceUri $referenceUri -AppTokens $AppTokens
                $score += Get-ContentContextTokenScore -Context ($fileName + ' ' + $detailContext) -AppTokens $AppTokens

                if ([string]$_.isPrimary -eq 'True') {
                    $score += 120
                }

                if ($fileName -match '(?i)(amd64|x64)') {
                    $score += 20
                }

                if ($fileName -match '(?i)(en-us|english)') {
                    $score += 10
                }

                [PSCustomObject]@{
                    url = $url
                    score = $score
                }
            } |
            Where-Object { $null -ne $_ } |
            Sort-Object -Property @(
                @{ Expression = 'score'; Descending = $true },
                'url'
            )
    )
}

function Get-AttributeDownloadCandidates {
    param(
        [Parameter(Mandatory)]
        [string]$PageUrl,

        [AllowNull()]
        [string]$RawContent,

        [Parameter(Mandatory)]
        [string[]]$AppTokens
    )

    if ([string]::IsNullOrWhiteSpace($RawContent)) {
        return @()
    }

    $referenceUri = [Uri]$PageUrl
    $decodedContent = [System.Net.WebUtility]::HtmlDecode($RawContent)
    $downloadUrlMatches = [regex]::Matches($decodedContent, '(?is)<[^>]*data-download-url\s*=\s*["''](?<url>[^"'']+)["''][^>]*>')
    if ($downloadUrlMatches.Count -eq 0) {
        return @()
    }

    $candidates = foreach ($match in $downloadUrlMatches) {
        $rawUrl = [string]$match.Groups['url'].Value
        if ([string]::IsNullOrWhiteSpace($rawUrl)) {
            continue
        }

        try {
            $absoluteUrl = [Uri]::new($referenceUri, $rawUrl).AbsoluteUri
        }
        catch {
            continue
        }

        $snippetStart = [Math]::Max(0, $match.Index - 1200)
        $snippetLength = [Math]::Min($decodedContent.Length - $snippetStart, 2400)
        $snippet = $decodedContent.Substring($snippetStart, $snippetLength)

        $score = Get-InstallerCandidateScore -Url $absoluteUrl -ReferenceUri $referenceUri -AppTokens $AppTokens
        $score += Get-ContentContextTokenScore -Context ($match.Value + ' ' + $snippet) -AppTokens $AppTokens

        [PSCustomObject]@{
            url = $absoluteUrl
            score = $score
        }
    }

    return @(
        $candidates |
            Group-Object -Property url |
            ForEach-Object {
                $_.Group |
                    Sort-Object -Property @(
                        @{ Expression = 'score'; Descending = $true },
                        'url'
                    ) |
                    Select-Object -First 1
            } |
            Sort-Object -Property @(
                @{ Expression = 'score'; Descending = $true },
                'url'
            )
    )
}

function Get-DownloadFileNameFromResponse {
    param(
        [AllowNull()]
        $Response,

        [AllowNull()]
        [string]$FallbackFileName,

        [Parameter(Mandatory)]
        [string]$DefaultFileName
    )

    $fileName = $null

    if ($null -ne $Response -and (Test-HasProperty -InputObject $Response -PropertyName 'Headers') -and $null -ne $Response.Headers) {
        $contentDisposition = $null
        if ($Response.Headers['Content-Disposition']) {
            $contentDisposition = [string]$Response.Headers['Content-Disposition']
        }

        if (-not [string]::IsNullOrWhiteSpace($contentDisposition)) {
            $fileNameMatch = [regex]::Match($contentDisposition, '(?i)filename\*?=(?:UTF-8''''|"|)(?<name>[^";]+)')
            if ($fileNameMatch.Success) {
                $fileName = [string]$fileNameMatch.Groups['name'].Value.Trim()
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($fileName) -and $null -ne $Response -and (Test-HasProperty -InputObject $Response -PropertyName 'BaseResponse') -and $null -ne $Response.BaseResponse) {
        $responseUri = $Response.BaseResponse.ResponseUri
        if ($null -ne $responseUri) {
            $candidate = [System.IO.Path]::GetFileName(([Uri]$responseUri).AbsolutePath)
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $fileName = $candidate
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = $FallbackFileName
    }

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = $DefaultFileName
    }

    return $fileName
}

function Resolve-ManualReferenceDownload {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        [string]$ReferenceUrl,

        [int]$MaxDepth = 2,

        [int]$TimeoutSeconds = 120
    )

    $searchTokens = @(
        (Get-SearchTokens -Value ([string]$App.name)) +
        (Get-SearchTokens -Value ([string]$App.publisher))
    ) | Select-Object -Unique

    $visitedUrls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pendingPages = [System.Collections.Generic.Queue[object]]::new()
    $pendingPages.Enqueue([PSCustomObject]@{
        url = $ReferenceUrl
        depth = 0
    })

    while ($pendingPages.Count -gt 0) {
        $page = $pendingPages.Dequeue()
        $currentUrl = [string]$page.url
        $currentDepth = [int]$page.depth

        if ([string]::IsNullOrWhiteSpace($currentUrl) -or $visitedUrls.Contains($currentUrl)) {
            continue
        }

        $null = $visitedUrls.Add($currentUrl)

        if (Test-DirectInstallerUrl -Url $currentUrl) {
            return [PSCustomObject]@{
                downloadUrl = $currentUrl
                sourcePage = $currentUrl
                detection = 'direct-url'
            }
        }

        $response = Invoke-WebContentRequest -Url $currentUrl -TimeoutSeconds $TimeoutSeconds
        if ($null -eq $response) {
            if ($currentDepth -lt $MaxDepth) {
                foreach ($bootstrapUrl in (Get-AdditionalReferencePageUrls -PageUrl $currentUrl)) {
                    if (-not $visitedUrls.Contains($bootstrapUrl)) {
                        $pendingPages.Enqueue([PSCustomObject]@{
                            url = $bootstrapUrl
                            depth = $currentDepth + 1
                        })
                    }
                }
            }

            continue
        }

        $currentUri = [Uri]$currentUrl
        $rawContent = ''
        if ((Test-HasProperty -InputObject $response -PropertyName 'Content') -and $null -ne $response.Content) {
            $rawContent = [string]$response.Content
        }

        $candidateUrls = @(
            (Get-WebResponseCandidateUrls -Response $response -BaseUrl $currentUrl) +
            (Get-AdditionalReferencePageUrls -PageUrl $currentUrl)
        ) | Sort-Object -Unique

        $specialDirectCandidates = @(
            (Get-MicrosoftDownloadCenterCandidates -PageUrl $currentUrl -RawContent $rawContent -AppTokens $searchTokens) +
            (Get-AttributeDownloadCandidates -PageUrl $currentUrl -RawContent $rawContent -AppTokens $searchTokens)
        )

        $directCandidates = @(
            $specialDirectCandidates +
            @(
                $candidateUrls |
                    Where-Object { Test-DirectInstallerUrl -Url $_ } |
                    ForEach-Object {
                        [PSCustomObject]@{
                            url = [string]$_
                            score = Get-InstallerCandidateScore -Url ([string]$_) -ReferenceUri $currentUri -AppTokens $searchTokens
                        }
                    }
            ) |
                Group-Object -Property url |
                ForEach-Object {
                    $_.Group |
                        Sort-Object -Property @(
                            @{ Expression = 'score'; Descending = $true },
                            'url'
                        ) |
                        Select-Object -First 1
                } |
                Sort-Object -Property @(
                    @{ Expression = 'score'; Descending = $true },
                    'url'
                )
        )

        if ($directCandidates.Count -gt 0) {
            return [PSCustomObject]@{
                downloadUrl = [string]$directCandidates[0].url
                sourcePage = $currentUrl
                detection = if ($currentDepth -eq 0) { 'page-link' } else { 'follow-up-page' }
            }
        }

        if ($currentDepth -ge $MaxDepth) {
            continue
        }

        $followUpCandidates = @(
            $candidateUrls |
                Where-Object {
                    -not $visitedUrls.Contains($_) -and
                    -not (Test-DirectInstallerUrl -Url $_)
                } |
                ForEach-Object {
                    $score = Get-InstallerCandidateScore -Url ([string]$_) -ReferenceUri $currentUri -AppTokens $searchTokens -ForFollowUpPage
                    if ($score -lt 25) {
                        return
                    }

                    [PSCustomObject]@{
                        url = [string]$_
                        score = $score
                    }
                } |
                Where-Object { $null -ne $_ } |
                Sort-Object -Property @(
                    @{ Expression = 'score'; Descending = $true },
                    'url'
                ) |
                Select-Object -First 5
        )

        foreach ($followUpCandidate in $followUpCandidates) {
            $pendingPages.Enqueue([PSCustomObject]@{
                url = [string]$followUpCandidate.url
                depth = $currentDepth + 1
            })
        }
    }

    return $null
}

function Stage-ManualReferenceInstaller {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        [string]$StageDirectory,

        [switch]$Force,

        [int]$MaxDepth = 2,

        [int]$TimeoutSeconds = 120
    )

    $referenceUrl = [string]$App.installer.manualReferenceUrl
    if ([string]::IsNullOrWhiteSpace($referenceUrl)) {
        return [PSCustomObject]@{
            status = 'no-reference-url'
            stagedPath = $null
            downloadUrl = $null
            detection = $null
            details = 'no manual reference URL recorded'
        }
    }

    $resolvedStageDirectory = Resolve-AppPath -Path $StageDirectory
    if (-not (Test-Path -LiteralPath $resolvedStageDirectory)) {
        New-Item -Path $resolvedStageDirectory -ItemType Directory -Force | Out-Null
    }

    if (-not $Force.IsPresent -and -not [string]::IsNullOrWhiteSpace([string]$App.installer.downloadedPath) -and (Test-Path -LiteralPath $App.installer.downloadedPath)) {
        $App.installer.preferredSource = 'manual-url'
        $App.installer.ready = $true
        return [PSCustomObject]@{
            status = 'downloaded'
            stagedPath = [string]$App.installer.downloadedPath
            downloadUrl = $null
            detection = 'existing-file'
            details = $null
        }
    }

    $safeDirectoryName = (Normalize-AppText -Value ([string]$App.name)) -replace ' ', '-'
    if ([string]::IsNullOrWhiteSpace($safeDirectoryName)) {
        $safeDirectoryName = 'app'
    }

    $appStageDirectory = Join-Path -Path $resolvedStageDirectory -ChildPath $safeDirectoryName
    if (-not (Test-Path -LiteralPath $appStageDirectory)) {
        New-Item -Path $appStageDirectory -ItemType Directory -Force | Out-Null
    }

    $existingFile = @(
        Get-ChildItem -LiteralPath $appStageDirectory -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { (Get-DefaultInstallerExtensions) -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object LastWriteTime -Descending
    ) | Select-Object -First 1

    if ($null -ne $existingFile -and -not $Force.IsPresent) {
        $App.installer.downloadedPath = $existingFile.FullName
        $App.installer.preferredSource = 'manual-url'
        $App.installer.ready = $true
        return [PSCustomObject]@{
            status = 'downloaded'
            stagedPath = $existingFile.FullName
            downloadUrl = $null
            detection = 'existing-file'
            details = $null
        }
    }

    try {
        $resolution = Resolve-ManualReferenceDownload -App $App -ReferenceUrl $referenceUrl -MaxDepth $MaxDepth -TimeoutSeconds $TimeoutSeconds
    }
    catch {
        return [PSCustomObject]@{
            status = 'error'
            stagedPath = $null
            downloadUrl = $null
            detection = $null
            details = $_.Exception.Message
        }
    }

    if ($null -eq $resolution) {
        return [PSCustomObject]@{
            status = 'no-direct-link'
            stagedPath = $null
            downloadUrl = $null
            detection = $null
            details = 'reference page did not expose a direct installer link'
        }
    }

    $downloadUrl = [string]$resolution.downloadUrl
    $fileName = [System.IO.Path]::GetFileName(([Uri]$downloadUrl).AbsolutePath)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = $safeDirectoryName + '.download'
    }

    $temporaryPath = Join-Path -Path $appStageDirectory -ChildPath $fileName
    $destinationPath = $temporaryPath
    try {
        if ($Force.IsPresent -or -not (Test-Path -LiteralPath $destinationPath)) {
            $downloadResponse = Invoke-WebRequest -Uri $downloadUrl -UseBasicParsing -OutFile $temporaryPath -PassThru -MaximumRedirection 5 -TimeoutSec $TimeoutSeconds -Headers (Get-DefaultWebRequestHeaders -Referer ([string]$resolution.sourcePage) -ForBinaryDownload)
            $resolvedFileName = Get-DownloadFileNameFromResponse -Response $downloadResponse -FallbackFileName $fileName -DefaultFileName ($safeDirectoryName + '.bin')
            $destinationPath = Join-Path -Path $appStageDirectory -ChildPath $resolvedFileName

            if ($temporaryPath -ne $destinationPath) {
                if (Test-Path -LiteralPath $destinationPath) {
                    Remove-Item -LiteralPath $destinationPath -Force -Confirm:$false -ErrorAction SilentlyContinue
                }

                Move-Item -LiteralPath $temporaryPath -Destination $destinationPath -Force -Confirm:$false
            }
        }
    }
    catch {
        return [PSCustomObject]@{
            status = 'error'
            stagedPath = $null
            downloadUrl = $downloadUrl
            detection = [string]$resolution.detection
            details = $_.Exception.Message
        }
    }

    if (-not (Test-Path -LiteralPath $destinationPath)) {
        return [PSCustomObject]@{
            status = 'download-failed'
            stagedPath = $null
            downloadUrl = $downloadUrl
            detection = [string]$resolution.detection
            details = 'download completed without a staged file'
        }
    }

    $App.installer.downloadedPath = $destinationPath
    $App.installer.preferredSource = 'manual-url'
    $App.installer.ready = $true

    return [PSCustomObject]@{
        status = 'downloaded'
        stagedPath = $destinationPath
        downloadUrl = $downloadUrl
        detection = [string]$resolution.detection
        details = $null
    }
}

function Get-InstallerFileInventory {
    param(
        [Parameter(Mandatory)]
        [string[]]$SearchRoots
    )

    $extensions = Get-DefaultInstallerExtensions
    $files = foreach ($root in $SearchRoots) {
        if ([string]::IsNullOrWhiteSpace($root)) {
            continue
        }

        $resolvedRoot = Resolve-AppPath -Path $root
        if (-not (Test-Path -LiteralPath $resolvedRoot)) {
            continue
        }

        Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
            ForEach-Object {
                [PSCustomObject]@{
                    Path = $_.FullName
                    FileName = $_.Name
                    Extension = $_.Extension.ToLowerInvariant()
                    SizeBytes = $_.Length
                    LastWriteTime = $_.LastWriteTime
                    NormalizedName = Normalize-AppText -Value $_.BaseName
                }
            }
    }

    return @($files)
}

function Get-SearchTokens {
    param(
        [AllowNull()]
        [string]$Value
    )

    $ignoredTokens = @('and', 'for', 'the', 'with', 'setup', 'installer', 'install', 'x64', 'x86', 'inc', 'llc', 'ltd', 'corp', 'corporation', 'company', 'co', 'software', 'systems', 'foundation')
    return @(
        (Normalize-AppText -Value $Value).Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) |
            Where-Object { $_.Length -ge 3 -and $_ -notin $ignoredTokens } |
            Select-Object -Unique
    )
}

function Get-AppInstallerCandidates {
    param(
        [Parameter(Mandatory)]
        $App,

        [Parameter(Mandatory)]
        [object[]]$InstallerFiles,

        [int]$Limit = 5
    )

    $appName = [string]$App.name
    $normalizedAppName = Normalize-AppText -Value $appName
    if ([string]::IsNullOrWhiteSpace($normalizedAppName)) {
        return @()
    }

    $nameTokens = Get-SearchTokens -Value $appName
    $publisherTokens = Get-SearchTokens -Value ([string]$App.publisher)

    $scoredCandidates = foreach ($file in $InstallerFiles) {
        $score = 0
        $fileName = [string]$file.NormalizedName
        if ([string]::IsNullOrWhiteSpace($fileName)) {
            continue
        }

        if ($fileName -eq $normalizedAppName) {
            $score += 120
        }
        elseif ($fileName.Contains($normalizedAppName)) {
            $score += 90
        }

        foreach ($token in $nameTokens) {
            if ($fileName.Contains($token)) {
                $score += 15
            }
        }

        foreach ($token in $publisherTokens) {
            if ($fileName.Contains($token)) {
                $score += 5
            }
        }

        if ($score -lt 30) {
            continue
        }

        [PSCustomObject]@{
            path = $file.Path
            fileName = $file.FileName
            extension = $file.Extension.TrimStart('.')
            sizeBytes = $file.SizeBytes
            lastWriteTime = $file.LastWriteTime
            score = $score
        }
    }

    return @(
        $scoredCandidates |
            Sort-Object -Property score, lastWriteTime -Descending |
            Select-Object -First $Limit
    )
}

function Convert-WingetSearchTextToPackages {
    param(
        [AllowNull()]
        [string[]]$Lines
    )

    $packages = @()
    $inTable = $false
    $headerLine = $null
    $idStart = -1
    $versionStart = -1
    $sourceStart = -1
    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($null -eq $headerLine -and $line -match '^Name\s+Id\s+Version') {
            $headerLine = [string]$line
            $idStart = $headerLine.IndexOf('Id')
            $versionStart = $headerLine.IndexOf('Version')
            $sourceStart = $headerLine.IndexOf('Source')
            continue
        }

        if ($line -match '^-{5,}$') {
            $inTable = $true
            continue
        }

        if (-not $inTable) {
            continue
        }

        if ($idStart -lt 0 -or $versionStart -lt 0) {
            continue
        }

        $workingLine = [string]$line
        if ($workingLine.Length -lt $versionStart) {
            continue
        }

        if ($sourceStart -ge 0 -and $workingLine.Length -lt $sourceStart) {
            $workingLine = $workingLine.PadRight($sourceStart)
        }

        $name = $workingLine.Substring(0, $idStart).Trim()
        $id = $workingLine.Substring($idStart, $versionStart - $idStart).Trim()
        $version = if ($sourceStart -ge 0) {
            $workingLine.Substring($versionStart, $sourceStart - $versionStart).Trim()
        }
        else {
            $workingLine.Substring($versionStart).Trim()
        }
        $source = if ($sourceStart -ge 0 -and $workingLine.Length -ge $sourceStart) {
            $workingLine.Substring($sourceStart).Trim()
        }
        else {
            'winget'
        }

        if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($id) -and -not [string]::IsNullOrWhiteSpace($version)) {
            $packages += [PSCustomObject]@{
                Name = $name
                Id = $id
                Version = $version
                Source = $source
            }
        }
    }

    return @($packages)
}

function Get-WingetPackageById {
    param(
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return $null
    }

    $searchLines = & winget.exe search --id $Id --exact --count 1 --source winget --accept-source-agreements --disable-interactivity 2>&1
    $packages = Convert-WingetSearchTextToPackages -Lines $searchLines
    if (@($packages).Count -eq 0) {
        return $null
    }

    return @($packages)[0]
}

function Get-WingetPackageCandidates {
    param(
        [Parameter(Mandatory)]
        $App,

        [int]$Count = 8
    )

    $name = [string]$App.name
    if ([string]::IsNullOrWhiteSpace($name)) {
        return @()
    }

    $queries = @($name)
    $versionlessName = Get-VersionlessAppText -Value $name
    if (-not [string]::IsNullOrWhiteSpace($versionlessName) -and $versionlessName -ne (Normalize-AppText -Value $name)) {
        $queries += $versionlessName
    }

    $appCoreTokens = @(Get-SearchTokens -Value $versionlessName)
    $publisherTokens = @(Get-SearchTokens -Value ([string]$App.publisher))

    $seenIds = @{}
    $scored = foreach ($query in ($queries | Select-Object -Unique)) {
        $searchLines = & winget.exe search --name $query --count $Count --source winget --accept-source-agreements --disable-interactivity 2>&1
        $packages = Convert-WingetSearchTextToPackages -Lines $searchLines
        foreach ($package in $packages) {
            if ($seenIds.ContainsKey($package.Id)) {
                continue
            }

            $seenIds[$package.Id] = $true
            $packageName = Normalize-AppText -Value ([string]$package.Name)
            $packageId = Normalize-AppText -Value ([string]$package.Id)
            $appName = Normalize-AppText -Value $name
            $appCoreName = Get-VersionlessAppText -Value $name
            $score = 0

            if ($packageName -eq $appName -or $packageId -eq $appName) {
                $score += 120
            }

            if (-not [string]::IsNullOrWhiteSpace($appCoreName)) {
                if ($packageName -eq $appCoreName -or $packageId -eq $appCoreName) {
                    $score += 110
                }
                elseif ($packageName.Contains($appCoreName) -or $appCoreName.Contains($packageName)) {
                    $score += 80
                }
            }

            foreach ($token in (Get-SearchTokens -Value $name)) {
                if ($packageName.Contains($token) -or $packageId.Contains($token)) {
                    $score += 15
                }
            }

            $publisherTokenMatches = 0
            foreach ($token in $publisherTokens) {
                if ($packageName.Contains($token) -or $packageId.Contains($token)) {
                    $score += 5
                    $publisherTokenMatches += 1
                }
            }

            if (
                $appCoreTokens.Count -le 1 -and
                $packageName -ne $appName -and
                $packageName -ne $appCoreName -and
                $packageId -ne $appName -and
                $packageId -ne $appCoreName -and
                $publisherTokenMatches -eq 0
            ) {
                continue
            }

            if ($score -lt 40) {
                continue
            }

            [PSCustomObject]@{
                name = [string]$package.Name
                id = [string]$package.Id
                version = [string]$package.Version
                source = [string]$package.Source
                score = $score
            }
        }
    }

    return @(
        $scored |
            Sort-Object -Property score, id -Descending |
            Select-Object -First $Count
    )
}

Export-ModuleMember -Function @(
    'Convert-EmptyToNull',
    'Convert-WingetSearchTextToPackages',
    'Find-InstalledProgramMatch',
    'Get-AppClassification',
    'Get-AppInstallerCandidates',
    'Get-AppNameCandidates',
    'Get-DefaultInstallerExtensions',
    'Get-DefaultWebRequestHeaders',
    'Get-InstalledProgramInventory',
    'Get-InstallerFileInventory',
    'Resolve-ManualReferenceDownload',
    'Get-VersionlessAppText',
    'Get-WingetPackageById',
    'Get-WingetPackageCandidates',
    'Initialize-AppEntry',
    'Normalize-AppText',
    'Read-AppCatalog',
    'Resolve-AppPath',
    'Save-AppCatalog',
    'Test-AppCatalogStructure',
    'Test-InstallQueueStructure',
    'Stage-ManualReferenceInstaller',
    'Test-HasProperty',
    'Update-AppCatalogStatuses'
)