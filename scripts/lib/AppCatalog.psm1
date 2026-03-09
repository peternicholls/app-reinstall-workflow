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
        'extension sdk',
        'sdk arm',
        'sdk addon',
        'app certification kit',
        'kits configuration installer',
        'winappdeploy',
        'msi development tools',
        'supportedapilist',
        'coreeditorfonts'
    )

    $systemPatterns = @(
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

    foreach ($propertyName in 'installArgs', 'preferredSource') {
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

    $matches = @(
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

    if ($matches.Count -eq 0) {
        return $null
    }

    return @(
        $matches |
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
    'Get-InstalledProgramInventory',
    'Get-InstallerFileInventory',
    'Get-VersionlessAppText',
    'Get-WingetPackageById',
    'Get-WingetPackageCandidates',
    'Initialize-AppEntry',
    'Normalize-AppText',
    'Read-AppCatalog',
    'Resolve-AppPath',
    'Save-AppCatalog',
    'Test-HasProperty',
    'Update-AppCatalogStatuses'
)