[CmdletBinding()]
param(
    [ValidateSet('Table', 'Json', 'Csv')]
    [string]$Format = 'Table',

    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'lib\AppCatalog.psm1') -Force

$installedPrograms = @(
    Get-InstalledProgramInventory |
        Sort-Object DisplayName, DisplayVersion
)

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $resolvedOutputPath = Resolve-AppPath -Path $OutputPath
    $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    switch ($Format) {
        'Json' {
            $installedPrograms | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8
        }
        'Csv' {
            $installedPrograms | Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation -Encoding UTF8
        }
        default {
            $installedPrograms | Format-Table -AutoSize | Out-String | Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8
        }
    }
}

switch ($Format) {
    'Json' { $installedPrograms | ConvertTo-Json -Depth 6 }
    'Csv' { $installedPrograms }
    default { $installedPrograms }
}