Describe 'AppCatalog validation' {
    BeforeAll {
        $modulePath = Join-Path -Path $PSScriptRoot -ChildPath '..\scripts\lib\AppCatalog.psm1'
        Import-Module $modulePath -Force
    }

    It 'accepts a structurally valid catalog document' {
        $catalogPath = Join-Path -Path $TestDrive -ChildPath 'valid-catalog.json'
        $catalog = [PSCustomObject]@{
            schemaVersion = 1
            generatedAt = '2026-03-11T00:00:00.0000000Z'
            apps = @(
                [PSCustomObject]@{
                    name = 'Contoso Tool'
                    publisher = 'Contoso'
                    expectedVersion = '1.0'
                    desired = $true
                    status = 'missing'
                    detection = [PSCustomObject]@{}
                    latest = [PSCustomObject]@{}
                    classification = [PSCustomObject]@{}
                    installer = [PSCustomObject]@{}
                    notes = $null
                }
            )
        }

        $catalog | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $catalogPath -Encoding UTF8

        $result = Test-AppCatalogStructure -CatalogPath $catalogPath

        $result.isValid | Should -BeTrue
        $result.errorCount | Should -Be 0
        $result.summary.appCount | Should -Be 1
    }

    It 'reports missing required app properties in a catalog document' {
        $catalogPath = Join-Path -Path $TestDrive -ChildPath 'invalid-catalog.json'
        $catalog = [PSCustomObject]@{
            schemaVersion = 1
            apps = @(
                [PSCustomObject]@{
                    name = 'Broken App'
                    desired = $true
                    status = 'missing'
                    detection = [PSCustomObject]@{}
                    latest = [PSCustomObject]@{}
                    classification = [PSCustomObject]@{}
                }
            )
        }

        $catalog | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $catalogPath -Encoding UTF8

        $result = Test-AppCatalogStructure -CatalogPath $catalogPath

        $result.isValid | Should -BeFalse
        $result.errorCount | Should -BeGreaterThan 0
        @($result.issues | Where-Object { $_.path -eq 'apps[0].installer' }).Count | Should -Be 1
    }

    It 'reports missing ready flags in an install queue document' {
        $queuePath = Join-Path -Path $TestDrive -ChildPath 'invalid-queue.json'
        $queueDocument = [PSCustomObject]@{
            generatedAt = '2026-03-11T00:00:00.0000000Z'
            items = @(
                [PSCustomObject]@{
                    name = 'Contoso Tool'
                    source = 'local'
                    stagedPath = 'C:\Installers\contoso-tool.msi'
                }
            )
        }

        $queueDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $queuePath -Encoding UTF8

        $result = Test-InstallQueueStructure -QueuePath $queuePath

        $result.isValid | Should -BeFalse
        $result.errorCount | Should -BeGreaterThan 0
        @($result.issues | Where-Object { $_.path -eq 'items[0].ready' }).Count | Should -Be 1
    }
}