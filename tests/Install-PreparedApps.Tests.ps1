Describe 'Install-PreparedApps' {
    BeforeAll {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\scripts\Install-PreparedApps.ps1'
    }

    It 'writes an install-plan document with summary data for ready MSI items' {
        $workspaceRoot = Join-Path -Path $TestDrive -ChildPath 'plan-ready-msi'
        $catalogDirectory = Join-Path -Path $workspaceRoot -ChildPath 'catalog'
        $outputDirectory = Join-Path -Path $workspaceRoot -ChildPath 'output'
        $stageDirectory = Join-Path -Path $workspaceRoot -ChildPath 'staged-installers'

        New-Item -Path $catalogDirectory -ItemType Directory -Force | Out-Null
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        New-Item -Path $stageDirectory -ItemType Directory -Force | Out-Null

        $catalogPath = Join-Path -Path $catalogDirectory -ChildPath 'apps.json'
        $queuePath = Join-Path -Path $outputDirectory -ChildPath 'install-queue.json'
        $planPath = Join-Path -Path $outputDirectory -ChildPath 'install-plan.json'
        $stagedInstallerPath = Join-Path -Path $stageDirectory -ChildPath 'contoso-tool.msi'

        Set-Content -LiteralPath $stagedInstallerPath -Value 'dummy msi content' -Encoding UTF8

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

        $queueDocument = [PSCustomObject]@{
            generatedAt = '2026-03-11T00:00:00.0000000Z'
            items = @(
                [PSCustomObject]@{
                    name = 'Contoso Tool'
                    status = 'missing'
                    source = 'local'
                    ready = $true
                    stagedPath = $stagedInstallerPath
                    details = $null
                }
            )
        }

        $catalog | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $catalogPath -Encoding UTF8
        $queueDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $queuePath -Encoding UTF8

        $result = & $scriptPath -CatalogPath $catalogPath -QueuePath $queuePath -Mode Plan -LogPath $planPath
        $persisted = Get-Content -LiteralPath $planPath -Raw | ConvertFrom-Json

        $result.documentType | Should -Be 'install-plan'
        $result.summary.totalItems | Should -Be 1
        $result.summary.readyItems | Should -Be 1
        $result.summary.plannedItems | Should -Be 1
        $result.items[0].method | Should -Be 'msi'
        $result.items[0].stagedPath | Should -Be $stagedInstallerPath

        (Test-Path -LiteralPath $planPath) | Should -BeTrue
        $persisted.documentType | Should -Be 'install-plan'
        $persisted.summary.readyItems | Should -Be 1
        $persisted.items[0].command | Should -Be 'msiexec.exe'
    }

    It 'marks EXE installers without silent arguments as skipped in plan mode' {
        $workspaceRoot = Join-Path -Path $TestDrive -ChildPath 'plan-unready-exe'
        $catalogDirectory = Join-Path -Path $workspaceRoot -ChildPath 'catalog'
        $outputDirectory = Join-Path -Path $workspaceRoot -ChildPath 'output'
        $stageDirectory = Join-Path -Path $workspaceRoot -ChildPath 'staged-installers'

        New-Item -Path $catalogDirectory -ItemType Directory -Force | Out-Null
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        New-Item -Path $stageDirectory -ItemType Directory -Force | Out-Null

        $catalogPath = Join-Path -Path $catalogDirectory -ChildPath 'apps.json'
        $queuePath = Join-Path -Path $outputDirectory -ChildPath 'install-queue.json'
        $planPath = Join-Path -Path $outputDirectory -ChildPath 'install-plan.json'
        $stagedInstallerPath = Join-Path -Path $stageDirectory -ChildPath 'fabrikam-tool.exe'

        Set-Content -LiteralPath $stagedInstallerPath -Value 'dummy exe content' -Encoding UTF8

        $catalog = [PSCustomObject]@{
            schemaVersion = 1
            generatedAt = '2026-03-11T00:00:00.0000000Z'
            apps = @(
                [PSCustomObject]@{
                    name = 'Fabrikam Tool'
                    publisher = 'Fabrikam'
                    expectedVersion = '2.0'
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

        $queueDocument = [PSCustomObject]@{
            generatedAt = '2026-03-11T00:00:00.0000000Z'
            items = @(
                [PSCustomObject]@{
                    name = 'Fabrikam Tool'
                    status = 'missing'
                    source = 'local'
                    ready = $true
                    stagedPath = $stagedInstallerPath
                    details = $null
                }
            )
        }

        $catalog | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $catalogPath -Encoding UTF8
        $queueDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $queuePath -Encoding UTF8

        $result = & $scriptPath -CatalogPath $catalogPath -QueuePath $queuePath -Mode Plan -LogPath $planPath

        $result.documentType | Should -Be 'install-plan'
        $result.summary.readyItems | Should -Be 0
        $result.summary.skippedItems | Should -Be 1
        $result.items[0].method | Should -Be 'exe'
        $result.items[0].ready | Should -BeFalse
        $result.items[0].executionState | Should -Be 'skipped'
        $result.items[0].details | Should -Be 'exe installer requires installer.installArgs or -AllowExeWithoutArgs'
    }
}