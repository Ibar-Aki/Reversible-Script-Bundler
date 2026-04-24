$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:BundleScriptPath = Join-Path -Path $script:RepoRoot -ChildPath 'bundle_system.ps1'
$script:BundleBatPath = Join-Path -Path $script:RepoRoot -ChildPath 'bundle_files.bat'
$script:RestoreBatPath = Join-Path -Path $script:RepoRoot -ChildPath 'restore_files.bat'

function New-TestWorkspace {
    $path = Join-Path -Path $env:TEMP -ChildPath ('bundle_system_test_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    return $path
}

function Remove-TestWorkspace {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Invoke-Converter {
    param(
        [string]$Mode,
        [string]$InputPath,
        [string]$OutputPath,
        [string]$RestoreInputPath,
        [string]$RestoreOutputPath,
        [string]$IgnoreFilePath,
        [string]$BundleRootName,
        [string]$ResultJsonPath
    )

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $script:BundleScriptPath,
        '-Mode', $Mode
    )

    if ($InputPath) {
        $arguments += @('-InputPath', $InputPath)
    }
    if ($OutputPath) {
        $arguments += @('-OutputPath', $OutputPath)
    }
    if ($RestoreInputPath) {
        $arguments += @('-RestoreInputPath', $RestoreInputPath)
    }
    if ($RestoreOutputPath) {
        $arguments += @('-RestoreOutputPath', $RestoreOutputPath)
    }
    if ($IgnoreFilePath) {
        $arguments += @('-IgnoreFilePath', $IgnoreFilePath)
    }
    if ($BundleRootName) {
        $arguments += @('-BundleRootName', $BundleRootName)
    }
    if ($ResultJsonPath) {
        $arguments += @('-ResultJsonPath', $ResultJsonPath)
    }

    $output = & powershell.exe @arguments | Out-String
    return [PSCustomObject]@{
        ExitCode = $LASTEXITCODE
        Output = $output
    }
}

function Invoke-BatchFile {
    param(
        [string]$BatchPath,
        [string[]]$Arguments
    )

    $parts = @('call', ('"' + $BatchPath + '"'))
    foreach ($argument in $Arguments) {
        if ($argument -match '^[A-Za-z0-9._\\/-]+$') {
            $parts += $argument
        }
        else {
            $parts += ('"' + $argument.Replace('"', '""') + '"')
        }
    }

    $commandLine = $parts -join ' '
    $output = & cmd.exe /c $commandLine | Out-String
    return [PSCustomObject]@{
        ExitCode = $LASTEXITCODE
        Output = $output
    }
}

function Get-ResultObject {
    param([string]$Path)

    return Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
}

function Get-FileBytes {
    param([string]$Path)

    return [System.IO.File]::ReadAllBytes($Path)
}

Describe 'bundle_system.ps1' {
    It 'bundles a directory and records metadata with .bundleignore support' {
        $workspace = New-TestWorkspace
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'input'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $resultJsonPath = Join-Path -Path $workspace -ChildPath 'bundle_result.json'
            $skipDirectory = Join-Path -Path $inputPath -ChildPath 'skipme'
            $keepDirectory = Join-Path -Path $inputPath -ChildPath 'keepme'
            New-Item -ItemType Directory -Path $skipDirectory -Force | Out-Null
            New-Item -ItemType Directory -Path $keepDirectory -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath 'main.ps1') -Value 'Write-Host "main"' -Encoding UTF8
            Set-Content -LiteralPath (Join-Path -Path $keepDirectory -ChildPath 'keep.ps1') -Value 'Write-Host "keep"' -Encoding UTF8
            Set-Content -LiteralPath (Join-Path -Path $skipDirectory -ChildPath 'skip.ps1') -Value 'Write-Host "skip"' -Encoding UTF8
            Set-Content -LiteralPath (Join-Path -Path $skipDirectory -ChildPath 'skip.exe') -Value 'ignored binary' -Encoding UTF8
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath '.bundleignore') -Value "skipme`r`n" -Encoding UTF8

            $invokeResult = Invoke-Converter -Mode 'Bundle' -InputPath $inputPath -OutputPath $outputPath -ResultJsonPath $resultJsonPath
            $invokeResult.ExitCode | Should Be 0

            $resultObject = Get-ResultObject -Path $resultJsonPath
            $bundleObject = Get-Content -Raw -LiteralPath $resultObject.BundlePath | ConvertFrom-Json

            $resultObject.Status | Should Be 'Success'
            $resultObject.BundledFileCount | Should Be 2
            $resultObject.SourceRoot | Should Be ([System.IO.Path]::GetFullPath($inputPath))
            $resultObject.ToolVersion | Should Be '1.2.0'
            $bundleObject.sourceRoot | Should Be ([System.IO.Path]::GetFullPath($inputPath))
            $bundleObject.toolVersion | Should Be '1.2.0'
            $bundleObject.createdBy | Should Not BeNullOrEmpty
            $bundleObject.hostname | Should Not BeNullOrEmpty
            (@($bundleObject.excludedDirectories) -contains 'skipme') | Should Be $true
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }

    It 'preserves a root directory name when BundleRootName is specified' {
        $workspace = New-TestWorkspace
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'project'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $restorePath = Join-Path -Path $workspace -ChildPath 'restore'
            $bundleResultPath = Join-Path -Path $workspace -ChildPath 'bundle_result.json'
            New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath 'main.ps1') -Value 'Write-Host "root"' -Encoding UTF8

            $bundleRun = Invoke-Converter -Mode 'Bundle' -InputPath $inputPath -OutputPath $outputPath -BundleRootName 'ProjectRoot' -ResultJsonPath $bundleResultPath
            $bundleRun.ExitCode | Should Be 0
            $bundleResult = Get-ResultObject -Path $bundleResultPath

            $restoreRun = Invoke-Converter -Mode 'Restore' -RestoreInputPath $bundleResult.BundlePath -RestoreOutputPath $restorePath
            $restoreRun.ExitCode | Should Be 0
            (Test-Path -LiteralPath (Join-Path -Path $restorePath -ChildPath 'ProjectRoot\main.ps1')) | Should Be $true
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }

    It 'restores from an explicit bundle file path' {
        $workspace = Join-Path -Path $env:TEMP -ChildPath ('bundle_system_test_bang!' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $workspace -Force | Out-Null
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'input'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $restorePath = Join-Path -Path $workspace -ChildPath 'restore'
            $bundleResultPath = Join-Path -Path $workspace -ChildPath 'bundle_result.json'
            $restoreResultPath = Join-Path -Path $workspace -ChildPath 'restore_result.json'
            New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath 'sample.md') -Value "line1`r`nline2" -Encoding UTF8

            $bundleRun = Invoke-Converter -Mode 'Bundle' -InputPath $inputPath -OutputPath $outputPath -ResultJsonPath $bundleResultPath
            $bundleRun.ExitCode | Should Be 0
            $bundleResult = Get-ResultObject -Path $bundleResultPath

            $restoreRun = Invoke-Converter -Mode 'Restore' -RestoreInputPath $bundleResult.BundlePath -RestoreOutputPath $restorePath -ResultJsonPath $restoreResultPath
            $restoreRun.ExitCode | Should Be 0

            $restoredFile = Join-Path -Path $restorePath -ChildPath 'sample.md'
            (Test-Path -LiteralPath $restoredFile) | Should Be $true
            (Get-Content -Raw -LiteralPath $restoredFile) | Should Be "line1`r`nline2`r`n"
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }

    It 'passes arguments through bundle_files.bat and restore_files.bat' {
        $workspace = Join-Path -Path $env:TEMP -ChildPath ('bundle_system_test_bang!_amp&percent%X' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $workspace -Force | Out-Null
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'input'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $restorePath = Join-Path -Path $workspace -ChildPath 'restore'
            $bundleResultPath = Join-Path -Path $workspace -ChildPath 'bundle_result.json'
            $restoreResultPath = Join-Path -Path $workspace -ChildPath 'restore_result.json'
            New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath 'tool.ps1') -Value 'Write-Host "bat"' -Encoding UTF8

            $bundleArgs = @(
                '--no-pause',
                '-InputPath', $inputPath,
                '-OutputPath', $outputPath,
                '-ResultJsonPath', $bundleResultPath
            )
            $bundleRun = Invoke-BatchFile -BatchPath $script:BundleBatPath -Arguments $bundleArgs
            $bundleRun.ExitCode | Should Be 0

            $directOutputPath = Join-Path -Path $workspace -ChildPath 'direct_output'
            $directBundleResultPath = Join-Path -Path $workspace -ChildPath 'direct_bundle_result.json'
            $directBundleOutput = & $script:BundleBatPath --no-pause -InputPath $inputPath -OutputPath $directOutputPath -ResultJsonPath $directBundleResultPath | Out-String
            $LASTEXITCODE | Should Be 0
            (Test-Path -LiteralPath $directBundleResultPath) | Should Be $true

            $bundleResult = Get-ResultObject -Path $directBundleResultPath
            $restoreArgs = @(
                '--no-pause',
                '-RestoreInputPath', $bundleResult.BundlePath,
                '-RestoreOutputPath', $restorePath,
                '-ResultJsonPath', $restoreResultPath
            )
            $restoreRun = Invoke-BatchFile -BatchPath $script:RestoreBatPath -Arguments $restoreArgs
            $restoreRun.ExitCode | Should Be 0

            $directRestorePath = Join-Path -Path $workspace -ChildPath 'direct_restore'
            $directRestoreResultPath = Join-Path -Path $workspace -ChildPath 'direct_restore_result.json'
            $directRestoreOutput = & $script:RestoreBatPath --no-pause -RestoreInputPath $bundleResult.BundlePath -RestoreOutputPath $directRestorePath -ResultJsonPath $directRestoreResultPath | Out-String
            $LASTEXITCODE | Should Be 0
            (Test-Path -LiteralPath (Join-Path -Path $restorePath -ChildPath 'tool.ps1')) | Should Be $true
            (Test-Path -LiteralPath (Join-Path -Path $directRestorePath -ChildPath 'tool.ps1')) | Should Be $true
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }

    It 'round-trips UTF-8 BOM, UTF-8, UTF-16 LE, and CP932 files byte-for-byte' {
        $workspace = New-TestWorkspace
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'input'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $restorePath = Join-Path -Path $workspace -ChildPath 'restore'
            $bundleResultPath = Join-Path -Path $workspace -ChildPath 'bundle_result.json'
            New-Item -ItemType Directory -Path $inputPath -Force | Out-Null

            [System.IO.File]::WriteAllText(
                (Join-Path -Path $inputPath -ChildPath 'utf8bom.ps1'),
                'Write-Host "こんにちは"',
                [System.Text.UTF8Encoding]::new($true)
            )
            [System.IO.File]::WriteAllText(
                (Join-Path -Path $inputPath -ChildPath 'utf8.md'),
                '日本語UTF8',
                [System.Text.UTF8Encoding]::new($false)
            )
            [System.IO.File]::WriteAllText(
                (Join-Path -Path $inputPath -ChildPath 'utf16.json'),
                '{"message":"こんにちは"}',
                [System.Text.UnicodeEncoding]::new($false, $true)
            )
            $cp932Bytes = [System.Text.Encoding]::GetEncoding(932).GetBytes("Attribute VB_Name = ""Mod1""`r`n' こんにちは")
            [System.IO.File]::WriteAllBytes((Join-Path -Path $inputPath -ChildPath 'cp932.bas'), $cp932Bytes)

            $bundleRun = Invoke-Converter -Mode 'Bundle' -InputPath $inputPath -OutputPath $outputPath -ResultJsonPath $bundleResultPath
            $bundleRun.ExitCode | Should Be 0
            $bundleResult = Get-ResultObject -Path $bundleResultPath

            $restoreRun = Invoke-Converter -Mode 'Restore' -RestoreInputPath $bundleResult.BundlePath -RestoreOutputPath $restorePath
            $restoreRun.ExitCode | Should Be 0

            foreach ($fileName in @('utf8bom.ps1', 'utf8.md', 'utf16.json', 'cp932.bas')) {
                $originalBytes = Get-FileBytes -Path (Join-Path -Path $inputPath -ChildPath $fileName)
                $restoredBytes = Get-FileBytes -Path (Join-Path -Path $restorePath -ChildPath $fileName)
                [System.BitConverter]::ToString($restoredBytes) | Should Be ([System.BitConverter]::ToString($originalBytes))
            }
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }

    It 'verifies round-trip output with verify mode' {
        $workspace = New-TestWorkspace
        try {
            $inputPath = Join-Path -Path $workspace -ChildPath 'input'
            $outputPath = Join-Path -Path $workspace -ChildPath 'output'
            $restorePath = Join-Path -Path $workspace -ChildPath 'restore'
            $resultJsonPath = Join-Path -Path $workspace -ChildPath 'verify_result.json'
            $nestedPath = Join-Path -Path $inputPath -ChildPath 'nested'
            New-Item -ItemType Directory -Path $nestedPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $inputPath -ChildPath 'root.ps1') -Value 'Write-Host "root"' -Encoding UTF8
            Set-Content -LiteralPath (Join-Path -Path $nestedPath -ChildPath 'child.md') -Value 'child' -Encoding UTF8

            $verifyRun = Invoke-Converter -Mode 'Verify' -InputPath $inputPath -OutputPath $outputPath -RestoreOutputPath $restorePath -ResultJsonPath $resultJsonPath
            $verifyRun.ExitCode | Should Be 0

            $verifyResult = Get-ResultObject -Path $resultJsonPath
            $verifyResult.Status | Should Be 'Success'
            $verifyResult.VerifiedFileCount | Should Be 2
            $verifyResult.VerifiedDirectoryCount | Should Be 1
            (Test-Path -LiteralPath $verifyResult.VerifyWorkspace) | Should Be $false
        }
        finally {
            Remove-TestWorkspace -Path $workspace
        }
    }
}
