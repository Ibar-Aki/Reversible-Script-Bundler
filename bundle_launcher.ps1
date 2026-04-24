[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Bundle', 'Restore')]
    [string]$DefaultMode,
    [Parameter(Mandatory = $true)]
    [string]$SystemPath,
    [string]$ArgumentFile,
    [string]$BatchPath,
    [string]$ExitParentFlagPath,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardedArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Split-CmdArgumentString {
    param([string]$ArgumentText)

    $result = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $inQuotes = $false

    for ($index = 0; $index -lt $ArgumentText.Length; $index++) {
        $character = $ArgumentText[$index]

        if ($character -eq '"') {
            if ($inQuotes -and $index + 1 -lt $ArgumentText.Length -and $ArgumentText[$index + 1] -eq '"') {
                [void]$current.Append('"')
                $index++
                continue
            }

            $inQuotes = -not $inQuotes
            continue
        }

        if (-not $inQuotes -and [char]::IsWhiteSpace($character)) {
            if ($current.Length -gt 0) {
                [void]$result.Add($current.ToString())
                [void]$current.Clear()
            }
            continue
        }

        [void]$current.Append($character)
    }

    if ($current.Length -gt 0) {
        [void]$result.Add($current.ToString())
    }

    return @($result)
}

function Get-BatchArgumentsFromCmdLine {
    param(
        [string]$CommandLine,
        [string]$CurrentBatchPath
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine) -or [string]::IsNullOrWhiteSpace($CurrentBatchPath)) {
        return $null
    }

    $normalizedBatchPath = [System.IO.Path]::GetFullPath($CurrentBatchPath)
    $quotedNeedle = '"' + $normalizedBatchPath + '"'
    $comparison = [System.StringComparison]::OrdinalIgnoreCase
    $batchIndex = $CommandLine.IndexOf($quotedNeedle, $comparison)
    $needleLength = $quotedNeedle.Length

    if ($batchIndex -lt 0) {
        $batchIndex = $CommandLine.IndexOf($normalizedBatchPath, $comparison)
        $needleLength = $normalizedBatchPath.Length
    }

    if ($batchIndex -lt 0) {
        $batchFileName = [System.IO.Path]::GetFileName($normalizedBatchPath)
        $quotedFileName = '"' + $batchFileName + '"'
        $batchIndex = $CommandLine.IndexOf($quotedFileName, $comparison)
        $needleLength = $quotedFileName.Length

        if ($batchIndex -lt 0) {
            $batchIndex = $CommandLine.IndexOf($batchFileName, $comparison)
            $needleLength = $batchFileName.Length
        }
    }

    if ($batchIndex -lt 0) {
        return $null
    }

    $argumentText = $CommandLine.Substring($batchIndex + $needleLength).Trim()
    if ($argumentText.EndsWith('"')) {
        $quoteCount = 0
        foreach ($character in $argumentText.ToCharArray()) {
            if ($character -eq '"') {
                $quoteCount++
            }
        }

        if (($quoteCount % 2) -eq 1) {
            $argumentText = $argumentText.Substring(0, $argumentText.Length - 1).TrimEnd()
        }
    }

    return @(Split-CmdArgumentString -ArgumentText $argumentText)
}

$usedCmdCommandLine = $false
$shouldExitParentCmd = $false
$mode = $DefaultMode
$shouldPause = $true
$converterArgs = New-Object System.Collections.Generic.List[string]
$rawArgs = @()

$originalCmdCommandLine = $env:BUNDLE_ORIGINAL_CMDCMDLINE
if ([string]::IsNullOrWhiteSpace($originalCmdCommandLine)) {
    $originalCmdCommandLine = $env:CMDCMDLINE
}

$cmdLineArgs = Get-BatchArgumentsFromCmdLine -CommandLine $originalCmdCommandLine -CurrentBatchPath $BatchPath
if ($null -ne $cmdLineArgs) {
    $rawArgs = @($cmdLineArgs)
    $usedCmdCommandLine = $true
    $shouldExitParentCmd = $originalCmdCommandLine -match '(?i)^\s*"?[^"]*cmd(?:\.exe)?"?\s+(?:/[a-z0-9:]+\s+)*?/c\b'
}
elseif (-not [string]::IsNullOrWhiteSpace($ArgumentFile)) {
    if (Test-Path -LiteralPath $ArgumentFile -PathType Leaf) {
        $rawArgs = [System.IO.File]::ReadAllLines($ArgumentFile)
    }
}
else {
    $rawArgs = @($ForwardedArgs)
}

foreach ($argument in @($rawArgs)) {
    if ($DefaultMode -eq 'Bundle' -and $argument -ieq 'verify') {
        $mode = 'Verify'
        continue
    }

    if ($DefaultMode -eq 'Restore' -and $argument -ieq 'structure') {
        $mode = 'RestoreStructure'
        continue
    }

    if ($argument -ieq '--no-pause') {
        $shouldPause = $false
        continue
    }

    [void]$converterArgs.Add($argument)
}

$processArgs = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', $SystemPath,
    '-Mode', $mode
)
$processArgs += @($converterArgs)

& powershell.exe @processArgs
$exitCode = $LASTEXITCODE

if ($shouldExitParentCmd -and -not [string]::IsNullOrWhiteSpace($ExitParentFlagPath)) {
    [System.IO.File]::WriteAllText($ExitParentFlagPath, 'exit-parent', [System.Text.Encoding]::ASCII)
}

if ($shouldPause) {
    Write-Host 'Press any key to continue . . .'
    [void][System.Console]::ReadKey($true)
}

exit $exitCode
