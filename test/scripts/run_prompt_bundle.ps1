[CmdletBinding()]
param(
    [string]$WorkbookJsonPath,
    [string]$JsonlPath,
    [ValidateSet('general', 'mechanical', 'accounting')]
    [string]$Scenario = 'general',
    [string]$OutputDir,
    [int]$MaxChunkPrompts = 3,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

try {
    $latestOutputDir = $null
    if (-not $WorkbookJsonPath) {
        $latestOutputDir = Get-LatestOutputDirectory
        $WorkbookJsonPath = Join-Path $latestOutputDir 'workbook.json'
    }
    elseif (Test-Path -LiteralPath $WorkbookJsonPath) {
        $WorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath
    }

    if (-not $JsonlPath) {
        if (-not [string]::IsNullOrWhiteSpace($WorkbookJsonPath)) {
            $JsonlPath = Join-Path (Split-Path -Path $WorkbookJsonPath -Parent) 'llm_package.jsonl'
        }
        else {
            if ($null -eq $latestOutputDir) {
                $latestOutputDir = Get-LatestOutputDirectory
            }
            $JsonlPath = Join-Path $latestOutputDir 'llm_package.jsonl'
        }
    }
    elseif (Test-Path -LiteralPath $JsonlPath) {
        $JsonlPath = Resolve-AbsolutePath -Path $JsonlPath
    }

    if (-not $OutputDir) {
        $promptSourceDir = if (-not [string]::IsNullOrWhiteSpace($WorkbookJsonPath)) {
            Split-Path -Path $WorkbookJsonPath -Parent
        }
        else {
            Split-Path -Path $JsonlPath -Parent
        }
        $OutputDir = Join-Path $promptSourceDir 'prompt_bundle'
    }

    & (Join-Path $PSScriptRoot 'export_prompt_bundle.ps1') `
        -WorkbookJsonPath $WorkbookJsonPath `
        -JsonlPath $JsonlPath `
        -Scenario $Scenario `
        -OutputDir $OutputDir `
        -MaxChunkPrompts $MaxChunkPrompts `
        -RedactPaths:$RedactPaths

    Write-Host '=== 指示文セット作成結果 ==='
    Write-Host ('  シナリオ:   {0}' -f $Scenario)
    Write-Host ('  出力先:     {0}' -f (Get-NormalizedFullPath -Path $OutputDir))
    Write-NextStepBlock -Steps @(
        ('prompt_*.txt を開いて LLM に貼り付ける'),
        ('テンプレート確認: docs\reference\LLM_PROMPT_FORMATS.md')
    )
}
catch {
    Write-ErrorRecoverySteps -CommandName 'run_prompt_bundle'
    throw "run_prompt_bundle.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
