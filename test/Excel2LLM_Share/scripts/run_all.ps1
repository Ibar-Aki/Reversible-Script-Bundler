[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$ExcelPath,
    [switch]$Verify,
    [string]$OutputDir,
    [string[]]$Sheets,
    [string[]]$ExcludeSheets,
    [switch]$CollectStyles,
    [switch]$SkipStyles,
    [switch]$NoRecalculate,
    [switch]$RedactPaths,
    [switch]$AllowWorkbookMacros
)

. (Join-Path $PSScriptRoot 'common.ps1')

$sw = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $projectRoot = Get-ProjectRoot
    $resolvedExcelPath = Resolve-AbsolutePath -Path $ExcelPath
    if (-not $OutputDir) {
        $OutputDir = Get-DefaultRunOutputDirectory -ExcelPath $resolvedExcelPath
    }
    $resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
    $workbookJsonPath = Join-Path $resolvedOutputDir 'workbook.json'
    $jsonlPath = Join-Path $resolvedOutputDir 'llm_package.jsonl'
    $extractParameters = @{
        ExcelPath = $resolvedExcelPath
        OutputDir = $resolvedOutputDir
        CollectStyles = $CollectStyles
        SkipStyles = $SkipStyles
        NoRecalculate = $NoRecalculate
        RedactPaths = $RedactPaths
        AllowWorkbookMacros = $AllowWorkbookMacros
    }
    if (@($Sheets).Count -gt 0) {
        $extractParameters['Sheets'] = $Sheets
    }
    if (@($ExcludeSheets).Count -gt 0) {
        $extractParameters['ExcludeSheets'] = $ExcludeSheets
    }

    # extract_excel.ps1 が mandatory preflight を先に実行し、危険なブックでは COM を起動しない。
    & (Join-Path $PSScriptRoot 'extract_excel.ps1') @extractParameters

    if ($Verify) {
        & (Join-Path $PSScriptRoot 'excel_verify.ps1') `
            -ExcelPath $resolvedExcelPath `
            -WorkbookJsonPath $workbookJsonPath `
            -OutputDir $resolvedOutputDir `
            -AllowWorkbookMacros:$AllowWorkbookMacros `
            -RedactPaths:$RedactPaths
    }

    & (Join-Path $PSScriptRoot 'pack_for_llm.ps1') -WorkbookJsonPath $workbookJsonPath -OutputPath $jsonlPath

    $sw.Stop()
    Write-Host '=== 一括実行結果 ==='
    Write-Host ('  対象ファイル: {0}' -f [System.IO.Path]::GetFileName($resolvedExcelPath))
    Write-Host ('  workbook.json: {0}' -f $workbookJsonPath)
    Write-Host ('  llm_package.jsonl: {0}' -f $jsonlPath)
    Write-Host ('  verify 実行: {0}' -f $(if ($Verify) { 'あり' } else { 'なし' }))
    Write-Host ('  処理時間: {0:hh\:mm\:ss}' -f $sw.Elapsed)
    Write-NextStepBlock -Steps @(
        ('LLM に渡す対象: {0}' -f $jsonlPath),
        ('必要なら Excel2LLM.bat -PromptBundle -Scenario general を実行する')
    )
}
catch {
    Write-ErrorRecoverySteps -CommandName 'run_all'
    throw "run_all.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
