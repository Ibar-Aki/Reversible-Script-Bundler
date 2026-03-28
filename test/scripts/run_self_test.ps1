[CmdletBinding()]
param()

. (Join-Path $PSScriptRoot 'common.ps1')

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
$samplesDir = Join-Path $projectRoot 'samples'
$outputDir = Join-Path $projectRoot 'output'
$rebuiltDir = Join-Path $outputDir 'rebuilt'
$roundTripDir = Join-Path $outputDir 'roundtrip'

& (Join-Path $PSScriptRoot 'create_sample_workbook.ps1') -OutputDir $samplesDir
& (Join-Path $PSScriptRoot 'extract_excel.ps1') -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $outputDir
& (Join-Path $PSScriptRoot 'pack_for_llm.ps1') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath (Join-Path $outputDir 'llm_package.jsonl') -ChunkBy range -MaxCells 250
& (Join-Path $PSScriptRoot 'excel_verify.ps1') -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputDir $outputDir
& (Join-Path $PSScriptRoot 'rebuild_excel.ps1') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath (Join-Path $rebuiltDir 'sample_rebuilt.xlsx') -Overwrite
& (Join-Path $PSScriptRoot 'extract_excel.ps1') -ExcelPath (Join-Path $rebuiltDir 'sample_rebuilt.xlsx') -OutputDir $roundTripDir

$workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json
$verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json
$jsonlLines = @(Get-Content -LiteralPath (Join-Path $outputDir 'llm_package.jsonl'))
$rebuildReport = Get-Content -LiteralPath (Join-Path $rebuiltDir 'rebuild_report.json') -Raw | ConvertFrom-Json
$roundTripWorkbookJson = Get-Content -LiteralPath (Join-Path $roundTripDir 'workbook.json') -Raw | ConvertFrom-Json

if ($workbookJson.sheets.Count -lt 3) {
    throw 'Expected at least three worksheets in the sample workbook.'
}

if (($workbookJson.cells | Where-Object { $_.has_formula }).Count -lt 3) {
    throw 'Expected formula cells were not extracted.'
}

$firstFormulaCell = $workbookJson.cells | Where-Object { $_.has_formula } | Select-Object -First 1
if ($null -eq $firstFormulaCell.PSObject.Properties['formula2']) {
    throw 'formula2 field is missing from extracted cells.'
}

if (($workbookJson.merged_ranges | Where-Object { $_.sheet -eq 'Summary' }).Count -lt 1) {
    throw 'Merged range extraction failed.'
}

if ($jsonlLines.Count -lt 2) {
    throw 'Expected multiple JSONL chunks.'
}

if (-not @('success', 'warning') -contains [string]$manifest.status) {
    throw 'manifest.json status is invalid.'
}

if (-not @('success', 'warning') -contains [string]$verifyReport.status) {
    throw 'verify_report.json status is invalid.'
}

if (-not (Test-Path -LiteralPath (Join-Path $rebuiltDir 'sample_rebuilt.xlsx'))) {
    throw 'Rebuilt workbook was not created.'
}

if ($rebuildReport.restored_sheets -lt 3) {
    throw 'Rebuild report did not record all worksheets.'
}

if ($roundTripWorkbookJson.workbook.sheet_count -ne $workbookJson.workbook.sheet_count) {
    throw 'Round-trip workbook sheet count does not match the source.'
}

Write-Host 'セルフテストが正常終了しました。'
