[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,
    [string]$WorkbookJsonPath,
    [string]$OutputDir,
    [switch]$AllowWorkbookMacros,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $WorkbookJsonPath) {
    $WorkbookJsonPath = Join-Path (Get-LatestOutputDirectory) 'workbook.json'
}

$resolvedExcelPath = Resolve-AbsolutePath -Path $ExcelPath
$resolvedWorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath

if (-not $OutputDir) {
    $OutputDir = Split-Path -Path $resolvedWorkbookJsonPath -Parent
}

Ensure-Directory -Path $OutputDir

$verifyReportPath = Join-Path $OutputDir 'verify_report.json'
$manifestPath = Join-Path $OutputDir 'manifest.json'

$workbookData = Get-Content -LiteralPath $resolvedWorkbookJsonPath -Raw | ConvertFrom-Json
$warnings = [System.Collections.Generic.List[string]]::new()
$mismatches = [System.Collections.Generic.List[object]]::new()
$excel = $null
$workbook = $null
$cellsBySheet = Group-CellsBySheet -Cells @($workbookData.cells)

try {
    $excel = New-ExcelApplication -AllowWorkbookMacros:$AllowWorkbookMacros
    $workbook = $excel.Workbooks.Open($resolvedExcelPath, 0, $true)
    $excel.CalculateFullRebuild()

    foreach ($sheetName in ($cellsBySheet.Keys | Sort-Object)) {
        $sheet = $null
        try {
            $sheet = $workbook.Worksheets.Item([string]$sheetName)

            foreach ($cellRecord in $cellsBySheet[$sheetName]) {
                $cell = $null
                try {
                    $cell = $sheet.Cells.Item([int]$cellRecord.row, [int]$cellRecord.column)

                    $liveFormula = if ([bool]$cell.HasFormula) { [string]$cell.Formula } else { $null }
                    $liveFormula2 = if ([bool]$cell.HasFormula) { Get-CellFormula2 -Cell $cell } else { $null }
                    $liveValue2 = Convert-VariantValue -Value $cell.Value2
                    $liveText = [string]$cell.Text

                    $expectedValue2 = if ($null -eq $cellRecord.value2) { $null } else { [string]$cellRecord.value2 }
                    $actualValue2 = if ($null -eq $liveValue2) { $null } else { [string]$liveValue2 }

                    if (($cellRecord.formula -ne $liveFormula) -or
                        ([string]$cellRecord.formula2 -ne [string]$liveFormula2) -or
                        ([string]$cellRecord.text -ne [string]$liveText) -or
                        ($expectedValue2 -ne $actualValue2)) {
                        [void]$mismatches.Add([ordered]@{
                            sheet = [string]$cellRecord.sheet
                            address = [string]$cellRecord.address
                            expected = [ordered]@{
                                formula = $cellRecord.formula
                                formula2 = $cellRecord.formula2
                                value2 = $cellRecord.value2
                                text = $cellRecord.text
                            }
                            actual = [ordered]@{
                                formula = $liveFormula
                                formula2 = $liveFormula2
                                value2 = $liveValue2
                                text = $liveText
                            }
                        })
                    }
                }
                catch {
                    Add-WarningMessage -Warnings $warnings -Message ("Cell verification failed for {0}!{1}: {2}" -f [string]$cellRecord.sheet, [string]$cellRecord.address, $_.Exception.Message)
                }
                finally {
                    if ($null -ne $cell) {
                        Release-ComReference $cell
                    }
                }
            }
        }
        catch {
            Add-WarningMessage -Warnings $warnings -Message ("Sheet verification failed for {0}: {1}" -f [string]$sheetName, $_.Exception.Message)
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }        
    }

    $verifyStatus = if ($mismatches.Count -gt 0 -or $warnings.Count -gt 0) { 'warning' } else { 'success' }
    $report = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Verifier'
        status = $verifyStatus
        workbook_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedExcelPath) } else { $resolvedExcelPath }
        workbook_json_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedWorkbookJsonPath) } else { $resolvedWorkbookJsonPath }
        mismatch_count = $mismatches.Count
        warnings = $warnings
        mismatches = $mismatches
    }
    Write-JsonFile -Data $report -Path $verifyReportPath

    if (Test-Path -LiteralPath $manifestPath) {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        $mergedWarnings = [System.Collections.Generic.List[string]]::new()

        if ($null -ne $manifest.warnings) {
            foreach ($existingWarning in $manifest.warnings) {
                $mergedWarnings.Add([string]$existingWarning)
            }
        }

        foreach ($warning in $warnings) {
            $mergedWarnings.Add([string]$warning)
        }

        if ($mismatches.Count -gt 0) {
            $mergedWarnings.Add("$($mismatches.Count) mismatch entries were written to verify_report.json.")
        }

        $manifest.warnings = $mergedWarnings
        $manifest.verify_status = $verifyStatus
        if ($verifyStatus -eq 'warning') {
            $manifest.status = 'warning'
        }

        Write-JsonFile -Data $manifest -Path $manifestPath
    }

    $warningSummary = if ($warnings.Count -eq 0) { 'なし' } else { [string]$warnings.Count }
    Write-Host '=== 検証結果 ==='
    if ($mismatches.Count -eq 0) {
        Write-Host ('  差分:        なし (mismatch_count: {0})' -f $mismatches.Count)
    }
    else {
        Write-Host ('  差分:        あり (mismatch_count: {0})' -f $mismatches.Count)
    }
    Write-Host ('  警告:        {0}' -f $warningSummary)
    Write-Host ('  詳細:        {0}' -f $verifyReportPath)
    if ($mismatches.Count -eq 0 -and $warnings.Count -eq 0) {
        Write-NextStepBlock -Steps @(
            ('tools\advanced\run_pack.bat "{0}"' -f $resolvedWorkbookJsonPath)
        )
    }
    else {
        Write-NextStepBlock -Steps @(
            ('verify_report.json を確認する: {0}' -f $verifyReportPath),
            ('差分解消後に tools\advanced\run_pack.bat "{0}"' -f $resolvedWorkbookJsonPath)
        )
    }
}
catch {
    Write-ErrorRecoverySteps -CommandName 'verify'
    throw "excel_verify.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
        }
        Release-ComReference $workbook
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        Release-ComReference $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
