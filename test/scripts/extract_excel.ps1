[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,
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

function Get-OptionalRangePropertyMatrix {
    param(
        [Parameter(Mandatory)]
        $Range,
        [Parameter(Mandatory)]
        [string]$PropertyName,
        [Parameter(Mandatory)]
        $Warnings,
        [Parameter(Mandatory)]
        [string]$SheetName
    )

    try {
        return ,$Range.$PropertyName
    }
    catch {
        Add-WarningMessage -Warnings $Warnings -Message ("UsedRange.{0} could not be read for {1}: {2}" -f $PropertyName, $SheetName, $_.Exception.Message)
        return $null
    }
}

function Get-NormalizedSheetFilterList {
    param(
        [string[]]$Names
    )

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @($Names)) {
        if ([string]::IsNullOrWhiteSpace([string]$name)) {
            continue
        }

        $segments = ([string]$name) -split ','
        foreach ($segment in $segments) {
            if (-not [string]::IsNullOrWhiteSpace($segment)) {
                [void]$normalized.Add($segment.Trim())
            }
        }
    }

    return [string[]]$normalized.ToArray()
}

if (-not $OutputDir) {
    $OutputDir = Get-DefaultRunOutputDirectory -ExcelPath $ExcelPath
}

if ($SkipStyles) {
    $CollectStyles = $false
}

$warnings = [System.Collections.Generic.List[string]]::new()
$excel = $null
$workbook = $null
$usedRange = $null
$requestedSheets = Get-NormalizedSheetFilterList -Names $Sheets
$excludedSheets = Get-NormalizedSheetFilterList -Names $ExcludeSheets

try {
    $resolvedExcelPath = Resolve-AbsolutePath -Path $ExcelPath
    Ensure-Directory -Path $OutputDir
    $resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir

    $workbookJsonPath = Join-Path $OutputDir 'workbook.json'
    $stylesJsonPath = Join-Path $OutputDir 'styles.json'
    $manifestJsonPath = Join-Path $OutputDir 'manifest.json'
    $preflightScriptPath = Join-Path $PSScriptRoot 'preflight_excel.ps1'

    try {
        & $preflightScriptPath -ExcelPath $resolvedExcelPath -OutputDir $resolvedOutputDir -RedactPaths:$RedactPaths
    }
    catch {
        foreach ($stalePath in @($workbookJsonPath, $stylesJsonPath, $manifestJsonPath)) {
            if (Test-Path -LiteralPath $stalePath) {
                Remove-Item -LiteralPath $stalePath -Force
            }
        }
        throw
    }

    $excel = New-ExcelApplication -AllowWorkbookMacros:$AllowWorkbookMacros
    $workbook = $excel.Workbooks.Open($resolvedExcelPath, 0, $true)
    $sourceSheetCount = [int]$workbook.Worksheets.Count
    if (-not $NoRecalculate) {
        try {
            $excel.CalculateFullRebuild()
        }
        catch {
            Add-WarningMessage -Warnings $warnings -Message "Recalculation failed: $($_.Exception.Message)"
        }
    }

    $sheetEntries = [System.Collections.Generic.List[object]]::new()
    $cells = New-Object System.Collections.Generic.List[object]
    $styles = New-Object System.Collections.Generic.List[object]
    $mergedRanges = New-Object System.Collections.Generic.List[object]
    $globalMergedKeys = [System.Collections.Generic.HashSet[string]]::new()
    $availableSheets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $selectedSheets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $selectedSheetOrder = [System.Collections.Generic.List[string]]::new()

    $totalFormulaCount = 0
    $totalCellCount = 0

    for ($sheetIndex = 1; $sheetIndex -le $sourceSheetCount; $sheetIndex++) {
        $sheet = $null
        try {
            $sheet = $workbook.Worksheets.Item($sheetIndex)
            $sheetName = [string]$sheet.Name
            [void]$availableSheets.Add($sheetName)

            $isIncluded = (@($requestedSheets).Count -eq 0 -or $requestedSheets -contains $sheetName)
            if ($isIncluded -and @($excludedSheets).Count -gt 0 -and ($excludedSheets -contains $sheetName)) {
                $isIncluded = $false
            }

            if ($isIncluded) {
                [void]$selectedSheetOrder.Add($sheetName)
            }
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }
    }

    $selectedSheetCount = $selectedSheetOrder.Count
    $currentSheetNumber = 0

    foreach ($sheet in $workbook.Worksheets) {
        $sheetName = [string]$sheet.Name

        $isIncluded = (@($requestedSheets).Count -eq 0 -or $requestedSheets -contains $sheetName)
        if ($isIncluded -and @($excludedSheets).Count -gt 0 -and ($excludedSheets -contains $sheetName)) {
            $isIncluded = $false
        }

        if (-not $isIncluded) {
            Release-ComReference $sheet
            continue
        }

        $currentSheetNumber++
        Write-Host ('[{0}/{1}] Sheet "{2}" を処理中...' -f $currentSheetNumber, $selectedSheetCount, $sheetName)
        [void]$selectedSheets.Add($sheetName)
        $sheetIndex = [int]$sheet.Index
        $sheetVisible = [int]$sheet.Visible
        $usedRange = $sheet.UsedRange
        $rangeInfo = Get-UsedRangeInfo -UsedRange $usedRange
        $freezePanes = Get-WorksheetFreezeState -Excel $excel -Worksheet $sheet
        $usedRangeValues = Get-OptionalRangePropertyMatrix -Range $usedRange -PropertyName 'Value2' -Warnings $warnings -SheetName $sheetName
        $usedRangeFormulas = Get-OptionalRangePropertyMatrix -Range $usedRange -PropertyName 'Formula' -Warnings $warnings -SheetName $sheetName
        $usedRangeFormula2 = Get-OptionalRangePropertyMatrix -Range $usedRange -PropertyName 'Formula2' -Warnings $warnings -SheetName $sheetName
        $usedRangeNumberFormats = Get-OptionalRangePropertyMatrix -Range $usedRange -PropertyName 'NumberFormat' -Warnings $warnings -SheetName $sheetName
        $usedRowCount = [int]$rangeInfo.row_count
        $usedColumnCount = [int]$rangeInfo.column_count

        $sheetMergedRanges = New-Object System.Collections.Generic.List[object]
        $sheetMergedKeys = [System.Collections.Generic.HashSet[string]]::new()
        $hiddenRows = New-Object System.Collections.Generic.List[int]
        $hiddenColumns = New-Object System.Collections.Generic.List[string]
        $rowHeights = New-Object System.Collections.Generic.List[object]
        $columnWidths = New-Object System.Collections.Generic.List[object]

        for ($rowIndex = $rangeInfo.first_row; $rowIndex -le $rangeInfo.last_row; $rowIndex++) {
            $rowRange = $null
            try {
                $rowRange = $sheet.Rows.Item($rowIndex)
                if ($rowRange.Hidden) {
                    $hiddenRows.Add($rowIndex)
                }
                $rowHeights.Add([ordered]@{
                    row = $rowIndex
                    height = [double]$rowRange.RowHeight
                })
            }
            finally {
                if ($null -ne $rowRange) {
                    Release-ComReference $rowRange
                }
            }
        }

        for ($columnIndex = $rangeInfo.first_column; $columnIndex -le $rangeInfo.last_column; $columnIndex++) {
            $columnRange = $null
            try {
                $columnRange = $sheet.Columns.Item($columnIndex)
                $columnLetters = (Convert-CoordinateToA1 -Row 1 -Column $columnIndex) -replace '\d', ''
                if ($columnRange.Hidden) {
                    $hiddenColumns.Add($columnLetters)
                }
                $columnWidths.Add([ordered]@{
                    column = $columnLetters
                    width = [double]$columnRange.ColumnWidth
                })
            }
            finally {
                if ($null -ne $columnRange) {
                    Release-ComReference $columnRange
                }
            }
        }

        $sheetFormulaCount = 0
        $sheetCellCount = 0

        for ($rowIndex = $rangeInfo.first_row; $rowIndex -le $rangeInfo.last_row; $rowIndex++) {
            for ($columnIndex = $rangeInfo.first_column; $columnIndex -le $rangeInfo.last_column; $columnIndex++) {
                $cell = $null
                $mergeArea = $null
                try {
                    $rowOffset = $rowIndex - $rangeInfo.first_row + 1
                    $columnOffset = $columnIndex - $rangeInfo.first_column + 1
                    $address = Convert-CoordinateToA1 -Row $rowIndex -Column $columnIndex
                    $cell = $sheet.Cells.Item($rowIndex, $columnIndex)
                    $value2 = if ($null -ne $usedRangeValues) {
                        Convert-VariantValue -Value (Get-RangeMatrixValue -Matrix $usedRangeValues -RowOffset $rowOffset -ColumnOffset $columnOffset -RowCount $usedRowCount -ColumnCount $usedColumnCount)
                    }
                    else {
                        Convert-VariantValue -Value $cell.Value2
                    }

                    if ($null -ne $usedRangeFormulas) {
                        $formula = Convert-FormulaValue -Value (Get-RangeMatrixValue -Matrix $usedRangeFormulas -RowOffset $rowOffset -ColumnOffset $columnOffset -RowCount $usedRowCount -ColumnCount $usedColumnCount)
                        $formula2 = Convert-FormulaValue -Value (Get-RangeMatrixValue -Matrix $usedRangeFormula2 -RowOffset $rowOffset -ColumnOffset $columnOffset -RowCount $usedRowCount -ColumnCount $usedColumnCount)
                        $hasFormula = ($null -ne $formula)
                    }
                    else {
                        $hasFormula = [bool]$cell.HasFormula
                        $formula = if ($hasFormula) { [string]$cell.Formula } else { $null }
                        $formula2 = if ($hasFormula) { Get-CellFormula2 -Cell $cell } else { $null }
                    }

                    if ($hasFormula -and $null -eq $formula2) {
                        $formula2 = $formula
                    }

                    $numberFormat = if ($null -ne $usedRangeNumberFormats) {
                        $numberFormatValue = Get-RangeMatrixValue -Matrix $usedRangeNumberFormats -RowOffset $rowOffset -ColumnOffset $columnOffset -RowCount $usedRowCount -ColumnCount $usedColumnCount
                        if ($null -eq $numberFormatValue) { $null } else { [string]$numberFormatValue }
                    }
                    else {
                        [string]$cell.NumberFormat
                    }

                    $mergeAreaAddress = $null
                    $isMergeAnchor = $false

                    if ([bool]$cell.MergeCells) {
                        $mergeArea = $cell.MergeArea
                        $mergeAreaAddress = [string]$mergeArea.Address($false, $false)
                        $isMergeAnchor = ([int]$mergeArea.Row -eq $rowIndex -and [int]$mergeArea.Column -eq $columnIndex)

                        if ($sheetMergedKeys.Add($mergeAreaAddress)) {
                            $anchorAddress = Convert-CoordinateToA1 -Row ([int]$mergeArea.Row) -Column ([int]$mergeArea.Column)
                            $mergeRecord = [ordered]@{
                                sheet = $sheetName
                                range = $mergeAreaAddress
                                anchor = $anchorAddress
                            }
                            $sheetMergedRanges.Add($mergeRecord)
                            if ($globalMergedKeys.Add(("{0}|{1}" -f $sheetName, $mergeAreaAddress))) {
                                $mergedRanges.Add($mergeRecord)
                            }
                        }
                    }

                    $cellRecord = [ordered]@{
                        sheet = $sheetName
                        address = $address
                        row = $rowIndex
                        column = $columnIndex
                        value2 = $value2
                        text = [string]$cell.Text
                        formula = $formula
                        formula2 = $formula2
                        has_formula = $hasFormula
                        number_format = $numberFormat
                        merge_area = $mergeAreaAddress
                        is_merge_anchor = $isMergeAnchor
                        comment = Get-CellCommentText -Cell $cell
                        comment_threaded = Get-CellThreadedComment -Cell $cell
                        hyperlink = Get-CellHyperlink -Cell $cell
                    }

                    $cells.Add($cellRecord)
                    $sheetCellCount++
                    $totalCellCount++
                    if ($hasFormula) {
                        $sheetFormulaCount++
                        $totalFormulaCount++
                    }

                    if ($CollectStyles) {
                        try {
                            $styleRecord = Get-StyleRecord -Cell $cell
                            $styles.Add([ordered]@{
                                sheet = $sheetName
                                address = $address
                                fill_color = $styleRecord.fill_color
                                font_color = $styleRecord.font_color
                                horizontal_alignment = $styleRecord.horizontal_alignment
                                vertical_alignment = $styleRecord.vertical_alignment
                                wrap_text = $styleRecord.wrap_text
                                borders = $styleRecord.borders
                            })
                        }
                        catch {
                            Add-WarningMessage -Warnings $warnings -Message ("styles.json export skipped for {0}!{1}: {2}" -f $sheetName, $address, $_.Exception.Message)
                        }
                    }
                }
                finally {
                    if ($null -ne $mergeArea) {
                        Release-ComReference $mergeArea
                    }
                    if ($null -ne $cell) {
                        Release-ComReference $cell
                    }
                }
            }
        }

        $sheetEntries.Add([ordered]@{
            sheet_name = $sheetName
            sheet_index = $sheetIndex
            visible = $sheetVisible
            used_range = $rangeInfo
            freeze_panes = $freezePanes
            hidden_rows = $hiddenRows
            hidden_columns = $hiddenColumns
            row_heights = $rowHeights
            column_widths = $columnWidths
            cell_count = $sheetCellCount
            formula_count = $sheetFormulaCount
            merged_ranges = $sheetMergedRanges
        })

        if ($null -ne $usedRange) {
            Release-ComReference $usedRange
            $usedRange = $null
        }

        Release-ComReference $sheet
    }

    foreach ($requestedSheet in $requestedSheets) {
        if (-not $availableSheets.Contains($requestedSheet)) {
            Add-WarningMessage -Warnings $warnings -Message ("Requested sheet was not found: {0}" -f $requestedSheet)
        }
    }

    foreach ($excludedSheet in $excludedSheets) {
        if (-not $availableSheets.Contains($excludedSheet)) {
            Add-WarningMessage -Warnings $warnings -Message ("Excluded sheet was not found: {0}" -f $excludedSheet)
        }
    }

    $workbookPayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        workbook = [ordered]@{
            name = [string]$workbook.Name
            path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedExcelPath) } else { $resolvedExcelPath }
            extension = [System.IO.Path]::GetExtension($resolvedExcelPath)
            sheet_count = $sheetEntries.Count
            has_vba = @('.xlsm', '.xlam') -contains ([System.IO.Path]::GetExtension($resolvedExcelPath).ToLowerInvariant())
        }
        sheets = $sheetEntries
        cells = $cells
        merged_ranges = $mergedRanges
    }

    $stylePayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        styles = $styles
    }

    $styleStatus = if (-not $CollectStyles) { 'skipped' } elseif ($styles.Count -gt 0) { 'generated' } else { 'empty' }
    $status = if ($warnings.Count -gt 0) { 'warning' } else { 'success' }

    $manifestPayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        status = $status
        warnings = $warnings
        workbook_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedExcelPath) } else { $resolvedExcelPath }
        output_directory = if ($RedactPaths) { [string](Split-Path -Path $resolvedOutputDir -Leaf) } else { $resolvedOutputDir }
        source_sheet_count = $sourceSheetCount
        sheet_count = $sheetEntries.Count
        cell_count = $totalCellCount
        formula_count = $totalFormulaCount
        merged_range_count = $mergedRanges.Count
        style_export_status = $styleStatus
        verify_status = 'not_run'
        sheet_filter = [ordered]@{
            include = @($requestedSheets)
            exclude = @($excludedSheets)
            selected = @($selectedSheets | Sort-Object)
        }
    }

    Write-JsonFile -Data $workbookPayload -Path $workbookJsonPath
    Write-JsonFile -Data $stylePayload -Path $stylesJsonPath
    Write-JsonFile -Data $manifestPayload -Path $manifestJsonPath
    Set-LatestOutputDirectory -OutputDir $resolvedOutputDir

    $warningSummary = if ($warnings.Count -eq 0) { 'なし' } else { [string]$warnings.Count }
    Write-Host '=== Excel2LLM 抽出結果 ==='
    Write-Host ('  対象ファイル: {0}' -f [System.IO.Path]::GetFileName($resolvedExcelPath))
    Write-Host ('  処理シート:  {0} / {1}' -f $sheetEntries.Count, $sourceSheetCount)
    Write-Host ('  セル数:      {0}' -f $totalCellCount)
    Write-Host ('  数式数:      {0}' -f $totalFormulaCount)
    Write-Host ('  結合セル:    {0}' -f $mergedRanges.Count)
    Write-Host ('  警告:        {0}' -f $warningSummary)
    Write-Host ('  出力先:      {0}' -f $workbookJsonPath)
    Write-NextStepBlock -Steps @(
        ('tools\advanced\run_pack.bat "{0}"' -f $workbookJsonPath),
        ('重要な資料なら tools\advanced\run_verify.bat "{0}" -WorkbookJsonPath "{1}"' -f $resolvedExcelPath, $workbookJsonPath)
    )
}
catch {
    Write-ErrorRecoverySteps -CommandName 'extract'
    throw "extract_excel.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
finally {
    if ($null -ne $usedRange) {
        Release-ComReference $usedRange
    }
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
