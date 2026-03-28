[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookJsonPath,
    [string]$StylesJsonPath,
    [string]$OutputPath,
    [switch]$Overwrite
)

. (Join-Path $PSScriptRoot 'common.ps1')

function Convert-JsonValueToExcelValue {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParseExact(
                $Value,
                'o',
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind,
                [ref]$parsedDate)) {
            return $parsedDate
        }
    }

    return $Value
}

function Convert-ThreadedCommentToPlainText {
    param(
        $ThreadedComment
    )

    if ($null -eq $ThreadedComment) {
        return $null
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace([string]$ThreadedComment.text)) {
        [void]$lines.Add([string]$ThreadedComment.text)
    }

    foreach ($reply in @($ThreadedComment.replies)) {
        $author = if ([string]::IsNullOrWhiteSpace([string]$reply.author)) { 'reply' } else { [string]$reply.author }
        $replyText = [string]$reply.text
        if (-not [string]::IsNullOrWhiteSpace($replyText)) {
            [void]$lines.Add(('{0}: {1}' -f $author, $replyText))
        }
    }

    if ($lines.Count -eq 0) {
        return $null
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-ResolvedStylesJsonPath {
    param(
        [Parameter(Mandatory)]
        [string]$ResolvedWorkbookJsonPath,
        [string]$StylesJsonPath
    )

    if (-not [string]::IsNullOrWhiteSpace($StylesJsonPath)) {
        return Resolve-AbsolutePath -Path $StylesJsonPath
    }

    $candidate = Join-Path (Split-Path -Path $ResolvedWorkbookJsonPath -Parent) 'styles.json'
    if (Test-Path -LiteralPath $candidate) {
        return Resolve-AbsolutePath -Path $candidate
    }

    return $null
}

function Get-DefaultOutputPath {
    param(
        [Parameter(Mandatory)]
        [string]$ResolvedWorkbookJsonPath,
        [Parameter(Mandatory)]
        $WorkbookMetadata
    )

    $rebuiltDir = Join-Path (Split-Path -Path $ResolvedWorkbookJsonPath -Parent) 'rebuilt'
    Ensure-Directory -Path $rebuiltDir

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension([string]$WorkbookMetadata.name)
    return Join-Path $rebuiltDir ($baseName + '.xlsx')
}

function Test-ValidMergeAddress {
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    return [System.Text.RegularExpressions.Regex]::IsMatch($Address.Trim().ToUpperInvariant(), '^[A-Z]+\d+:[A-Z]+\d+$')
}

function Test-IsDefaultNumberFormat {
    param(
        [string]$NumberFormat
    )

    if ([string]::IsNullOrWhiteSpace($NumberFormat)) {
        return $true
    }

    $normalized = $NumberFormat.Trim().ToUpperInvariant()
    return @('GENERAL', 'G/標準', '標準') -contains $normalized
}

function Set-WorksheetBaseValues {
    param(
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        $SheetRecord,
        [Parameter(Mandatory)]
        [object[]]$SheetCells
    )

    $rangeInfo = $SheetRecord.used_range
    if ($null -eq $rangeInfo) {
        return
    }

    $rowCount = [int]$rangeInfo.row_count
    $columnCount = [int]$rangeInfo.column_count
    if ($rowCount -lt 1 -or $columnCount -lt 1) {
        return
    }

    $values = New-Object 'object[,]' $rowCount, $columnCount
    foreach ($cellRecord in $SheetCells) {
        $rowOffset = [int]$cellRecord.row - [int]$rangeInfo.first_row
        $columnOffset = [int]$cellRecord.column - [int]$rangeInfo.first_column
        if ($rowOffset -lt 0 -or $columnOffset -lt 0 -or $rowOffset -ge $rowCount -or $columnOffset -ge $columnCount) {
            continue
        }

        $values[$rowOffset, $columnOffset] = Convert-JsonValueToExcelValue -Value $cellRecord.value2
    }

    $targetRange = $null
    try {
        $targetRange = $Worksheet.Range([string]$rangeInfo.address)
        $targetRange.Value2 = $values
    }
    finally {
        if ($null -ne $targetRange) {
            Release-ComReference $targetRange
        }
    }
}

function Set-CellFormulaOrValue {
    param(
        [Parameter(Mandatory)]
        $Cell,
        [Parameter(Mandatory)]
        $CellRecord,
        [Parameter(Mandatory)]
        $Warnings,
        [Parameter(Mandatory)]
        [string]$CellLabel
    )

    if (-not [bool]$CellRecord.has_formula) {
        $excelValue = Convert-JsonValueToExcelValue -Value $CellRecord.value2
        if ($null -eq $excelValue) {
            $Cell.ClearContents()
        }
        else {
            $Cell.Value2 = $excelValue
        }
        return 'value2'
    }

    $formula2 = [string]$CellRecord.formula2
    $formula = [string]$CellRecord.formula
    $preferFormula2 = (-not [string]::IsNullOrWhiteSpace($formula2)) -and ($formula2 -ne $formula)

    if ($preferFormula2) {
        try {
            $Cell.Formula2 = $formula2
            return 'formula2'
        }
        catch {
            Add-WarningMessage -Warnings $Warnings -Message ("Formula2 restore failed for {0}: {1}" -f $CellLabel, $_.Exception.Message)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($formula)) {
        try {
            $Cell.Formula = $formula
            return 'formula'
        }
        catch {
            Add-WarningMessage -Warnings $Warnings -Message ("Formula restore failed for {0}: {1}" -f $CellLabel, $_.Exception.Message)
        }
    }

    if (-not $preferFormula2 -and -not [string]::IsNullOrWhiteSpace($formula2)) {
        try {
            $Cell.Formula2 = $formula2
            return 'formula2'
        }
        catch {
            Add-WarningMessage -Warnings $Warnings -Message ("Formula2 restore failed for {0}: {1}" -f $CellLabel, $_.Exception.Message)
        }
    }

    Add-WarningMessage -Warnings $Warnings -Message ("Formula metadata missing or invalid for {0}; restored cached value instead." -f $CellLabel)
    $excelValue = Convert-JsonValueToExcelValue -Value $CellRecord.value2
    if ($null -eq $excelValue) {
        $Cell.ClearContents()
    }
    else {
        $Cell.Value2 = $excelValue
    }
    return 'value2'
}

function Apply-CellStyleRecord {
    param(
        [Parameter(Mandatory)]
        $Cell,
        [Parameter(Mandatory)]
        $StyleRecord
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$StyleRecord.fill_color)) {
        $Cell.Interior.Color = Convert-HexColorToExcelColor -Color ([string]$StyleRecord.fill_color)
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$StyleRecord.font_color)) {
        $Cell.Font.Color = Convert-HexColorToExcelColor -Color ([string]$StyleRecord.font_color)
    }

    if ($null -ne $StyleRecord.horizontal_alignment) {
        $Cell.HorizontalAlignment = [int]$StyleRecord.horizontal_alignment
    }

    if ($null -ne $StyleRecord.vertical_alignment) {
        $Cell.VerticalAlignment = [int]$StyleRecord.vertical_alignment
    }

    if ($null -ne $StyleRecord.wrap_text) {
        $Cell.WrapText = [bool]$StyleRecord.wrap_text
    }

    foreach ($pair in (Get-BorderNames).GetEnumerator()) {
        $borderRecord = $StyleRecord.borders.$($pair.Key)
        if ($null -eq $borderRecord) {
            continue
        }

        $border = $null
        try {
            $border = $Cell.Borders.Item($pair.Value)
            if ($null -ne $borderRecord.line_style) {
                $border.LineStyle = [int]$borderRecord.line_style
            }
            if ($null -ne $borderRecord.weight) {
                $border.Weight = [int]$borderRecord.weight
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$borderRecord.color)) {
                $border.Color = Convert-HexColorToExcelColor -Color ([string]$borderRecord.color)
            }
        }
        finally {
            if ($null -ne $border) {
                Release-ComReference $border
            }
        }
    }
}

function Assert-RebuildInput {
    param(
        [Parameter(Mandatory)]
        $WorkbookData
    )

    if ($null -eq $WorkbookData.workbook) {
        throw 'workbook.json is missing workbook metadata.'
    }

    $sheetRecords = @($WorkbookData.sheets)
    if ($sheetRecords.Count -eq 0) {
        throw 'workbook.json does not contain any sheets.'
    }

    $sheetNames = @{}
    $sheetIndexes = @{}
    foreach ($sheet in $sheetRecords) {
        $sheetName = [string]$sheet.sheet_name
        if ([string]::IsNullOrWhiteSpace($sheetName)) {
            throw 'Sheet name is missing in workbook.json.'
        }

        if ($sheetNames.ContainsKey($sheetName)) {
            throw "Duplicate sheet name found in workbook.json: $sheetName"
        }

        $sheetNames[$sheetName] = $true

        $sheetIndex = [int]$sheet.sheet_index
        if ($sheetIndexes.ContainsKey($sheetIndex)) {
            throw "Duplicate sheet_index found in workbook.json: $sheetIndex"
        }

        $sheetIndexes[$sheetIndex] = $true
    }

    $cellLookup = @{}
    foreach ($cell in @($WorkbookData.cells)) {
        $sheetName = [string]$cell.sheet
        if (-not $sheetNames.ContainsKey($sheetName)) {
            throw "Cell references unknown sheet: $sheetName"
        }

        if ([string]::IsNullOrWhiteSpace([string]$cell.address)) {
            throw "Cell address is missing for sheet: $sheetName"
        }

        if (-not $cellLookup.ContainsKey($sheetName)) {
            $cellLookup[$sheetName] = @{}
        }

        $address = [string]$cell.address
        if ($cellLookup[$sheetName].ContainsKey($address)) {
            throw "Duplicate cell address found in workbook.json: $sheetName!$address"
        }

        $cellLookup[$sheetName][$address] = $true
    }
}

$warnings = [System.Collections.Generic.List[string]]::new()
$excel = $null
$workbook = $null

try {
    $resolvedWorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath
    $resolvedStylesJsonPath = Get-ResolvedStylesJsonPath -ResolvedWorkbookJsonPath $resolvedWorkbookJsonPath -StylesJsonPath $StylesJsonPath
    $workbookData = Get-Content -LiteralPath $resolvedWorkbookJsonPath -Raw | ConvertFrom-Json
    Assert-RebuildInput -WorkbookData $workbookData

    if (-not $OutputPath) {
        $OutputPath = Get-DefaultOutputPath -ResolvedWorkbookJsonPath $resolvedWorkbookJsonPath -WorkbookMetadata $workbookData.workbook
    }

    $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = Split-Path -Path $resolvedOutputPath -Parent
    Ensure-Directory -Path $outputDirectory

    if ((Test-Path -LiteralPath $resolvedOutputPath) -and (-not $Overwrite)) {
        throw "Output file already exists: $resolvedOutputPath"
    }

    $rebuildReportPath = Join-Path $outputDirectory 'rebuild_report.json'

    $styleLookup = @{}
    if ($resolvedStylesJsonPath) {
        $styleData = Get-Content -LiteralPath $resolvedStylesJsonPath -Raw | ConvertFrom-Json
        if ($null -ne $styleData.styles) {
            $styleLookup = Get-StyleLookupBySheetAndAddress -Styles @($styleData.styles)
        }
    }

    $sheetRecords = @($workbookData.sheets | Sort-Object { [int]$_.sheet_index })
    $sheetLookup = Get-SheetLookupByName -Sheets $sheetRecords
    $cellsBySheet = Group-CellsBySheet -Cells @($workbookData.cells)
    $cellLookup = Get-CellLookupBySheetAndAddress -Cells @($workbookData.cells)

    $reportCounters = [ordered]@{
        restored_sheets = 0
        restored_cells = 0
        restored_formulas = 0
        restored_comments = 0
        restored_hyperlinks = 0
        restored_styles = 0
        restored_merged_ranges = 0
        threaded_comment_fallbacks = 0
    }

    if ([bool]$workbookData.workbook.has_vba) {
        Add-WarningMessage -Warnings $warnings -Message 'VBA modules and macros are not restored. The rebuilt workbook is saved as .xlsx only.'
    }

    $excel = New-ExcelApplication
    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt $sheetRecords.Count) {
        $createdSheet = $null
        try {
            $createdSheet = $workbook.Worksheets.Add()
        }
        finally {
            if ($null -ne $createdSheet) {
                Release-ComReference $createdSheet
            }
        }
    }

    while ($workbook.Worksheets.Count -gt $sheetRecords.Count) {
        $extraSheet = $null
        try {
            $extraSheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
            $extraSheet.Delete()
        }
        finally {
            if ($null -ne $extraSheet) {
                Release-ComReference $extraSheet
            }
        }
    }

    for ($worksheetIndex = 1; $worksheetIndex -le $workbook.Worksheets.Count; $worksheetIndex++) {
        $sheet = $null
        try {
            $sheet = $workbook.Worksheets.Item($worksheetIndex)
            $sheet.Name = ('Excel2LLM_Temp_{0}' -f $worksheetIndex)
            $sheet.Visible = -1
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }
    }

    foreach ($sheetRecord in $sheetRecords) {
        $sheet = $null
        $sheetName = [string]$sheetRecord.sheet_name
        try {
            $sheet = $workbook.Worksheets.Item([int]$sheetRecord.sheet_index)
            $sheet.Name = $sheetName
            $reportCounters.restored_sheets++
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }
    }

    foreach ($sheetRecord in $sheetRecords) {
        $sheet = $null
        try {
            $sheetName = [string]$sheetRecord.sheet_name
            Write-Host ('シートを復元中 -> {0}' -f $sheetName)
            $sheet = $workbook.Worksheets.Item([int]$sheetRecord.sheet_index)
            $sheetCells = @($cellsBySheet[$sheetName] | Sort-Object row, column)
            Write-Host ('  基本値を書き込み中 -> {0}' -f $sheetName)
            Set-WorksheetBaseValues -Worksheet $sheet -SheetRecord $sheetRecord -SheetCells $sheetCells
            $reportCounters.restored_cells += $sheetCells.Count

            $specialCells = @(
                $sheetCells | Where-Object {
                    [bool]$_.has_formula -or
                    -not (Test-IsDefaultNumberFormat -NumberFormat ([string]$_.number_format)) -or
                    $null -ne $_.hyperlink -or
                    -not [string]::IsNullOrWhiteSpace([string]$_.comment) -or
                    $null -ne $_.comment_threaded
                }
            )
            Write-Host ('  数式・コメント・リンクを復元中 -> {0}' -f $sheetName)

            foreach ($cellRecord in $specialCells) {
                $cell = $null
                $cellLabel = '{0}!{1}' -f $sheetName, [string]$cellRecord.address
                try {
                    if ($sheetCells.Count -le 100) {
                        Write-Host ('    特別処理セル -> {0}' -f $cellLabel)
                    }
                    $cell = $sheet.Cells.Item([int]$cellRecord.row, [int]$cellRecord.column)

                    if ([bool]$cellRecord.has_formula) {
                        $restoreMode = Set-CellFormulaOrValue -Cell $cell -CellRecord $cellRecord -Warnings $warnings -CellLabel $cellLabel
                        if ($restoreMode -ne 'value2') {
                            $reportCounters.restored_formulas++
                        }
                    }

                    if (-not (Test-IsDefaultNumberFormat -NumberFormat ([string]$cellRecord.number_format))) {
                        $cell.NumberFormat = [string]$cellRecord.number_format
                    }

                    if ($null -ne $cellRecord.hyperlink -and
                        (-not [string]::IsNullOrWhiteSpace([string]$cellRecord.hyperlink.address) -or
                         -not [string]::IsNullOrWhiteSpace([string]$cellRecord.hyperlink.sub_address))) {
                        try {
                            $address = [string]$cellRecord.hyperlink.address
                            $subAddress = [string]$cellRecord.hyperlink.sub_address
                            if (-not [string]::IsNullOrWhiteSpace($address) -and -not [string]::IsNullOrWhiteSpace($subAddress)) {
                                [void]$sheet.Hyperlinks.Add($cell, $address, $subAddress)
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($address)) {
                                [void]$sheet.Hyperlinks.Add($cell, $address)
                            }
                            else {
                                [void]$sheet.Hyperlinks.Add($cell, '', $subAddress)
                            }

                            if (-not [string]::IsNullOrWhiteSpace([string]$cellRecord.hyperlink.text_to_display)) {
                                $cell.Value2 = [string]$cellRecord.hyperlink.text_to_display
                            }

                            $reportCounters.restored_hyperlinks++
                        }
                        catch {
                            Add-WarningMessage -Warnings $warnings -Message ("Hyperlink restore failed for {0}: {1}" -f $cellLabel, $_.Exception.Message)
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace([string]$cellRecord.comment)) {
                        [void]$cell.AddComment([string]$cellRecord.comment)
                        $reportCounters.restored_comments++
                    }
                    elseif ($null -ne $cellRecord.comment_threaded) {
                        $plainText = Convert-ThreadedCommentToPlainText -ThreadedComment $cellRecord.comment_threaded
                        if (-not [string]::IsNullOrWhiteSpace($plainText)) {
                            [void]$cell.AddComment($plainText)
                            $reportCounters.restored_comments++
                            $reportCounters.threaded_comment_fallbacks++
                            Add-WarningMessage -Warnings $warnings -Message ("Threaded comment restored as legacy comment for {0}." -f $cellLabel)
                        }
                    }
                }
                finally {
                    if ($null -ne $cell) {
                        Release-ComReference $cell
                    }
                }
            }

            Write-Host ('  行高・列幅・非表示を復元中 -> {0}' -f $sheetName)
            foreach ($rowInfo in @($sheetRecord.row_heights)) {
                $rowRange = $null
                try {
                    $rowRange = $sheet.Rows.Item([int]$rowInfo.row)
                    $rowRange.RowHeight = [double]$rowInfo.height
                }
                finally {
                    if ($null -ne $rowRange) {
                        Release-ComReference $rowRange
                    }
                }
            }

            foreach ($columnInfo in @($sheetRecord.column_widths)) {
                $columnRange = $null
                try {
                    $columnIndex = Convert-ColumnLettersToNumber -ColumnLetters ([string]$columnInfo.column)
                    $columnRange = $sheet.Columns.Item($columnIndex)
                    $columnRange.ColumnWidth = [double]$columnInfo.width
                }
                finally {
                    if ($null -ne $columnRange) {
                        Release-ComReference $columnRange
                    }
                }
            }

            foreach ($hiddenRow in @($sheetRecord.hidden_rows)) {
                $rowRange = $null
                try {
                    $rowRange = $sheet.Rows.Item([int]$hiddenRow)
                    $rowRange.Hidden = $true
                }
                finally {
                    if ($null -ne $rowRange) {
                        Release-ComReference $rowRange
                    }
                }
            }

            foreach ($hiddenColumn in @($sheetRecord.hidden_columns)) {
                $columnRange = $null
                try {
                    $columnIndex = Convert-ColumnLettersToNumber -ColumnLetters ([string]$hiddenColumn)
                    $columnRange = $sheet.Columns.Item($columnIndex)
                    $columnRange.Hidden = $true
                }
                finally {
                    if ($null -ne $columnRange) {
                        Release-ComReference $columnRange
                    }
                }
            }

            Write-Host ('  結合セルを復元中 -> {0}' -f $sheetName)
            foreach ($mergeRecord in @($sheetRecord.merged_ranges)) {
                if (-not (Test-ValidMergeAddress -Address ([string]$mergeRecord.range))) {
                    Add-WarningMessage -Warnings $warnings -Message ("Invalid merge range skipped for {0}: {1}" -f $sheetName, [string]$mergeRecord.range)
                    continue
                }

                $mergeRange = $null
                try {
                    $mergeRange = $sheet.Range([string]$mergeRecord.range)
                    $mergeRange.Merge()
                    $reportCounters.restored_merged_ranges++
                }
                catch {
                    Add-WarningMessage -Warnings $warnings -Message ("Merge restore failed for {0}!{1}: {2}" -f $sheetName, [string]$mergeRecord.range, $_.Exception.Message)
                }
                finally {
                    if ($null -ne $mergeRange) {
                        Release-ComReference $mergeRange
                    }
                }
            }

            Write-Host ('  書式を復元中 -> {0}' -f $sheetName)
            if ($styleLookup.ContainsKey($sheetName)) {
                foreach ($styleRecord in $styleLookup[$sheetName].Values) {
                    $cellRecord = $cellLookup[$sheetName][[string]$styleRecord.address]
                    if ($null -ne $cellRecord.merge_area -and -not [bool]$cellRecord.is_merge_anchor) {
                        continue
                    }

                    $styleCell = $null
                    try {
                        $styleCell = $sheet.Range([string]$styleRecord.address)
                        Apply-CellStyleRecord -Cell $styleCell -StyleRecord $styleRecord
                        $reportCounters.restored_styles++
                    }
                    catch {
                        Add-WarningMessage -Warnings $warnings -Message ("Style restore failed for {0}!{1}: {2}" -f $sheetName, [string]$styleRecord.address, $_.Exception.Message)
                    }
                    finally {
                        if ($null -ne $styleCell) {
                            Release-ComReference $styleCell
                        }
                    }
                }
            }

            Write-Host ('  ウィンドウ枠固定を復元中 -> {0}' -f $sheetName)
            if ($null -ne $sheetRecord.freeze_panes) {
                try {
                    Set-WorksheetFreezeState -Excel $excel -Worksheet $sheet -FreezeState $sheetRecord.freeze_panes
                }
                catch {
                    Add-WarningMessage -Warnings $warnings -Message ("Freeze panes restore failed for {0}: {1}" -f $sheetName, $_.Exception.Message)
                }
            }
            Write-Host ('シート復元完了 -> {0}' -f $sheetName)
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }
    }

    $visibleSheets = @($sheetRecords | Where-Object { [int]$_.visible -eq -1 })
    if ($visibleSheets.Count -eq 0) {
        Add-WarningMessage -Warnings $warnings -Message 'No visible sheet was defined in workbook.json. The first sheet remains visible.'
        $sheetLookup[[string]$sheetRecords[0].sheet_name].visible = -1
    }

    foreach ($sheetRecord in ($sheetRecords | Sort-Object { [int]$_.visible -eq -1 } -Descending)) {
        $sheet = $null
        try {
            $sheet = $workbook.Worksheets.Item([int]$sheetRecord.sheet_index)
            $sheet.Visible = [int]$sheetRecord.visible
        }
        finally {
            if ($null -ne $sheet) {
                Release-ComReference $sheet
            }
        }
    }

    if (Test-Path -LiteralPath $resolvedOutputPath) {
        Remove-Item -LiteralPath $resolvedOutputPath -Force
    }

    Write-Host ('復元した Excel を保存中 -> {0}' -f $resolvedOutputPath)
    $workbook.SaveAs($resolvedOutputPath, 51)
    $excel.CalculateFullRebuild()

    $reportPayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Rebuilder'
        status = if ($warnings.Count -gt 0) { 'warning' } else { 'success' }
        warnings = $warnings
        workbook_json_path = $resolvedWorkbookJsonPath
        styles_json_path = $resolvedStylesJsonPath
        output_path = $resolvedOutputPath
        output_extension = '.xlsx'
        source_has_vba = [bool]$workbookData.workbook.has_vba
        restored_sheets = $reportCounters.restored_sheets
        restored_cells = $reportCounters.restored_cells
        restored_formulas = $reportCounters.restored_formulas
        restored_comments = $reportCounters.restored_comments
        restored_hyperlinks = $reportCounters.restored_hyperlinks
        restored_styles = $reportCounters.restored_styles
        restored_merged_ranges = $reportCounters.restored_merged_ranges
        threaded_comment_fallbacks = $reportCounters.threaded_comment_fallbacks
    }

    Write-JsonFile -Data $reportPayload -Path $rebuildReportPath

    Write-NextStepBlock -Steps @(
        ('復元した Excel を開いて確認する: {0}' -f $resolvedOutputPath),
        ('必要なら tools\advanced\run_extract.bat "{0}"' -f $resolvedOutputPath)
    )
    Write-Host ('復元した Excel     -> {0}' -f $resolvedOutputPath)
    Write-Host ('rebuild_report.json -> {0}' -f $rebuildReportPath)
}
catch {
    Write-ErrorRecoverySteps -CommandName 'rebuild'
    throw "rebuild_excel.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($true)
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
