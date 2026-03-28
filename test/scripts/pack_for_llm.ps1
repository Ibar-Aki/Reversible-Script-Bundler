[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookJsonPath,
    [string]$OutputPath,
    [ValidateSet('sheet', 'range')]
    [string]$ChunkBy = 'sheet',
    [int]$MaxCells = 500,
    [switch]$IncludeStyles,
    [string]$StylesJsonPath
)

. (Join-Path $PSScriptRoot 'common.ps1')

function Get-ChunkRange {
    param(
        [Parameter(Mandatory)]
        [object[]]$ChunkCells
    )

    $minRow = [int]$ChunkCells[0].row
    $maxRow = [int]$ChunkCells[0].row
    $minColumn = [int]$ChunkCells[0].column
    $maxColumn = [int]$ChunkCells[0].column

    foreach ($cell in $ChunkCells) {
        $row = [int]$cell.row
        $column = [int]$cell.column

        if ($row -lt $minRow) { $minRow = $row }
        if ($row -gt $maxRow) { $maxRow = $row }
        if ($column -lt $minColumn) { $minColumn = $column }
        if ($column -gt $maxColumn) { $maxColumn = $column }
    }

    $start = Convert-CoordinateToA1 -Row $minRow -Column $minColumn
    $end = Convert-CoordinateToA1 -Row $maxRow -Column $maxColumn
    return '{0}:{1}' -f $start, $end
}

function Get-ChunkPayload {
    param(
        [Parameter(Mandatory)]
        [string]$SheetName,
        [Parameter(Mandatory)]
        [string]$ChunkRange,
        [Parameter(Mandatory)]
        [object[]]$ChunkCells,
        [Parameter(Mandatory)]
        [hashtable]$StyleLookup,
        [switch]$IncludeStyles
    )

    $payloadCells = foreach ($cell in $ChunkCells) {
        $entry = [ordered]@{
            address = $cell.address
            row = [int]$cell.row
            column = [int]$cell.column
            value2 = $cell.value2
            text = $cell.text
            formula = $cell.formula
            formula2 = $cell.formula2
            has_formula = [bool]$cell.has_formula
            number_format = $cell.number_format
            merge_area = $cell.merge_area
            is_merge_anchor = [bool]$cell.is_merge_anchor
            comment = $cell.comment
            comment_threaded = $cell.comment_threaded
            hyperlink = $cell.hyperlink
        }

        if ($IncludeStyles) {
            $styleKey = '{0}|{1}' -f $SheetName, $cell.address
            if ($StyleLookup.ContainsKey($styleKey)) {
                $entry['style'] = $StyleLookup[$styleKey]
            }
        }

        $entry
    }

    return [ordered]@{
        sheet_name = $SheetName
        range = $ChunkRange
        cell_count = $ChunkCells.Count
        cells = $payloadCells
    }
}

function Add-ChunkRecord {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Chunks,
        [Parameter(Mandatory)]
        [string]$SheetName,
        [Parameter(Mandatory)]
        [object[]]$ChunkCells,
        [Parameter(Mandatory)]
        [hashtable]$StyleLookup,
        [switch]$IncludeStyles,
        [Parameter(Mandatory)]
        [ref]$ChunkIndex
    )

    if ($ChunkCells.Count -eq 0) {
        return
    }

    $chunkRange = Get-ChunkRange -ChunkCells $ChunkCells
    $payload = Get-ChunkPayload -SheetName $SheetName -ChunkRange $chunkRange -ChunkCells $ChunkCells -StyleLookup $StyleLookup -IncludeStyles:$IncludeStyles
    $payloadJson = $payload | ConvertTo-Json -Depth 40 -Compress
    $formulaCells = @($ChunkCells | Where-Object { $_.has_formula } | ForEach-Object { $_.address })

    [void]$Chunks.Add([ordered]@{
        chunk_id = ('{0}-{1:D4}' -f $SheetName, $ChunkIndex.Value)
        sheet_name = $SheetName
        range = $chunkRange
        cell_addresses = @($ChunkCells | ForEach-Object { $_.address })
        payload = $payload
        formula_cells = $formulaCells
        token_estimate = [Math]::Ceiling($payloadJson.Length / 4)
        includes_styles = [bool]$IncludeStyles
    })

    $ChunkIndex.Value++
}

try {
    $resolvedWorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath
    $workbookOutputDir = Split-Path -Path $resolvedWorkbookJsonPath -Parent

    if (-not $OutputPath) {
        $OutputPath = Join-Path $workbookOutputDir 'llm_package.jsonl'
    }

    if (-not $StylesJsonPath) {
        $StylesJsonPath = Join-Path $workbookOutputDir 'styles.json'
    }

    $workbookData = Get-Content -LiteralPath $resolvedWorkbookJsonPath -Raw | ConvertFrom-Json
    $styleLookup = @{}
    $cellsBySheet = Group-CellsBySheet -Cells @($workbookData.cells)

    if ($IncludeStyles -and (Test-Path -LiteralPath $StylesJsonPath)) {
        $stylesData = Get-Content -LiteralPath $StylesJsonPath -Raw | ConvertFrom-Json
        foreach ($style in $stylesData.styles) {
            $styleLookup['{0}|{1}' -f $style.sheet, $style.address] = [ordered]@{
                fill_color = $style.fill_color
                font_color = $style.font_color
                horizontal_alignment = $style.horizontal_alignment
                vertical_alignment = $style.vertical_alignment
                wrap_text = $style.wrap_text
                borders = $style.borders
            }
        }
    }

    $chunks = New-Object System.Collections.Generic.List[object]
    $chunkIndex = 0

    foreach ($sheet in $workbookData.sheets) {
        if (-not $cellsBySheet.ContainsKey([string]$sheet.sheet_name)) {
            continue
        }

        $sheetCells = @($cellsBySheet[[string]$sheet.sheet_name] | Sort-Object row, column)
        if ($sheetCells.Count -eq 0) {
            continue
        }

        if ($ChunkBy -eq 'range') {
            for ($offset = 0; $offset -lt $sheetCells.Count; $offset += $MaxCells) {
                $upperBound = [Math]::Min($offset + $MaxCells - 1, $sheetCells.Count - 1)
                $chunkCells = @($sheetCells[$offset..$upperBound])
                Add-ChunkRecord -Chunks $chunks -SheetName $sheet.sheet_name -ChunkCells $chunkCells -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles -ChunkIndex ([ref]$chunkIndex)
            }
            continue
        }

        $rowGroups = @($sheetCells | Group-Object row | Sort-Object { [int]$_.Name })
        $currentChunk = New-Object System.Collections.Generic.List[object]
        foreach ($rowGroup in $rowGroups) {
            $rowCells = @($rowGroup.Group | Sort-Object column)

            if ($currentChunk.Count -gt 0 -and ($currentChunk.Count + $rowCells.Count) -gt $MaxCells) {
                Add-ChunkRecord -Chunks $chunks -SheetName $sheet.sheet_name -ChunkCells ([object[]]$currentChunk.ToArray()) -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles -ChunkIndex ([ref]$chunkIndex)
                $currentChunk = New-Object System.Collections.Generic.List[object]
            }

            if ($rowCells.Count -gt $MaxCells) {
                if ($currentChunk.Count -gt 0) {
                    Add-ChunkRecord -Chunks $chunks -SheetName $sheet.sheet_name -ChunkCells ([object[]]$currentChunk.ToArray()) -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles -ChunkIndex ([ref]$chunkIndex)
                    $currentChunk = New-Object System.Collections.Generic.List[object]
                }

                for ($offset = 0; $offset -lt $rowCells.Count; $offset += $MaxCells) {
                    $upperBound = [Math]::Min($offset + $MaxCells - 1, $rowCells.Count - 1)
                    $rowSlice = @($rowCells[$offset..$upperBound])
                    Add-ChunkRecord -Chunks $chunks -SheetName $sheet.sheet_name -ChunkCells $rowSlice -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles -ChunkIndex ([ref]$chunkIndex)
                }
                continue
            }

            foreach ($cell in $rowCells) {
                [void]$currentChunk.Add($cell)
            }
        }

        if ($currentChunk.Count -gt 0) {
            Add-ChunkRecord -Chunks $chunks -SheetName $sheet.sheet_name -ChunkCells ([object[]]$currentChunk.ToArray()) -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles -ChunkIndex ([ref]$chunkIndex)
        }
    }

    Ensure-Directory -Path (Split-Path -Path $OutputPath -Parent)
    Write-JsonLineFile -Items $chunks -Path $OutputPath
    $chunkArray = @($chunks.ToArray())
    $chunkCount = $chunkArray.Count
    $cellCounts = @($chunkArray | ForEach-Object { $_.cell_addresses.Count })
    $tokenEstimates = @($chunkArray | ForEach-Object { $_.token_estimate })
    $averageCellCount = if ($chunkCount -eq 0) { 0 } else { [Math]::Round((($cellCounts | Measure-Object -Average).Average), 2) }
    $maxTokenEstimate = if ($chunkCount -eq 0) { 0 } else { ($tokenEstimates | Measure-Object -Maximum).Maximum }
    Write-Host '=== パッキング結果 ==='
    Write-Host ('  チャンク数:       {0}' -f $chunkCount)
    Write-Host ('  チャンク方式:     {0}' -f $ChunkBy)
    Write-Host ('  平均セル数:       {0}' -f $averageCellCount)
    Write-Host ('  最大トークン推定: {0}' -f $maxTokenEstimate)
    Write-Host ('  出力:             {0}' -f $OutputPath)
    Write-NextStepBlock -Steps @(
        ('llm_package.jsonl から必要チャンクを LLM に渡す: {0}' -f $OutputPath)
    )
}
catch {
    Write-ErrorRecoverySteps -CommandName 'pack'
    throw "pack_for_llm.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
