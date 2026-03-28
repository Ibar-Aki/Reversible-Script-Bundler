Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

function Get-ProjectRoot {
    return (Split-Path -Path $PSScriptRoot -Parent)
}

function Get-OutputRootDirectory {
    $outputRoot = Join-Path (Get-ProjectRoot) 'output'
    Ensure-Directory -Path $outputRoot
    return $outputRoot
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path)
}

function Get-NormalizedFullPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

function Convert-ToSafeDirectoryName {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $invalidCharacters = [System.IO.Path]::GetInvalidFileNameChars()
    $sanitized = [string]$Name
    foreach ($character in $invalidCharacters) {
        $sanitized = $sanitized.Replace([string]$character, '_')
    }

    $sanitized = $sanitized -replace '\s+', '_'
    $sanitized = $sanitized.Trim(' ', '.')
    if ([string]::IsNullOrWhiteSpace($sanitized)) {
        return 'workbook'
    }

    return $sanitized
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory)]
        [object]$Data,
        [Parameter(Mandatory)]
        [string]$Path,
        [int]$Depth = 100
    )

    $json = ($Data | ConvertTo-Json -Depth $Depth) -replace "`r`n", "`n"
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Write-JsonLineFile {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Items,
        [Parameter(Mandatory)]
        [string]$Path,
        [int]$Depth = 50
    )

    $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($item in $Items) {
            $line = $item | ConvertTo-Json -Depth $Depth -Compress
            $writer.WriteLine($line)
        }
    }
    finally {
        $writer.Dispose()
    }
}

function Get-RangeMatrixValue {
    param(
        $Matrix,
        [Parameter(Mandatory)]
        [int]$RowOffset,
        [Parameter(Mandatory)]
        [int]$ColumnOffset,
        [Parameter(Mandatory)]
        [int]$RowCount,
        [Parameter(Mandatory)]
        [int]$ColumnCount
    )

    if ($null -eq $Matrix -or $Matrix -is [System.DBNull]) {
        return $null
    }

    if ($Matrix -isnot [System.Array]) {
        if ($RowOffset -eq 1 -and $ColumnOffset -eq 1) {
            return $Matrix
        }

        return $null
    }

    if ($Matrix.Rank -eq 2) {
        $rowIndex = $Matrix.GetLowerBound(0) + $RowOffset - 1
        $columnIndex = $Matrix.GetLowerBound(1) + $ColumnOffset - 1
        return $Matrix.GetValue($rowIndex, $columnIndex)
    }

    if ($Matrix.Rank -eq 1) {
        $index = $Matrix.GetLowerBound(0)
        if ($RowCount -eq 1) {
            $index += $ColumnOffset - 1
        }
        elseif ($ColumnCount -eq 1) {
            $index += $RowOffset - 1
        }
        elseif ($RowOffset -ne 1 -or $ColumnOffset -ne 1) {
            return $null
        }

        return $Matrix.GetValue($index)
    }

    return $null
}

function Convert-FormulaValue {
    param(
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    $formulaText = [string]$Value
    if ([string]::IsNullOrWhiteSpace($formulaText)) {
        return $null
    }

    if ($formulaText.StartsWith('=')) {
        return $formulaText
    }

    return $null
}

function Get-TimestampJst {
    return (Get-Date).ToString("yyyy-MM-dd HH:mm 'JST'")
}

function Get-TimestampJstForPath {
    return (Get-Date).ToString('yyyyMMdd-HHmmss')
}

function Get-DefaultRunOutputDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$ExcelPath
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath)
    $safeBaseName = Convert-ToSafeDirectoryName -Name $baseName
    return Join-Path (Get-OutputRootDirectory) ('{0}_{1}' -f $safeBaseName, (Get-TimestampJstForPath))
}

function Get-LatestOutputDirectoryPointerPath {
    return Join-Path (Get-OutputRootDirectory) 'latest_run.txt'
}

function Set-LatestOutputDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$OutputDir
    )

    $resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
    [System.IO.File]::WriteAllText((Get-LatestOutputDirectoryPointerPath), $resolvedOutputDir, [System.Text.Encoding]::UTF8)
}

function Get-LatestOutputDirectory {
    $pointerPath = Get-LatestOutputDirectoryPointerPath
    if (-not (Test-Path -LiteralPath $pointerPath)) {
        throw '最新の実行結果が見つかりません。先に Excel2LLM.bat または tools\advanced\run_extract.bat を実行してください。'
    }

    $resolvedOutputDir = [string]([System.IO.File]::ReadAllText($pointerPath, [System.Text.Encoding]::UTF8)).Trim()
    if ([string]::IsNullOrWhiteSpace($resolvedOutputDir)) {
        throw 'latest_run.txt の内容が空です。先に Excel2LLM.bat または tools\advanced\run_extract.bat を実行してください。'
    }

    if (-not (Test-Path -LiteralPath $resolvedOutputDir)) {
        throw ("最新の出力フォルダが存在しません: {0}" -f $resolvedOutputDir)
    }

    return (Get-NormalizedFullPath -Path $resolvedOutputDir)
}

function Write-NextStepBlock {
    param(
        [Parameter(Mandatory)]
        [string[]]$Steps,
        [string]$Title = '次のおすすめ'
    )

    $normalizedSteps = @($Steps | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($normalizedSteps.Count -eq 0) {
        return
    }

    Write-Host ('=== {0} ===' -f $Title)
    foreach ($step in $normalizedSteps) {
        Write-Host ('  - {0}' -f [string]$step)
    }
}

function Write-ErrorRecoverySteps {
    param(
        [string]$CommandName = 'このコマンド'
    )

    Write-Host ('{0} の実行中にエラーが発生しました。' -f $CommandName)
    Write-Host '対処の目安:'
    Write-Host '  1. Excel を閉じる'
    Write-Host '  2. コマンドをもう一度実行する'
    Write-Host '  3. まだダメなら Excel2LLM.bat -SelfTest を実行する'
    Write-Host '  ※ このあとに表示される英語メッセージは、技術調査用の詳細情報です。'
}

function Group-CellsBySheet {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Cells
    )

    $lookup = @{}
    foreach ($cell in $Cells) {
        $sheetName = [string]$cell.sheet
        if (-not $lookup.ContainsKey($sheetName)) {
            $lookup[$sheetName] = [System.Collections.Generic.List[object]]::new()
        }

        [void]$lookup[$sheetName].Add($cell)
    }

    return $lookup
}

function Convert-ExcelColor {
    param(
        $ColorValue
    )

    if ($null -eq $ColorValue) {
        return $null
    }

    try {
        $number = [int64]$ColorValue
    }
    catch {
        return $null
    }

    if ($number -lt 0) {
        return $null
    }

    $red = $number -band 0xFF
    $green = ($number -shr 8) -band 0xFF
    $blue = ($number -shr 16) -band 0xFF
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
}

function Convert-VariantValue {
    param(
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    if ($Value -is [DateTime]) {
        return $Value.ToString('o')
    }

    if ($Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return $Value
    }

    return [string]$Value
}

function Add-WarningMessage {
    param(
        [Parameter(Mandatory)]
        $Warnings,
        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $Warnings.Add($Message)
    }
}

function Release-ComReference {
    param(
        $Reference
    )

    if ($null -ne $Reference -and [System.Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Reference)
    }
}

function New-ExcelApplication {
    param(
        [switch]$AllowWorkbookMacros
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    if (-not $AllowWorkbookMacros) {
        try {
            $excel.AutomationSecurity = 3
        }
        catch {
        }
    }
    return $excel
}

function Test-PathWithinDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$DirectoryPath
    )

    $resolvedPath = Get-NormalizedFullPath -Path $Path
    $resolvedDirectoryPath = Get-NormalizedFullPath -Path $DirectoryPath
    $directoryPrefix = if ($resolvedDirectoryPath.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
        $resolvedDirectoryPath
    }
    else {
        $resolvedDirectoryPath + [System.IO.Path]::DirectorySeparatorChar
    }

    return $resolvedPath.Equals($resolvedDirectoryPath, [System.StringComparison]::OrdinalIgnoreCase) -or
        $resolvedPath.StartsWith($directoryPrefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-BorderNames {
    return [ordered]@{
        left = 7
        top = 8
        bottom = 9
        right = 10
        inside_vertical = 11
        inside_horizontal = 12
        diagonal_down = 5
        diagonal_up = 6
    }
}

function Get-CellHyperlink {
    param(
        $Cell
    )

    $link = $null
    try {
        if ($Cell.Hyperlinks.Count -gt 0) {
            $link = $Cell.Hyperlinks.Item(1)
            return [ordered]@{
                address = if ([string]::IsNullOrWhiteSpace([string]$link.Address)) { $null } else { [string]$link.Address }
                sub_address = if ([string]::IsNullOrWhiteSpace([string]$link.SubAddress)) { $null } else { [string]$link.SubAddress }
                text_to_display = if ([string]::IsNullOrWhiteSpace([string]$link.TextToDisplay)) { $null } else { [string]$link.TextToDisplay }
            }
        }
    }
    catch {
        return $null
    }
    finally {
        if ($null -ne $link) {
            Release-ComReference $link
        }
    }

    return $null
}

function Get-CellCommentText {
    param(
        $Cell
    )

    try {
        if ($null -ne $Cell.Comment) {
            return [string]$Cell.Comment.Text()
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-CellThreadedComment {
    param(
        $Cell
    )

    $commentThreaded = $null
    $repliesCollection = $null

    try {
        try {
            $commentThreaded = $Cell.CommentThreaded
        }
        catch {
            return $null
        }

        if ($null -eq $commentThreaded) {
            return $null
        }

        $text = $null
        try {
            $text = [string]$commentThreaded.Text()
        }
        catch {
            try {
                $text = [string]$commentThreaded.Text
            }
            catch {
                $text = $null
            }
        }

        $author = $null
        try {
            $author = [string]$commentThreaded.Author.Name
        }
        catch {
            try {
                $author = [string]$commentThreaded.Author
            }
            catch {
                $author = $null
            }
        }

        $createdAt = $null
        try {
            $createdAt = ([datetime]$commentThreaded.Date).ToString('o')
        }
        catch {
            $createdAt = $null
        }

        $replyList = [System.Collections.Generic.List[object]]::new()
        try {
            $repliesCollection = $commentThreaded.Replies
            if ($null -ne $repliesCollection) {
                foreach ($reply in $repliesCollection) {
                    try {
                        $replyText = $null
                        $replyAuthor = $null
                        $replyDate = $null

                        try {
                            $replyText = [string]$reply.Text()
                        }
                        catch {
                            try {
                                $replyText = [string]$reply.Text
                            }
                            catch {
                                $replyText = $null
                            }
                        }

                        try {
                            $replyAuthor = [string]$reply.Author.Name
                        }
                        catch {
                            try {
                                $replyAuthor = [string]$reply.Author
                            }
                            catch {
                                $replyAuthor = $null
                            }
                        }

                        try {
                            $replyDate = ([datetime]$reply.Date).ToString('o')
                        }
                        catch {
                            $replyDate = $null
                        }

                        [void]$replyList.Add([ordered]@{
                            text = $replyText
                            author = $replyAuthor
                            created_at = $replyDate
                        })
                    }
                    finally {
                        Release-ComReference $reply
                    }
                }
            }
        }
        catch {
        }

        return [ordered]@{
            text = $text
            author = $author
            created_at = $createdAt
            replies = $replyList
        }
    }
    finally {
        if ($null -ne $repliesCollection) {
            Release-ComReference $repliesCollection
        }
        if ($null -ne $commentThreaded) {
            Release-ComReference $commentThreaded
        }
    }
}

function Get-CellFormula2 {
    param(
        $Cell
    )

    try {
        $formula2 = $Cell.Formula2
        if ($null -eq $formula2 -or [string]$formula2 -eq '') {
            $formula2 = $null
        }
    }
    catch {
        $formula2 = $null
    }

    if ($null -eq $formula2) {
        try {
            $fallbackFormula = $Cell.Formula
            if ($null -ne $fallbackFormula -and [string]$fallbackFormula -ne '') {
                return [string]$fallbackFormula
            }
        }
        catch {
        }

        return $null
    }

    return [string]$formula2
}

function Get-WorksheetFreezeState {
    param(
        [Parameter(Mandatory)]
        $Excel,
        [Parameter(Mandatory)]
        $Worksheet
    )

    $state = [ordered]@{
        enabled = $false
        split_row = 0
        split_column = 0
    }

    try {
        [void]$Worksheet.Activate()
        $window = $Excel.ActiveWindow
        if ($null -ne $window) {
            $state.enabled = [bool]$window.FreezePanes
            $state.split_row = [int]$window.SplitRow
            $state.split_column = [int]$window.SplitColumn
        }
    }
    catch {
        $state.enabled = $false
    }
    finally {
        if ($null -ne $window) {
            Release-ComReference $window
        }
    }

    return $state
}

function Get-UsedRangeInfo {
    param(
        [Parameter(Mandatory)]
        $UsedRange
    )

    $firstRow = [int]$UsedRange.Row
    $firstColumn = [int]$UsedRange.Column
    $rowCount = [int]$UsedRange.Rows.Count
    $columnCount = [int]$UsedRange.Columns.Count

    return [ordered]@{
        address = [string]$UsedRange.Address($false, $false)
        first_row = $firstRow
        first_column = $firstColumn
        last_row = $firstRow + $rowCount - 1
        last_column = $firstColumn + $columnCount - 1
        row_count = $rowCount
        column_count = $columnCount
    }
}

function Get-StyleRecord {
    param(
        [Parameter(Mandatory)]
        $Cell
    )

    $record = [ordered]@{
        fill_color = $null
        font_color = $null
        horizontal_alignment = $null
        vertical_alignment = $null
        wrap_text = $null
        borders = [ordered]@{}
    }

    try {
        $record.fill_color = Convert-ExcelColor $Cell.Interior.Color
    }
    catch {
        $record.fill_color = $null
    }

    try {
        $record.font_color = Convert-ExcelColor $Cell.Font.Color
    }
    catch {
        $record.font_color = $null
    }

    try {
        $record.horizontal_alignment = [int]$Cell.HorizontalAlignment
    }
    catch {
        $record.horizontal_alignment = $null
    }

    try {
        $record.vertical_alignment = [int]$Cell.VerticalAlignment
    }
    catch {
        $record.vertical_alignment = $null
    }

    try {
        $record.wrap_text = [bool]$Cell.WrapText
    }
    catch {
        $record.wrap_text = $null
    }

    $borders = $null
    try {
        $borders = $Cell.Borders
        $overallLineStyle = $null
        try {
            $overallLineStyle = [int]$borders.LineStyle
        }
        catch {
            $overallLineStyle = $null
        }

        if ($null -eq $overallLineStyle -or $overallLineStyle -eq -4142) {
            foreach ($pair in (Get-BorderNames).GetEnumerator()) {
                $record.borders[$pair.Key] = $null
            }
        }
        else {
            foreach ($pair in (Get-BorderNames).GetEnumerator()) {
                $border = $null
                try {
                    $border = $borders.Item($pair.Value)
                    $lineStyle = $null
                    $weight = $null
                    $color = $null

                    try {
                        $lineStyle = [int]$border.LineStyle
                    }
                    catch {
                        $lineStyle = $null
                    }

                    try {
                        $weight = [int]$border.Weight
                    }
                    catch {
                        $weight = $null
                    }

                    try {
                        $color = Convert-ExcelColor $border.Color
                    }
                    catch {
                        $color = $null
                    }

                    $record.borders[$pair.Key] = [ordered]@{
                        line_style = $lineStyle
                        weight = $weight
                        color = $color
                    }
                }
                catch {
                    $record.borders[$pair.Key] = $null
                }
                finally {
                    if ($null -ne $border) {
                        Release-ComReference $border
                    }
                }
            }
        }
    }
    finally {
        if ($null -ne $borders) {
            Release-ComReference $borders
        }
    }

    return $record
}

function Convert-CoordinateToA1 {
    param(
        [Parameter(Mandatory)]
        [int]$Row,
        [Parameter(Mandatory)]
        [int]$Column
    )

    $col = $Column
    $letters = ''
    while ($col -gt 0) {
        $remainder = [int](($col - 1) % 26)
        $letters = [char][int](65 + $remainder) + $letters
        $col = [int][math]::Floor(($col - 1) / 26)
    }

    return '{0}{1}' -f $letters, $Row
}

function Convert-ColumnLettersToNumber {
    param(
        [Parameter(Mandatory)]
        [string]$ColumnLetters
    )

    $normalized = $ColumnLetters.Trim().ToUpperInvariant()
    if ($normalized -notmatch '^[A-Z]+$') {
        throw "Invalid column letters: $ColumnLetters"
    }

    $columnNumber = 0
    foreach ($letter in $normalized.ToCharArray()) {
        $columnNumber = ($columnNumber * 26) + ([int][char]$letter - [int][char]'A' + 1)
    }

    return $columnNumber
}

function Convert-A1ToCoordinate {
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    $match = [System.Text.RegularExpressions.Regex]::Match($Address.Trim().ToUpperInvariant(), '^([A-Z]+)(\d+)$')
    if (-not $match.Success) {
        throw "Invalid A1 address: $Address"
    }

    return [ordered]@{
        row = [int]$match.Groups[2].Value
        column = Convert-ColumnLettersToNumber -ColumnLetters $match.Groups[1].Value
    }
}

function Get-SheetLookupByName {
    param(
        [Parameter(Mandatory)]
        [object[]]$Sheets
    )

    $lookup = @{}
    foreach ($sheet in $Sheets) {
        $sheetName = [string]$sheet.sheet_name
        if ($lookup.ContainsKey($sheetName)) {
            throw "Duplicate sheet name detected: $sheetName"
        }

        $lookup[$sheetName] = $sheet
    }

    return $lookup
}

function Get-CellLookupBySheetAndAddress {
    param(
        [Parameter(Mandatory)]
        [object[]]$Cells
    )

    $lookup = @{}
    foreach ($cell in $Cells) {
        $sheetName = [string]$cell.sheet
        $address = [string]$cell.address

        if (-not $lookup.ContainsKey($sheetName)) {
            $lookup[$sheetName] = @{}
        }

        if ($lookup[$sheetName].ContainsKey($address)) {
            throw "Duplicate cell address detected: $sheetName!$address"
        }

        $lookup[$sheetName][$address] = $cell
    }

    return $lookup
}

function Get-StyleLookupBySheetAndAddress {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Styles
    )

    $lookup = @{}
    foreach ($style in $Styles) {
        $sheetName = [string]$style.sheet
        $address = [string]$style.address

        if (-not $lookup.ContainsKey($sheetName)) {
            $lookup[$sheetName] = @{}
        }

        $lookup[$sheetName][$address] = $style
    }

    return $lookup
}

function Convert-HexColorToExcelColor {
    param(
        [string]$Color
    )

    if ([string]::IsNullOrWhiteSpace($Color)) {
        return $null
    }

    $normalized = $Color.Trim()
    if ($normalized -notmatch '^#?([0-9A-Fa-f]{6})$') {
        throw "Invalid color code: $Color"
    }

    $hex = $Matches[1]
    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $red + ($green -shl 8) + ($blue -shl 16)
}

function Set-WorksheetFreezeState {
    param(
        [Parameter(Mandatory)]
        $Excel,
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        $FreezeState
    )

    $window = $null
    try {
        [void]$Worksheet.Activate()
        $window = $Excel.ActiveWindow
        if ($null -eq $window) {
            return
        }

        $window.FreezePanes = $false
        $window.SplitRow = [int]$FreezeState.split_row
        $window.SplitColumn = [int]$FreezeState.split_column
        $window.FreezePanes = [bool]$FreezeState.enabled
    }
    finally {
        if ($null -ne $window) {
            Release-ComReference $window
        }
    }
}
