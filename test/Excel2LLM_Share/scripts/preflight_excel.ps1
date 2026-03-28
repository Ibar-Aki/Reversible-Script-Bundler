[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,
    [string]$OutputDir,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

$ErrorActionPreference = 'Stop'

if (-not $OutputDir) {
    $OutputDir = Get-DefaultRunOutputDirectory -ExcelPath $ExcelPath
}

$warningCellThreshold = 1000000L
$blockingCellThreshold = 5000000L
$blockingSheetCellThreshold = 2000000L
$warningFileSizeBytes = 50MB
$blockingFileSizeBytes = 200MB

function Get-DisplayPathValue {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$LeafOnly
    )

    if ($LeafOnly) {
        return [System.IO.Path]::GetFileName($Path)
    }

    return $Path
}

function Get-DisplayDirectoryValue {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$LeafOnly
    )

    if ($LeafOnly) {
        return [string](Split-Path -Path $Path -Leaf)
    }

    return $Path
}

function Get-FileSizeLabel {
    param(
        [long]$Bytes
    )

    if ($Bytes -lt 1KB) {
        return ('{0} B' -f $Bytes)
    }

    return ('{0:N2} MB' -f ($Bytes / 1MB))
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        $Archive,
        [Parameter(Mandatory)]
        [string]$EntryPath
    )

    $entry = $Archive.GetEntry($EntryPath)
    if ($null -eq $entry) {
        return $null
    }

    $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Resolve-WorkbookTargetPath {
    param(
        [Parameter(Mandatory)]
        [string]$Target
    )

    $normalized = $Target.Replace('\', '/')
    if ($normalized.StartsWith('/')) {
        return $normalized.TrimStart('/')
    }

    if ($normalized.StartsWith('xl/')) {
        return $normalized
    }

    return ('xl/{0}' -f $normalized.TrimStart('./'))
}

function Get-DimensionInfo {
    param(
        [string]$Reference
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return [ordered]@{
            has_dimension = $false
            reference = $null
            estimated_row_count = $null
            estimated_column_count = $null
            estimated_cell_count = $null
        }
    }

    $normalizedReference = $Reference.Trim().ToUpperInvariant().Replace('$', '')
    $rangeParts = $normalizedReference -split ':'
    if ($rangeParts.Count -gt 2) {
        throw "Invalid worksheet dimension reference: $Reference"
    }

    $startCoordinate = Convert-A1ToCoordinate -Address $rangeParts[0]
    $endCoordinate = if ($rangeParts.Count -eq 2) {
        Convert-A1ToCoordinate -Address $rangeParts[1]
    }
    else {
        $startCoordinate
    }

    if ([int]$endCoordinate.row -lt [int]$startCoordinate.row -or [int]$endCoordinate.column -lt [int]$startCoordinate.column) {
        throw "Invalid worksheet dimension ordering: $Reference"
    }

    $rowCount = ([int64]$endCoordinate.row - [int64]$startCoordinate.row) + 1
    $columnCount = ([int64]$endCoordinate.column - [int64]$startCoordinate.column) + 1

    return [ordered]@{
        has_dimension = $true
        reference = $normalizedReference
        estimated_row_count = $rowCount
        estimated_column_count = $columnCount
        estimated_cell_count = ($rowCount * $columnCount)
    }
}

function Write-PreflightSummary {
    param(
        [Parameter(Mandatory)]
        [object]$Report
    )

    $largestSheetLabel = if ($null -eq $Report.largest_sheet -or [string]::IsNullOrWhiteSpace([string]$Report.largest_sheet.name)) {
        'なし'
    }
    else {
        '{0} ({1} セル)' -f $Report.largest_sheet.name, $Report.largest_sheet.estimated_cell_count
    }

    $statusLabel = switch ([string]$Report.status) {
        'success' { '成功' }
        'warning' { '警告あり' }
        'blocked' { '停止' }
        default { [string]$Report.status }
    }

    Write-Host '=== 事前チェック結果 ==='
    Write-Host ('  対象ファイル:     {0}' -f [System.IO.Path]::GetFileName([string]$Report.workbook_path))
    Write-Host ('  判定:             {0}' -f $statusLabel)
    Write-Host ('  ファイルサイズ:   {0}' -f (Get-FileSizeLabel -Bytes ([int64]$Report.file_size_bytes)))
    Write-Host ('  シート数:         {0}' -f [int]$Report.sheet_count)
    Write-Host ('  推定総セル数:     {0}' -f $(if ($null -eq $Report.estimated_total_cells) { '不明' } else { [string]$Report.estimated_total_cells }))
    Write-Host ('  最大シート:       {0}' -f $largestSheetLabel)
    Write-Host ('  レポート:         {0}' -f [string]$Report.report_path)

    foreach ($warning in @($Report.warnings)) {
        Write-Host ('  警告:             {0}' -f [string]$warning)
    }

    foreach ($reason in @($Report.reasons)) {
        Write-Host ('  停止理由:         {0}' -f [string]$reason)
    }
}

function Write-PreflightStopGuidance {
    Write-Host '事前チェックで処理を中止しました。'
    Write-Host '対処:'
    Write-Host '  1. 対象シートを絞る'
    Write-Host '  2. 不要列を削る'
    Write-Host '  3. 集計用シートだけを別ブック化する'
}

if (-not $OutputDir) {
    $OutputDir = Join-Path (Get-ProjectRoot) 'output'
}

$resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
Ensure-Directory -Path $resolvedOutputDir

$reportPath = Join-Path $resolvedOutputDir 'preflight_report.json'
$displayWorkbookPath = if ($RedactPaths) { [System.IO.Path]::GetFileName($ExcelPath) } else { $ExcelPath }
$displayOutputDir = if ($RedactPaths) { [string](Split-Path -Path $resolvedOutputDir -Leaf) } else { $resolvedOutputDir }

$warnings = [System.Collections.Generic.List[string]]::new()
$reasons = [System.Collections.Generic.List[string]]::new()
$sheetReports = [System.Collections.Generic.List[object]]::new()
$sheetCount = 0
$estimatedTotalCells = 0L
$largestSheet = $null

try {
    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        [void]$reasons.Add(("ファイルが見つかりません: {0}" -f $ExcelPath))
    }
    else {
        $resolvedExcelPath = Resolve-AbsolutePath -Path $ExcelPath
        $displayWorkbookPath = Get-DisplayPathValue -Path $resolvedExcelPath -LeafOnly:$RedactPaths
        $fileItem = Get-Item -LiteralPath $resolvedExcelPath
        $fileSizeBytes = [int64]$fileItem.Length
        $extension = [System.IO.Path]::GetExtension($resolvedExcelPath).ToLowerInvariant()

        if (@('.xlsx', '.xlsm') -notcontains $extension) {
            [void]$reasons.Add(("対応していない拡張子です: {0}。対応形式は .xlsx と .xlsm だけです。" -f $extension))
        }

        if ($fileSizeBytes -gt $blockingFileSizeBytes) {
            [void]$reasons.Add(("ファイルサイズが上限を超えています: {0} > {1}" -f (Get-FileSizeLabel -Bytes $fileSizeBytes), (Get-FileSizeLabel -Bytes $blockingFileSizeBytes)))
        }
        elseif ($fileSizeBytes -gt $warningFileSizeBytes) {
            [void]$warnings.Add(("ファイルサイズが大きめです: {0}" -f (Get-FileSizeLabel -Bytes $fileSizeBytes)))
        }

        if ($reasons.Count -eq 0) {
            $zipStream = $null
            $zipArchive = $null
            try {
                $zipStream = [System.IO.File]::OpenRead($resolvedExcelPath)
                $zipArchive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)

                foreach ($requiredEntry in @('[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels')) {
                    if ($null -eq $zipArchive.GetEntry($requiredEntry)) {
                        [void]$reasons.Add(("必須の OpenXML エントリが不足しています: {0}" -f $requiredEntry))
                    }
                }

                if ($reasons.Count -eq 0) {
                    $workbookXmlText = Get-ZipEntryText -Archive $zipArchive -EntryPath 'xl/workbook.xml'
                    $workbookRelsText = Get-ZipEntryText -Archive $zipArchive -EntryPath 'xl/_rels/workbook.xml.rels'

                    if ([string]::IsNullOrWhiteSpace($workbookXmlText) -or [string]::IsNullOrWhiteSpace($workbookRelsText)) {
                        [void]$reasons.Add('workbook.xml の内容を読み取れませんでした。')
                    }
                    else {
                        [xml]$workbookXml = $workbookXmlText
                        [xml]$workbookRelsXml = $workbookRelsText

                        $workbookNs = [System.Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
                        $workbookNs.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                        $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

                        $relsNs = [System.Xml.XmlNamespaceManager]::new($workbookRelsXml.NameTable)
                        $relsNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')

                        $relationshipLookup = @{}
                        foreach ($relationshipNode in @($workbookRelsXml.SelectNodes('/p:Relationships/p:Relationship', $relsNs))) {
                            $relationshipLookup[[string]$relationshipNode.Id] = Resolve-WorkbookTargetPath -Target ([string]$relationshipNode.Target)
                        }

                        $sheetNodes = @($workbookXml.SelectNodes('/s:workbook/s:sheets/s:sheet', $workbookNs))
                        $sheetCount = $sheetNodes.Count

                        foreach ($sheetNode in $sheetNodes) {
                            $sheetName = [string]$sheetNode.GetAttribute('name')
                            $relationshipId = [string]$sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                            $sheetEntryPath = if ($relationshipLookup.ContainsKey($relationshipId)) { [string]$relationshipLookup[$relationshipId] } else { $null }

                            $sheetReport = [ordered]@{
                                name = $sheetName
                                relationship_id = $relationshipId
                                entry_path = $sheetEntryPath
                                dimension = $null
                                dimension_missing = $false
                                estimated_row_count = $null
                                estimated_column_count = $null
                                estimated_cell_count = $null
                                status = 'ok'
                            }

                            if ([string]::IsNullOrWhiteSpace($sheetEntryPath)) {
                                $sheetReport.status = 'blocked'
                                [void]$reasons.Add(("シートの関連付けが見つかりません: {0}" -f $sheetName))
                                [void]$sheetReports.Add($sheetReport)
                                continue
                            }

                            $sheetEntry = $zipArchive.GetEntry($sheetEntryPath)
                            if ($null -eq $sheetEntry) {
                                $sheetReport.status = 'blocked'
                                [void]$reasons.Add(("シート XML が見つかりません: {0} ({1})" -f $sheetName, $sheetEntryPath))
                                [void]$sheetReports.Add($sheetReport)
                                continue
                            }

                            $sheetXmlText = Get-ZipEntryText -Archive $zipArchive -EntryPath $sheetEntryPath
                            if ([string]::IsNullOrWhiteSpace($sheetXmlText)) {
                                $sheetReport.status = 'blocked'
                                [void]$reasons.Add(("シート XML を読み取れませんでした: {0}" -f $sheetName))
                                [void]$sheetReports.Add($sheetReport)
                                continue
                            }

                            try {
                                [xml]$sheetXml = $sheetXmlText
                                $sheetNs = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
                                $sheetNs.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

                                $dimensionNode = $sheetXml.SelectSingleNode('/s:worksheet/s:dimension', $sheetNs)
                                $dimensionRef = if ($null -eq $dimensionNode) { $null } else { [string]$dimensionNode.GetAttribute('ref') }
                                $dimensionInfo = Get-DimensionInfo -Reference $dimensionRef

                                $sheetReport.dimension = $dimensionInfo.reference
                                $sheetReport.dimension_missing = -not [bool]$dimensionInfo.has_dimension
                                $sheetReport.estimated_row_count = $dimensionInfo.estimated_row_count
                                $sheetReport.estimated_column_count = $dimensionInfo.estimated_column_count
                                $sheetReport.estimated_cell_count = $dimensionInfo.estimated_cell_count

                                if (-not $dimensionInfo.has_dimension) {
                                    $sheetReport.status = 'warning'
                                    if ($fileSizeBytes -ge $warningFileSizeBytes) {
                                        $sheetReport.status = 'blocked'
                                        [void]$reasons.Add(("シート範囲情報が見つかりません: {0}。ファイルサイズが 50MB 以上のため停止しました。" -f $sheetName))
                                    }
                                    else {
                                        [void]$warnings.Add(("シート範囲情報が見つかりません: {0}" -f $sheetName))
                                    }
                                }
                                else {
                                    $estimatedTotalCells += [int64]$dimensionInfo.estimated_cell_count
                                    if ($null -eq $largestSheet -or [int64]$dimensionInfo.estimated_cell_count -gt [int64]$largestSheet.estimated_cell_count) {
                                        $largestSheet = [ordered]@{
                                            name = $sheetName
                                            dimension = $dimensionInfo.reference
                                            estimated_cell_count = [int64]$dimensionInfo.estimated_cell_count
                                        }
                                    }

                                    if ([int64]$dimensionInfo.estimated_cell_count -gt $blockingSheetCellThreshold) {
                                        $sheetReport.status = 'blocked'
                                        [void]$reasons.Add(("単一シートの推定セル数が上限を超えています: {0} ({1})" -f $sheetName, [int64]$dimensionInfo.estimated_cell_count))
                                    }
                                }
                            }
                            catch {
                                $sheetReport.status = 'blocked'
                                [void]$reasons.Add(("シート XML を解析できませんでした: {0} ({1})" -f $sheetName, $_.Exception.Message))
                            }

                            [void]$sheetReports.Add($sheetReport)
                        }

                        if ($estimatedTotalCells -gt $blockingCellThreshold) {
                            [void]$reasons.Add(("推定総セル数が上限を超えています: {0}" -f $estimatedTotalCells))
                        }
                        elseif ($estimatedTotalCells -gt $warningCellThreshold) {
                            [void]$warnings.Add(("推定総セル数が大きめです: {0}" -f $estimatedTotalCells))
                        }
                    }
                }
            }
            catch {
                [void]$reasons.Add(("OpenXML ZIP として開けませんでした: {0}" -f $_.Exception.Message))
            }
            finally {
                if ($null -ne $zipArchive) {
                    $zipArchive.Dispose()
                }
                if ($null -ne $zipStream) {
                    $zipStream.Dispose()
                }
            }
        }
    }
}
catch {
    [void]$reasons.Add(("事前チェックが予期せず失敗しました: {0}" -f $_.Exception.Message))
}

$fileSizeValue = if (Test-Path -LiteralPath $ExcelPath) { [int64](Get-Item -LiteralPath $ExcelPath).Length } else { 0L }
$status = if ($reasons.Count -gt 0) { 'blocked' } elseif ($warnings.Count -gt 0) { 'warning' } else { 'success' }
$blocked = ($status -eq 'blocked')

$report = [ordered]@{
    generated_at = Get-TimestampJst
    generator = 'Excel2LLM Preflight Checker'
    status = $status
    blocked = $blocked
    workbook_path = $displayWorkbookPath
    output_directory = $displayOutputDir
    file_size_bytes = $fileSizeValue
    sheet_count = $sheetCount
    estimated_total_cells = $estimatedTotalCells
    largest_sheet = if ($null -eq $largestSheet) { $null } else { $largestSheet }
    reasons = @($reasons)
    warnings = @($warnings)
    sheets = @($sheetReports)
    report_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($reportPath) } else { $reportPath }
}

Write-JsonFile -Data $report -Path $reportPath
Write-PreflightSummary -Report $report

if ($blocked) {
    Write-PreflightStopGuidance
    throw "事前チェックで抽出を停止しました。詳細: $reportPath"
}
