[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookJsonPath,
    [Parameter(Mandatory)]
    [string]$JsonlPath,
    [ValidateSet('general', 'mechanical', 'accounting')]
    [string]$Scenario = 'general',
    [string]$OutputDir,
    [int]$MaxChunkPrompts = 3,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

function Get-ScenarioInstructions {
    param(
        [Parameter(Mandatory)]
        [string]$Scenario
    )

    switch ($Scenario) {
        'mechanical' {
            return [ordered]@{
                role = 'あなたは機械設計向けの計算シートレビューアシスタントです。'
                objective = '計算シートの構造、計算手順、数式の役割を読み解き、妥当性と改善余地を整理してください。'
                tasks = @(
                    'formula と formula2 を優先して確認する',
                    '入力値、中間計算、最終結果を分類する',
                    '計算手順を上流から下流まで説明する',
                    '重複計算、分かりにくい参照、冗長なセルを指摘する',
                    '可読性、保守性、説明しやすさの改善案を出す'
                )
            }
        }
        'accounting' {
            return [ordered]@{
                role = 'あなたは会計管理表レビューアシスタントです。'
                objective = '会計管理表の構造、集計ロジック、異常値候補、確認すべき箇所を整理してください。'
                tasks = @(
                    'formula と formula2 を優先して確認する',
                    '売上、費用、利益、予算差異の流れを整理する',
                    '異常値や確認すべき箇所を指摘する',
                    '入力ミスや集計漏れの可能性を挙げる',
                    '改善案を優先順位付きで提案する'
                )
            }
        }
        default {
            return [ordered]@{
                role = 'あなたは Excel 分析支援アシスタントです。'
                objective = 'Excel データの構造、数式、主要な問題点を整理してください。'
                tasks = @(
                    'formula と formula2 を優先して確認する',
                    'value2 と text の違いがあれば補足する',
                    '重要セルと問題点を整理する',
                    '改善案を提案する'
                )
            }
        }
    }
}

function Convert-ChunkToPromptText {
    param(
        [Parameter(Mandatory)]
        $Chunk,
        [Parameter(Mandatory)]
        [hashtable]$ScenarioInstructions,
        [Parameter(Mandatory)]
        [string]$WorkbookName
    )

    $taskLines = $ScenarioInstructions.tasks | ForEach-Object { '- ' + $_ }
    $inputJson = $Chunk | ConvertTo-Json -Depth 50

    return @(
        $ScenarioInstructions.role
        ''
        '目的:'
        ('- ' + $ScenarioInstructions.objective)
        ''
        '対象:'
        ('- 対象ファイル: ' + $WorkbookName)
        ('- 対象シート: ' + [string]$Chunk.sheet_name)
        ('- 対象範囲: ' + [string]$Chunk.range)
        ''
        'データの見方:'
        '- 数式は formula と formula2 を優先して確認する'
        '- 値は value2 を基準に確認する'
        '- 表示差が必要なときだけ text を参照する'
        '- style 情報は補助扱いとする'
        ''
        '作業指示:'
        $taskLines
        ''
        '出力形式:'
        '- 1. 概要'
        '- 2. 発見事項'
        '- 3. 問題点'
        '- 4. 改善案'
        '- 5. 追加確認点'
        ''
        '制約:'
        '- 根拠のない推測はしない'
        '- 不明点は不明と書く'
        '- 数式に基づく説明では対象セルや列の関係を示す'
        ''
        '入力データ:'
        $inputJson
    ) -join [Environment]::NewLine
}

try {
    if (-not $OutputDir) {
        $OutputDir = Join-Path (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'output') 'prompt_bundle'
    }

    Ensure-Directory -Path $OutputDir
    $resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
    $manifestPath = Join-Path $resolvedOutputDir 'prompt_bundle_manifest.json'
    $warnings = [System.Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $manifestPath) {
        try {
            $existingManifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            foreach ($existingPrompt in @($existingManifest.prompts)) {
                $existingPath = [string]$existingPrompt.path
                if ([string]::IsNullOrWhiteSpace($existingPath)) {
                    continue
                }

                $resolvedExistingPath = $null
                try {
                    $pathCandidate = if ([System.IO.Path]::IsPathRooted($existingPath)) {
                        $existingPath
                    }
                    else {
                        Join-Path $resolvedOutputDir $existingPath
                    }
                    $resolvedExistingPath = Get-NormalizedFullPath -Path $pathCandidate
                }
                catch {
                    Add-WarningMessage -Warnings $warnings -Message ("Skipped prompt cleanup for invalid path: {0}" -f $existingPath)
                    continue
                }

                $fileName = [System.IO.Path]::GetFileName($resolvedExistingPath)
                if (-not (Test-PathWithinDirectory -Path $resolvedExistingPath -DirectoryPath $resolvedOutputDir)) {
                    Add-WarningMessage -Warnings $warnings -Message ("Skipped prompt cleanup outside output directory: {0}" -f $resolvedExistingPath)
                    continue
                }

                if ($fileName -notlike 'prompt_*.txt') {
                    Add-WarningMessage -Warnings $warnings -Message ("Skipped prompt cleanup for unmanaged file name: {0}" -f $resolvedExistingPath)
                    continue
                }

                if (Test-Path -LiteralPath $resolvedExistingPath) {
                    Remove-Item -LiteralPath $resolvedExistingPath -Force
                }
            }
        }
        catch {
            Add-WarningMessage -Warnings $warnings -Message ("Prompt bundle cleanup skipped because manifest could not be read: {0}" -f $_.Exception.Message)
        }
    }

    $resolvedWorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath
    $resolvedJsonlPath = Resolve-AbsolutePath -Path $JsonlPath
    $workbookData = Get-Content -LiteralPath $resolvedWorkbookJsonPath -Raw | ConvertFrom-Json
    $chunks = Get-Content -LiteralPath $resolvedJsonlPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_ | ConvertFrom-Json }
    $scenarioInstructions = Get-ScenarioInstructions -Scenario $Scenario

    $selectedChunks = @(
        $chunks |
        Sort-Object @{ Expression = { @($_.formula_cells).Count }; Descending = $true }, @{ Expression = { $_.payload.cell_count }; Descending = $true }, @{ Expression = { [string]$_.sheet_name } } |
        Select-Object -First $MaxChunkPrompts
    )
    $promptFiles = [System.Collections.Generic.List[object]]::new()
    $index = 1

    foreach ($chunk in $selectedChunks) {
        $fileName = 'prompt_{0:D2}_{1}_{2}.txt' -f $index, $Scenario, ([string]$chunk.sheet_name).Replace(' ', '_')
        $promptPath = Join-Path $resolvedOutputDir $fileName
        $content = Convert-ChunkToPromptText -Chunk $chunk -ScenarioInstructions $scenarioInstructions -WorkbookName ([string]$workbookData.workbook.name)
        [System.IO.File]::WriteAllText($promptPath, $content, [System.Text.Encoding]::UTF8)
        [void]$promptFiles.Add([ordered]@{
            order = $index
            path = if ($RedactPaths) { $fileName } else { $promptPath }
            chunk_id = $chunk.chunk_id
            sheet_name = $chunk.sheet_name
            range = $chunk.range
            cell_count = $chunk.payload.cell_count
        })
        $index++
    }

    $manifest = [ordered]@{
        generated_at = Get-TimestampJst
        scenario = $Scenario
        workbook_json_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedWorkbookJsonPath) } else { $resolvedWorkbookJsonPath }
        jsonl_path = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedJsonlPath) } else { $resolvedJsonlPath }
        prompt_count = $promptFiles.Count
        warnings = @($warnings)
        prompts = @($promptFiles)
    }

    Write-JsonFile -Data $manifest -Path $manifestPath
    Write-Host "prompt bundle を出力しました -> $resolvedOutputDir"
}
catch {
    Write-ErrorRecoverySteps -CommandName 'prompt bundle'
    throw "export_prompt_bundle.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
