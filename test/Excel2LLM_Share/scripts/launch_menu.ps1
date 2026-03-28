[CmdletBinding()]
param(
    [string]$ProjectRoot
)

if (-not $ProjectRoot) {
    $ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
}

$runAllBat = Join-Path $ProjectRoot 'tools\user\run_all.bat'
$promptBundleBat = Join-Path $ProjectRoot 'tools\user\run_prompt_bundle.bat'
$selfTestBat = Join-Path $ProjectRoot 'tools\user\run_self_test.bat'
$advancedDir = Join-Path $ProjectRoot 'tools\advanced'

Write-Host '=== Excel2LLM ==='
Write-Host '1. Excel を処理する'
Write-Host '2. 最新結果から指示文セットを作る'
Write-Host '3. 動作確認をする'
Write-Host '4. 詳細機能フォルダを開く'
Write-Host '5. 終了する'

$choice = Read-Host '番号を入力してください'

switch ($choice) {
    '1' {
        $excelPath = Read-Host 'Excel ファイルのパスを入力してください'
        if ([string]::IsNullOrWhiteSpace($excelPath)) {
            Write-Host 'Excel ファイルのパスが入力されていないため終了します。'
            exit 1
        }

        & $runAllBat $excelPath
        exit $LASTEXITCODE
    }
    '2' {
        & $promptBundleBat -Scenario general
        exit $LASTEXITCODE
    }
    '3' {
        & $selfTestBat
        exit $LASTEXITCODE
    }
    '4' {
        Start-Process explorer.exe $advancedDir | Out-Null
        exit 0
    }
    '5' {
        exit 0
    }
    default {
        Write-Host '1 から 5 の番号を入力してください。'
        exit 1
    }
}
