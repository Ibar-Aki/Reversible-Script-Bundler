# VBA 補助モジュール

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)

`templates/Excel2LLM_Helper.bas` は任意の補助モジュールです。主処理は PowerShell のみで完結しますが、次の用途で使えます。

- `CalculateFullRebuild` を Excel 側で手動実行する
- 表示値確認のために TSV を吐き出す
- PowerShell 抽出との差分確認を手動で補う

## 取り込み手順

1. Excel を開く
2. `Alt + F11` で VBA エディターを開く
3. 対象ブックで `ファイル > ファイルのインポート` を選び、`templates/Excel2LLM_Helper.bas` を読み込む
4. `Excel2LLM_RecalculateWorkbook` または `Excel2LLM_ExportDisplaySnapshot` を実行する

## 注意

- `Excel2LLM_ExportDisplaySnapshot` は UTF-16 の TSV を出力します
- マクロ実行ポリシーや VB プロジェクトへのアクセス権は組織設定の影響を受けます
- 条件付き書式の最終見た目や画面状態の確認は Excel UI 上で行ってください
