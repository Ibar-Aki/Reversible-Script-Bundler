# 活用事例

- 作成日: 2026-03-10 01:30 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## 目的別の使い方

このドキュメントでは、`Excel2LLM` を実務でどう使うかを具体例でまとめます。

この文書では、`Excel2LLM.bat` 実行後に作られる最新の実行結果フォルダを `output\<実行結果フォルダ>\` と表記します。

## 事例 1: 見積もり表の内容を LLM に説明させる

### 状況

- 3 シートの見積もりブックがある
- 単価、数量、合計が数式で計算されている
- 担当者は「どの列が何を意味しているか」をすぐ説明したい

### 実行手順

```bat
Excel2LLM.bat "C:\Data\estimate.xlsx"
```

### LLM への渡し方

- `llm_package.jsonl` から見積もりシートのチャンクを選ぶ
- `formula_cells` を見ながら、合計計算のセルを重点的に渡す

### 期待する出力

- 各列の意味
- 合計の計算方法
- どの行が高額か
- 数量や単価の違和感

## 事例 2: 監査前に数式の意味を一覧化する

### 状況

- 管理会計ブックに多数の数式がある
- 監査前に「どのセルが何を計算しているか」を説明文にしたい

### 実行手順

```bat
Excel2LLM.bat "C:\Data\finance.xlsx" -Verify
tools\advanced\run_pack.bat "output\<実行結果フォルダ>\workbook.json" -ChunkBy range -MaxCells 200
```

### ポイント

- 先に `verify_report.json` を見て、差分がないことを確認する
- `formula` と `formula2` を LLM に読ませる

### 期待する出力

- 数式セルの一覧
- 入力セルと出力セルの関係
- 複雑な依存関係の説明

## 事例 3: 大きな在庫表を分割して分析する

### 状況

- 50 行 100 列クラスの表がある
- 1 回のプロンプトでは大きすぎる
- 列ごとの異常値や欠損を見たい

### 実行手順

```bat
Excel2LLM.bat "C:\Data\inventory.xlsx"
tools\advanced\run_pack.bat "output\<実行結果フォルダ>\workbook.json" -ChunkBy range -MaxCells 150
```

### ポイント

- `range` 分割でサイズを均一にする
- LLM には 1 チャンクずつ順番に渡す
- 最後にチャンクごとの所見をまとめて統合する

### 期待する出力

- 異常値の候補
- 空欄や欠損の傾向
- 列定義の推定

## 事例 4: 複数シートの業務ブックを FAQ 化する

### 状況

- 申請一覧、マスタ、計算シートが分かれている
- 社内メンバーが「このセルは何か」「このシートの役割は何か」をよく聞く

### 実行手順

```bat
Excel2LLM.bat "C:\Data\operations.xlsx"
```

### 活用方法

- `sheet_name` と `range` をキーに検索できるようにする
- 必要チャンクだけ LLM に渡して回答させる

### 向いている質問

- このシートは何を管理しているか
- この数式は何を計算しているか
- この列は何を意味しているか

## 事例 5: 色や罫線も補助的に見たい

### 状況

- 基本は値と数式が重要
- ただし一部案件では、色付きセルや罫線位置も参考にしたい

### 実行手順

```bat
Excel2LLM.bat "C:\Data\report.xlsx" -CollectStyles
tools\advanced\run_pack.bat "output\<実行結果フォルダ>\workbook.json" -IncludeStyles
```

### ポイント

- style は補助情報として扱う
- まず `workbook.json` を主に見て、必要なときだけ style を足す

### 向いている用途

- ヘッダー行の推定
- 強調セルの把握
- 表区切りの補助解釈

## 事例 6: 再計算が怪しいブックを検証する

### 状況

- 開くたびに値が変わる
- 外部リンクや volatile 関数がある
- LLM に渡す前に安全確認したい

### 実行手順

```bat
Excel2LLM.bat "C:\Data\volatile.xlsx" -Verify
```

### 見るべきファイル

- `manifest.json`
- `verify_report.json`

### 期待する運用

- 差分 0 件ならそのまま使う
- 差分ありなら、差分セルだけ再確認してから LLM に渡す

## 事例 7: LLM への指示文を標準化する

### 目的

- 担当者ごとに質問の質がぶれないようにする

### 例 1: シート要約

```text
以下は Excel チャンクです。sheet_name、range、cells を使って、シートの目的、主要列、主要数式、注意すべき値を要約してください。
```

### 例 2: 数式レビュー

```text
以下の Excel チャンクについて、formula と formula2 を確認し、各数式の計算意図、依存セル、誤入力リスクを説明してください。
```

### 例 3: データ品質チェック

```text
以下の Excel チャンクを見て、空欄、重複、不自然な値、列の意味の不整合を指摘してください。text と value2 の差がある場合はその理由も推定してください。
```

## どの事例でも共通の運用

- 最初は `workbook.json` を正本として保存する
- まずは `Excel2LLM.bat` で `workbook.json` と `llm_package.jsonl` をまとめて作る
- LLM に渡すのは `llm_package.jsonl` の必要チャンクだけにする
- 重要なブックは `Excel2LLM.bat -Verify` を使う
- style は本当に必要な案件だけ追加する
