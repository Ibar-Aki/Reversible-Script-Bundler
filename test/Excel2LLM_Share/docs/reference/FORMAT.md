# データ形式

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## preflight_report.json

`preflight_report.json` は抽出前の事前チェック結果です。

- `status`: `success | warning | blocked`
- `blocked`: 停止したかどうか
- `workbook_path`: 元 Excel のパス。`-RedactPaths` 指定時はファイル名
- `output_directory`: 出力先。`-RedactPaths` 指定時は末尾ディレクトリ名
- `file_size_bytes`
- `sheet_count`
- `estimated_total_cells`
- `largest_sheet`
- `reasons`: 停止理由
- `warnings`: 警告一覧
- `sheets[]`: シートごとの推定 dimension と推定セル数

## workbook.json

`workbook.json` は主データです。トップレベルは以下の構造です。

```json
{
  "workbook": {},
  "sheets": [],
  "cells": [],
  "merged_ranges": [],
  "generated_at": "",
  "generator": ""
}
```

### workbook

- `name`: ファイル名
- `path`: 元ファイルの絶対パス。`-RedactPaths` 指定時はファイル名
- `extension`: `.xlsx` または `.xlsm`
- `sheet_count`: 抽出結果に含まれるシート数
- `has_vba`: VBA を含む可能性があるか

### sheets[]

- `sheet_name`
- `sheet_index`
- `visible`
- `used_range`
- `freeze_panes`
- `hidden_rows`
- `hidden_columns`
- `row_heights`
- `column_widths`
- `cell_count`
- `formula_count`
- `merged_ranges`

### cells[]

- `sheet`
- `address`
- `row`
- `column`
- `value2`
- `text`
- `formula`
- `formula2`
- `has_formula`
- `number_format`
- `merge_area`
- `is_merge_anchor`
- `comment`
- `comment_threaded`
- `hyperlink`

## styles.json

`styles.json` は補助情報です。既定では空またはスキップ状態で生成し、`-CollectStyles` 指定時だけ best effort で内容を埋めます。取得失敗時も主処理は継続します。

```json
{
  "generated_at": "",
  "generator": "",
  "styles": []
}
```

### styles[]

- `sheet`
- `address`
- `fill_color`
- `font_color`
- `horizontal_alignment`
- `vertical_alignment`
- `wrap_text`
- `borders`

## llm_package.jsonl

1 行 1 チャンクの JSONL です。

- `chunk_id`
- `sheet_name`
- `range`
- `cell_addresses`
- `payload`
- `formula_cells`
- `token_estimate`
- `includes_styles`

## manifest.json

`manifest.json` は抽出全体の概要です。

- `status`
- `warnings`
- `workbook_path`: 元ファイルの絶対パス。`-RedactPaths` 指定時はファイル名
- `output_directory`: 出力先。`-RedactPaths` 指定時は末尾ディレクトリ名
- `source_sheet_count`: 元ブックのシート総数
- `sheet_count`: 抽出結果に含まれるシート数
- `sheet_filter.include`: `-Sheets` で指定した一覧
- `sheet_filter.exclude`: `-ExcludeSheets` で指定した一覧
- `sheet_filter.selected`: 実際に抽出されたシート一覧
- `cell_count`
- `formula_count`
- `merged_range_count`
- `style_export_status`
- `verify_status`

## verify_report.json

`verify_report.json` は再計算後の差分検証です。

- `workbook_path`: 元 Excel の絶対パス。`-RedactPaths` 指定時はファイル名
- `workbook_json_path`: 検証対象 JSON の絶対パス。`-RedactPaths` 指定時はファイル名
- `mismatch_count`
- `warnings`
- `mismatches`

## rebuild_report.json

`rebuild_report.json` は `workbook.json` から `.xlsx` を逆生成した結果の記録です。

```json
{
  "generated_at": "",
  "generator": "",
  "status": "",
  "warnings": [],
  "workbook_json_path": "",
  "styles_json_path": "",
  "output_path": "",
  "output_extension": ".xlsx",
  "source_has_vba": false,
  "restored_sheets": 0,
  "restored_cells": 0,
  "restored_formulas": 0,
  "restored_comments": 0,
  "restored_hyperlinks": 0,
  "restored_styles": 0,
  "restored_merged_ranges": 0,
  "threaded_comment_fallbacks": 0
}
```

### 注意

- `output_extension` は常に `.xlsx` です
- `source_has_vba=true` でも VBA 本体は復元されません
- `threaded_comment_fallbacks` はスレッドコメントを通常コメントへ落とした件数です
