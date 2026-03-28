# Excel2LLM はじめに

- 作成日: 2026-03-28 00:20 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## この文書だけ読めば使えます

利用者向けの手順は、この文書に統合しました。
まずはこの文書のとおりに進めてください。

## 使う前に必要なもの

| 項目 | 内容 |
| --- | --- |
| OS | Windows 10 / 11 |
| Excel | Microsoft 365 Excel（デスクトップ版） |
| 追加インストール | 不要 |
| 主な使い方 | `Excel2LLM.bat` へ Excel をドラッグアンドドロップ |

## まずやること

| 手順 | 何をするか | 補足 |
| --- | --- | --- |
| 1 | 配布フォルダをそのまま任意の場所へ置く | フォルダ名は変えてもかまいません |
| 2 | `Excel2LLM.bat` をダブルクリックして「3. 動作確認」を選ぶ | 最初の 1 回だけで十分です |
| 3 | `セルフテストが正常終了しました。` と出ることを確認する | これで使える状態です |
| 4 | 自分の Excel を `Excel2LLM.bat` にドラッグアンドドロップする | これが基本の使い方です |

## 基本の使い方

### いちばん簡単な使い方

1. 処理したい Excel ファイルを見つける
2. その Excel ファイルを `Excel2LLM.bat` の上へドラッグアンドドロップする
3. 処理が終わるまで待つ
4. 画面に出た出力先フォルダを開く

コマンドで実行する場合:

```bat
Excel2LLM.bat "C:\Data\book.xlsx"
```

重要な資料で、Excel との突き合わせもしたい場合:

```bat
Excel2LLM.bat "C:\Data\book.xlsx" -Verify
```

## 実行すると何が作られるか

実行するたびに、`output` の中へ **新しい実行結果フォルダ** が作られます。
フォルダ名は **ファイル名 + 実行日時** です。

例:

```text
output\estimate_20260328-143500
```

同じ Excel を何回実行しても、前回結果を上書きしにくい作りです。

## 出力ファイル一覧

| ファイル / フォルダ | いつできるか | 意味 |
| --- | --- | --- |
| `preflight_report.json` | `Excel2LLM.bat` / `tools\advanced\run_extract.bat` / `tools\advanced\run_preflight.bat` | 事前チェック結果 |
| `workbook.json` | `Excel2LLM.bat` / `tools\advanced\run_extract.bat` | Excel 全体を保存した正本 |
| `styles.json` | `-CollectStyles` を付けたとき | 色や罫線などの補助情報 |
| `manifest.json` | `Excel2LLM.bat` / `tools\advanced\run_extract.bat` | 抽出結果の要約 |
| `llm_package.jsonl` | `Excel2LLM.bat` / `tools\advanced\run_pack.bat` | LLM に渡しやすい分割済みデータ |
| `verify_report.json` | `-Verify` または `tools\advanced\run_verify.bat` | Excel との突き合わせ結果 |
| `prompt_bundle\` | `-PromptBundle` | LLM に貼り付ける指示文セット |
| `rebuilt\` | `tools\advanced\run_rebuild.bat` | `workbook.json` から作り直した Excel |

## 最初に見るべきもの

| 見るもの | 何を見るか |
| --- | --- |
| `workbook.json` | Excel の内容が取れているか |
| `llm_package.jsonl` | LLM に渡すデータができているか |
| `verify_report.json` | 差分があるかどうか |
| `preflight_report.json` | 重すぎる・壊れているなどで止まっていないか |

## よく出る言葉

| 用語 | 意味 |
| --- | --- |
| `preflight` | Excel を開く前の事前チェック |
| `workbook.json` | Excel 全体の内容を保存したファイル |
| `JSONL` | 1 行に 1 件ずつデータが入る形式 |
| `チャンク` | LLM に一度に渡すデータのかたまり |
| `prompt bundle` | LLM に貼り付ける指示文セット |
| `verify` | 抽出結果と Excel を突き合わせる確認 |

## 使い分け表

| やりたいこと | 使うもの |
| --- | --- |
| まず普通に使いたい | `Excel2LLM.bat` にドラッグアンドドロップ |
| 重要な Excel を慎重に扱いたい | `Excel2LLM.bat` + `-Verify` |
| 危険な Excel か先に確認したい | `tools\advanced\run_preflight.bat` |
| LLM に貼り付ける文を作りたい | `Excel2LLM.bat -PromptBundle` |
| `workbook.json` から Excel を作り直したい | `tools\advanced\run_rebuild.bat` |

## LLM に渡すまでの流れ

| 手順 | 何をするか |
| --- | --- |
| 1 | `Excel2LLM.bat` で `llm_package.jsonl` を作る |
| 2 | `llm_package.jsonl` をテキストエディタで開く |
| 3 | 必要な行だけをコピーする |
| 4 | ChatGPT などへ貼り付ける |
| 5 | 必要なら `Excel2LLM.bat -PromptBundle` で指示文も作る |

LLM 向けの指示文例は `docs\reference\LLM_PROMPT_FORMATS.md` にあります。

## 困ったとき

| 状況 | まずやること |
| --- | --- |
| うまく動かない | Excel を閉じて、もう一度実行する |
| 途中で止まった | `preflight_report.json` を見る |
| 差分が出た | `verify_report.json` を見る |
| まだだめ | `Excel2LLM.bat -SelfTest` を実行する |

## 補足

- `Excel2LLM.bat` は、抽出前に自動で事前チェックを行います
- 重すぎる Excel や壊れている疑いがある Excel は、Excel を開く前に停止します
- `Excel2LLM.bat -PromptBundle` は、直前の実行結果フォルダを自動で使います
- 詳細機能を直接使いたい場合は `tools\advanced\` の `bat` も利用できます
- 絶対パスを減らしたい場合は `-RedactPaths` を使います
