# Excel2LLM

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

あなたの Excel ファイルを、ChatGPT などの LLM に渡しやすい形式へ変換するツールです。
Excel を開いて手でコピペしなくても、コマンド 1 つで必要なデータを取り出せます。

利用者向けの手順は `GETTING_STARTED.md` に統合しました。

## 3ステップで使う

### 1. Excel を `Excel2LLM.bat` にドラッグアンドドロップする

```bat
Excel2LLM.bat "C:\Data\book.xlsx"
```

### 2. 重要なら検証も入れる

```bat
Excel2LLM.bat "C:\Data\book.xlsx" -Verify
```

### 3. 必要なら prompt bundle を作る

```bat
Excel2LLM.bat -PromptBundle -Scenario general
```

主に使う出力:

- `output\<ファイル名_実行日時>\preflight_report.json`
- `output\<ファイル名_実行日時>\workbook.json`
- `output\<ファイル名_実行日時>\llm_package.jsonl`
- `output\<ファイル名_実行日時>\verify_report.json`

## このあと読む文書

- `GETTING_STARTED.md`
  - 利用者向けの手順をまとめた唯一の案内
- `docs/reference/LLM_PROMPT_FORMATS.md`
  - LLM へ渡すときの用途別テンプレート
- `docs/README.md`
  - 文書全体の案内

## 主要コマンド

```bat
Excel2LLM.bat "C:\path\to\book.xlsx"
Excel2LLM.bat "C:\path\to\book.xlsx" -Verify
Excel2LLM.bat -PromptBundle -Scenario general
```

詳細機能を直接使いたい場合は `tools\user\` と `tools\advanced\` の `bat` を使えます。

## できること

- 複数シートのセル値、表示値、数式、結合セルを `workbook.json` に保存
- 色や罫線などの補助情報を `styles.json` に分離保存
- LLM 向けの `llm_package.jsonl` を生成
- `prompt_*.txt` の prompt bundle を生成
- 抽出結果と Excel 再計算結果の差分を `verify_report.json` で確認
- `workbook.json` から `.xlsx` を逆生成

## セキュリティ上の既定動作

- `extract` の前に必須の preflight を行い、重すぎる Excel や破損疑いのある Excel は抽出開始前に停止します
- `extract` と `verify` は、既定で Excel ブックマクロを無効化して開きます
- 絶対パスを減らしたい場合は `-RedactPaths` を使います
- 配布用フォルダ再生成は、既定で `distribution\` 配下のみ安全に削除します

## 補足

- 対応形式は `.xlsx` と `.xlsm`
- 色や罫線は補助情報であり、主処理の成功条件ではありません
- 逆生成の出力は常に `.xlsx` です
- ルート直下の入口は `Excel2LLM.bat` だけです
