# バッチ／PowerShellファイル可逆変換システム

- 作成日: 2026-03-28 17:42 JST
- 更新日: 2026-04-24
- 作成者: Codex (GPT-5)

## 概要

`.bat`、`.ps1`、`.psm1`、`.md`、`.json`、`.jsonl`、`.bas` を 1 つの UTF-8 JSON テキストへ集約し、その集約ファイルから元の複数ファイルとフォルダ構成を復元する Windows 向けツールです。起動は `.bat`、内部処理は PowerShell で行います。

固定フォルダ運用に加えて、任意の入出力パス指定、`.bundleignore`、`verify`、機械可読な結果 JSON 出力に対応します。

## ファイル構成

- `bundle_files.bat`: 集約または検証を実行します。
- `restore_files.bat`: 復元またはフォルダ構成のみ復元を実行します。
- `bundle_launcher.ps1`: `.bat` 起動時のモード判定、引数透過、一時停止制御を行います。
- `bundle_system.ps1`: 集約・復元・検証の共通処理です。
- `tests\BundleSystem.Tests.ps1`: Pester テストです。

## 対応拡張子

- `.bat`
- `.ps1`
- `.psm1`
- `.md`
- `.json`
- `.jsonl`
- `.bas`

## 従来の固定フォルダ構成

引数を省略した場合は、従来通り次のフォルダを使います。存在しなければ自動作成します。

- `input_files`
- `output_bundle`
- `restore_input`
- `restore_output`

## 主要な使い方

### 1. 従来運用のまま集約

```cmd
bundle_files.bat
```

### 2. 任意のフォルダを直接集約

```cmd
bundle_files.bat --no-pause -InputPath "C:\Work\MyProject" -OutputPath "D:\bundle_out"
```

### 3. `.bundleignore` を使って集約

`.bundleignore` を入力ルート直下に置くか、明示的に `-IgnoreFilePath` を渡します。

```cmd
bundle_files.bat --no-pause -InputPath "C:\Work\MyProject" -OutputPath "D:\bundle_out" -IgnoreFilePath "C:\Work\MyProject\.bundleignore"
```

### 4. 復元

```cmd
restore_files.bat --no-pause -RestoreInputPath "D:\bundle_out\bundle_260424_MyProject.txt" -RestoreOutputPath "D:\restore_out"
```

`-RestoreInputPath` は、`bundle*.txt` を 1 件だけ置いたフォルダでも、個別の bundle ファイルでも指定できます。

### 5. フォルダ構成のみ復元

```cmd
restore_files.bat structure --no-pause -RestoreInputPath "D:\bundle_out\bundle_260424_MyProject.txt" -RestoreOutputPath "D:\restore_out"
```

### 6. 往復検証

```cmd
bundle_files.bat verify --no-pause -InputPath "C:\Work\MyProject" -OutputPath "D:\bundle_out" -RestoreOutputPath "D:\verify_restore"
```

## PowerShell 引数

`bundle_system.ps1` は次のモードに対応します。

- `-Mode Bundle`
- `-Mode Restore`
- `-Mode RestoreStructure`
- `-Mode Verify`

主な引数:

- `-InputPath`: 集約または検証対象の入力ルート
- `-OutputPath`: bundle 出力先ディレクトリ
- `-RestoreInputPath`: 復元対象の bundle ファイル、またはその格納ディレクトリ
- `-RestoreOutputPath`: 復元先ディレクトリ
- `-IgnoreFilePath`: `.bundleignore` の明示パス
- `-BundleRootName`: bundle 内の相対パス先頭に付けるルートフォルダ名
- `-ResultJsonPath`: 実行結果サマリを JSON で出力するパス

## `.bundleignore`

入力ルート直下の `.bundleignore` を自動検出します。明示パスを渡した場合はそちらを優先します。

書式:

- 空行は無視
- `#` で始まる行はコメント
- `node_modules` のような単純名は、同名セグメントを含む配下を除外
- `dist/output` のような相対パスは、その配下を除外
- `docs/*.md` のようなワイルドカードも利用可能

例:

```text
# build artifacts
node_modules
dist
coverage
docs/private
```

## 出力 bundle JSON の追加メタデータ

bundle ファイルには次を含めます。

- `sourceRoot`
- `excludedDirectories`
- `toolVersion`
- `createdBy`
- `hostname`

従来の `format`, `version`, `directories`, `files` なども維持します。

## 結果 JSON

`-ResultJsonPath` を指定すると、実行結果サマリを UTF-8 JSON で出力します。成功時は次のような情報を含みます。

- `Status`
- `Operation`
- `ExitCode`
- `ToolVersion`
- `SourceRoot`
- `BundlePath`
- `RestoreOutputPath`
- `BundledFileCount`
- `BundledDirectoryCount`
- `VerifiedFileCount`
- `VerifiedDirectoryCount`

失敗時も `Status=Error` と `Message` を含む JSON を出力します。

## 仕様上のポイント

- bundle 本体は UTF-8 JSON テキストです。
- 各ファイル本文は Base64 で保持し、復元時は生バイト列をそのまま書き戻します。
- 通常復元ではフォルダ構成を先に再構築してからファイルを書き戻します。
- `structure` モードではフォルダ構成のみ復元し、ファイルは生成しません。
- `verify` モードでは bundle 作成、復元、SHA-256 とフォルダ構成の照合を連続で行います。
- 復元前に `id`, `relativePath`, `fileName`, `extension`, `byteLength`, `sha256`, `newlineStyle`, `bomType` を検証します。
- `.bat` 起動時は処理結果を確認できるよう、終了前に一時停止します。自動実行したい場合は `--no-pause` を付けます。

## 上書き方針

- 復元先に同名ファイルが 1 件でも存在した場合は、何も復元せずに停止します。
- 親ディレクトリ位置のファイル衝突も検出して停止します。
- 標準動作で上書きは行いません。

## 制約事項

- 本ツールは暗号化、圧縮、署名付与、ネットワーク連携、GUI を行いません。
- 対象外ファイルが混在していても、対象内ファイルが 1 件以上あれば一覧表示したうえでスキップします。
- 対象外ファイルしか存在しない場合は終了コード `11` で停止します。
- 復元時は危険な相対パス、絶対パス、予約名、禁止文字を含むパスを拒否します。

## 終了コード

| 終了コード | 内容 |
| --- | --- |
| `0` | 正常終了 |
| `10` | 集約対象 0 件 |
| `11` | 対象外ファイルのみ |
| `12` | 読み取り失敗 |
| `13` | 書き込み失敗 |
| `20` | 集約ファイル未検出 |
| `21` | 集約ファイル複数件 |
| `22` | フォーマット不正 |
| `23` | パス不正 |
| `24` | 復元先衝突 |
| `25` | verify 照合失敗 |
| `30` | 権限不足 |
| `99` | 想定外エラー |

## テスト

Pester 3.4 で次を検証します。

- 任意パスからの bundle 作成
- `.bundleignore`
- `-BundleRootName` による復元ルート保持
- bundle / restore のバッチ引数透過
- `!`, `&`, `%` を含むパスのバッチ引数透過
- UTF-8 BOM / UTF-8 / UTF-16 LE / CP932 の往復一致
- `verify` モード

実行例:

```powershell
Invoke-Pester -Path .\tests\BundleSystem.Tests.ps1
```
