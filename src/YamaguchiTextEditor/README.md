# Yamaguchi Text Editor

山口さん専用テキストエディタです。

## 機能一覧

### テキストモード
- **ファイル操作**: 新規作成、開く、上書き保存、名前を付けて保存
- **ドラッグ＆ドロップ**: ファイルをウィンドウにドロップして開く（文字コード自動判定）
- **編集機能**: 元に戻す、やり直し、切り取り、コピー、貼り付け、すべて選択
- **特殊文字表示**: 空白・改行コードの可視化（Ctrl+Shift+Hで切替、デフォルトON）
  - 半角空白 → `_`（アンダースコア）
  - 全角空白 → `□`（正方形）
  - タブ → `→`（右矢印）
  - LF改行 → `↓`（下向き矢印）
  - CR改行 → `←`（左向き矢印）
- **検索・置換**: 文字列検索（前方/後方）、置換、全置換、正規表現対応
- **行番号表示**: 編集領域の左側に常時表示
- **カーソル位置表示**: ステータスバーに行・列・選択文字数を表示
- **文字コード表示**: カーソル位置の文字のUnicodeコードポイントをステータスバーに表示
- **文字コード対応**: UTF-8、UTF-8(BOM)、Shift-JIS、EUC-JP、JIS、UTF-16 LE/BE
- **フォント変更**: 任意のフォントに変更可能
- **行ジャンプ**: 指定行へ移動

### 表モード（CSV/TSV対応）
- **表形式表示**: CSV/TSVファイルを表形式で表示・編集
- **区切り文字選択**: カンマ(,)またはタブを選択可能
- **囲み文字選択**: ダブルクォーテーション有り/無しを選択可能
- **1行目ヘッダー**: 1行目を列ヘッダーとして扱うオプション
- **列選択**: 列ヘッダークリックで列全体を選択
- **列コピー**: 選択した列をクリップボードにコピー
- **列検索・置換**: 指定列内での検索・置換（正規表現対応）
- **セル編集**: セルをダブルクリックして直接編集
- **行の絞り込み**: 条件を指定して行を絞り込み表示（元の行番号を保持）
  - 完全一致 / を含む / で始まる / で終わる
  - 文字数以上 / 文字数以下

## キーボードショートカット

| ショートカット | 機能 |
|---------------|------|
| Ctrl+N | 新規作成 |
| Ctrl+O | ファイルを開く |
| Ctrl+S | 上書き保存 |
| Ctrl+Shift+S | 名前を付けて保存 |
| Ctrl+Z | 元に戻す |
| Ctrl+Y | やり直し |
| Ctrl+X | 切り取り |
| Ctrl+C | コピー |
| Ctrl+V | 貼り付け |
| Ctrl+A | すべて選択 |
| Ctrl+F | 検索 |
| F3 | 次を検索 |
| Shift+F3 | 前を検索 |
| Ctrl+H | 置換 |
| Ctrl+G | 指定行へ移動 |
| Ctrl+T | 表モード切り替え |
| Ctrl+L | 列の検索・置換 |
| Ctrl+Shift+H | 特殊文字表示切り替え |

## 動作環境

- Windows 10/11
- .NET Framework 4.8.1

## ビルド手順（Visual Studio 2022）

1. Visual Studio 2022を起動
2. 「ファイル」→「開く」→「プロジェクト/ソリューション」
3. `YamaguchiTextEditor.csproj` を選択して開く
4. 「ビルド」→「ソリューションのビルド」(Ctrl+Shift+B)
5. ビルド成功後、`bin\Debug\YamaguchiTextEditor.exe` が生成される

## プロジェクト構成

```
YamaguchiTextEditor/
├── Program.cs              # エントリーポイント
├── MainForm.cs             # メインフォーム（エディタ本体）
├── SearchForm.cs           # 検索・置換ダイアログ（モードレス）
├── GoToLineForm.cs         # 行ジャンプダイアログ
├── LineNumberPanel.cs      # 行番号表示用カスタムパネル
├── SpecialCharRichTextBox.cs # 特殊文字表示対応RichTextBox
├── EditorSettings.cs       # エディタ設定クラス
├── EncodingSelectForm.cs   # 文字コード選択ダイアログ
├── EncodingDetector.cs     # 文字コード自動判定クラス
├── TableModePanel.cs       # 表モード用パネル
├── ColumnSearchForm.cs     # 列検索・置換ダイアログ
├── CsvParser.cs            # CSV/TSVパーサー
├── YamaguchiTextEditor.csproj # プロジェクトファイル
├── app.config              # アプリケーション設定
└── README.md               # このファイル
```

## 各ファイルの説明

| ファイル | 説明 |
|---------|------|
| Program.cs | アプリケーションのエントリーポイント。コマンドライン引数対応 |
| MainForm.cs | メインウィンドウ。メニュー、ツールバー、エディタ領域を含む |
| SearchForm.cs | モードレス検索・置換ダイアログ。正規表現対応 |
| GoToLineForm.cs | 指定行へ移動するダイアログ |
| LineNumberPanel.cs | 行番号を表示するカスタムパネル |
| SpecialCharRichTextBox.cs | 特殊文字（空白・改行）を可視化するRichTextBox |
| EditorSettings.cs | フォント、文字コード、改行コードなどの設定を保持 |
| EncodingSelectForm.cs | ファイルを開く際の文字コード選択ダイアログ |
| EncodingDetector.cs | 文字コード自動判定クラス |
| TableModePanel.cs | 表モード用パネル（DataGridView含む） |
| ColumnSearchForm.cs | 表モードでの列検索・置換ダイアログ |
| CsvParser.cs | CSV/TSVファイルのパース・出力ユーティリティ |

## 注意事項

- このプロジェクトはDesignerファイル（*.Designer.cs）を使用していません
- すべてのUIはコードで生成されています
- 外部ライブラリは使用していません（.NET Framework標準のみ）

## バージョン

- Version 1.4.0 - 文字コード表示機能追加、表モードに絞り込み機能追加
- Version 1.3.0 - 特殊文字表示機能追加、表モードのタブ区切り修正
- Version 1.2.0 - ドラッグ＆ドロップ・文字コード自動判定機能追加
- Version 1.1.0 - 表モード機能追加
