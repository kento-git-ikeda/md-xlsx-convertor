# md-xlsx-convertor

Excelのテスト仕様書をmael Markdownに変換し、Markdownのテスト行を既存のExcelブックに反映するWindows GUIアプリケーションです。

## 主な機能

- `#` の表ヘッダーを含む `.xlsx` ワークシートを、同じフォルダーにMarkdownファイルとして出力します。
- mael Markdownファイルの内容を既存の `.xlsx` ブックに反映します。
- ファイルパスを直接入力するか、ファイル選択画面から指定できます。
- Excelで読み込む開始行、または更新先の開始行を指定できます。
- 自動選択ができない場合は、ワークシートを指定できます。
- maelと実行時依存パッケージを、単体で実行できるWindows実行ファイルにまとめます。
- ブック更新時に、他のワークシートと既存の書式を保持します。
- maelの`output`ディレクトリに無関係なファイルがある場合、それらを削除しません。

## 実行ファイルのビルド

必要な環境: Windows、Python 3.14、PowerShell、および固定バージョンのパッケージをインストールするためのネットワーク接続。

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

生成されるファイル:

```text
dist/md-xlsx-convertor.exe
dist/THIRD_PARTY_NOTICES.txt
dist/ThirdPartyLicenses/
```

ビルド依存パッケージのバージョンは`requirements-build.txt`で固定しています。

## ソースから実行

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
python app.py
```

ソースからMarkdownからExcelへの更新を実行する際、maelがインストールされていなければ、アプリケーションが固定バージョンの`mael==0.0.3.35`をインストールします。パッケージ化された実行ファイルにはmaelが含まれているため、別途インストールする必要はありません。

## 操作方法

### ExcelからMarkdownを生成

**ExcelからMarkdownを生成**を選び、`.xlsx`ファイルを入力するか参照して選択します。必要に応じてワークシートと変換開始行を指定して実行してください。開始行を指定しない場合、検出した`#`ヘッダーの次の行から変換します。空のセルと対応するH3見出しは出力されません。上書きを有効にしない限り、既存のMarkdownファイルは置き換えません。

### MarkdownからExcelを更新

**MarkdownからExcelを更新**を選び、`.md`ファイルを入力するか参照して選択します。必要に応じて更新先のブック、ワークシート、書き込み開始行を指定してください。MarkdownのH1を、maelが生成するワークシートの選択に使用します。ブックを指定しない場合、同じフォルダーにある互換性のある`.xlsx`ファイルが1つであれば自動選択します。実行前に更新対象のブックを閉じてください。

更新に成功すると、maelが生成した`output`ディレクトリを削除します。そこに無関係なファイルやサブディレクトリがある場合は、それらを削除せずに処理を中止します。

## テスト

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests -v
```

## プロジェクト構成

```text
app.py
pyproject.toml
src/md_xlsx_convertor/
  excel_to_markdown.py
  gui.py
  markdown_to_excel.py
scripts/collect_licenses.py
tests/test_excel_to_markdown.py
tests/test_markdown_to_excel.py
md-xlsx-convertor.spec
requirements-build.txt
build.ps1
```

## ライセンス

このアプリケーションは、MIT、Apache-2.0/BSD、BSD-3-Clause、およびブートローダー例外付きのPyInstaller GPLv2以降のライセンスで提供されるサードパーティパッケージを使用しています。ビルド成果物には、パッケージのライセンス文書とNOTICEが含まれます。
