# StandardDocApp

`stdharvest`（収集）と `stdsearch`（検索）を 1 つの Windows GUI に統合した Tk
アプリケーションです。エンドユーザーは `StandardDocApp.exe` だけを起動し、
画面上のタブで「収集」と「検索」を切り替えて利用できます。

- 入口は **`StandardDocApp.exe`** に統一（`stdharvest.exe` / `stdsearch.exe`
  は作成しません）。
- 既存の CLI は引き続き利用できます:
  - `stdharvest run --excel ...`
  - `stdsearch run --excel ...`
- 既存の Excel 仕様、出力フォルダ構成、ログ仕様は変更していません。
- PDF 化は **Microsoft Office を優先**し、利用不可なら **LibreOffice
  (soffice)** にフォールバックします（`stdharvest/pdf_converter.py` 既存実装
  を再利用）。Playwright / Chromium には依存しません。

## タブ構成

### 収集タブ
- `sample_download.xlsx` 形式の Excel を選択
- `Sheet1!B1:B3` から SourceType / OutputRootFolder / JobName を表示
- 予測される `JobFolder`（`<OutputRoot>\<YYYYMMDD>_<JobName>`）を表示
- 「Run harvest」で `stdharvest.cli.run` を別スレッドで実行
- ログをリアルタイム表示（ファイルにも保存）
- 完了後、JobFolder と `html/index.html` を OS の関連付けで開ける

### 検索タブ
- `sample_search.xlsx` 形式の Excel を選択
- ProjectName / JobFolder / OutputFolder と検索ルール一覧を表示
- `manifest.json` の存在確認結果を表示
- 「Run search」で `stdsearch.cli.run` を別スレッドで実行
- 完了後、`search_results.html` / `.csv` / 出力フォルダを開ける

### 設定・ログ・About タブ
- アプリバージョン / Python / OS 情報
- Microsoft Office / LibreOffice 検出状況（PDF 変換に使う優先順位を表示）
- アプリログ保存先（`%LOCALAPPDATA%\StandardDocApp\logs`）
- **サンプル Excel ジェネレータ（カスタムダイアログ）**:
  - "Build sample download.xlsx... (3GPP / IEEE)" を押すとモーダルダイアログが開き、
    SourceType / 保存先フォルダ / ファイル名 / JobName / ProxyURL /
    OutputRootFolder / 参照元 Excel をその場で指定できます。
  - **OutputRootFolder の初期値はユーザーの Downloads フォルダ**
    (`%USERPROFILE%\Downloads`) に固定。
  - 未指定項目は既存サンプルのデフォルト値が使われます (JobName は
    `<source>_sample_job`、ProxyURL は空、URL リストは下記デフォルト)。
  - 3GPP デフォルト URL: TSGR1_120b の R1-2502624 / 2502715 / 2502726 /
    2502776 / 2502814 / 2502821 (.zip 6 件)
  - IEEE デフォルト URL: 802.11 mentor の mu-edca / edca rules 系 6 件
    (11-16-1424 / 1425 / 1368 / 0998 / 0963 / 0962)
  - **参照元 Excel** を指定すると、その A 列をスキャンして:
    - `cell.hyperlink.target` があればそれを URL、セル文字列を title に
    - なければセル文字列を URL（http(s)/ftp で始まる場合）として採用
    - 結果を `sample_download.xlsx` の Sheet1 C 列に **ハイパーリンク付きで** 書き込み。
  - 検索用 `sample_search.xlsx` も「Build sample search.xlsx...」ボタンで作成可能。
- **HTML 整合性チェッカ**:
  - 任意の job フォルダを選択 → `html/index.html` / `html/files/**/index.html` /
    `html/combined/*.html` / `html/combine_full/*.html` を **読み取り専用** で検査
  - 構文不正、`<body>` 欠如、`<title>` 空、`<!doctype>` 欠如などを検出
  - 内部リンク (`<a href>`, `<img src>`, …) と内部アンカー (`#id`) の整合性チェック
  - 件数進捗バー + ログ + ファイル保存
- リポジトリ内のサンプル Excel の場所をエクスプローラーで開く
- README を OS 既定のビューワで開く

## Open ボタンの保証

収集タブ・検索タブの「Open job folder」「Open html/index.html」「Open
search_results.{html,csv}」「Open output folder」は、**ボタンを押した瞬間に
現在指定されている Excel を読み直して**、`OutputRootFolder / JobName /
ProjectName / OutputFolder` から正しいパスを再解決します。Excel を変更しても、
古いキャッシュではなく現在の設定が指す場所が必ず開きます。

- 収集: `<OutputRootFolder>/<YYYYMMDD>_<JobName>` を試し、なければ
  `<OutputRootFolder>/*_<JobName>` の最新 mtime のフォルダを開きます。
- 検索: `<OutputFolder or JobFolder/search>/<YYYYMMDD>_<ProjectName>` を試し、
  なければ同様に最新 mtime のフォルダを開きます。

加えて、**Excel ファイルが外部エディタで保存される（mtime 更新）と自動で
プレビューが再読み込みされ**、いつでも「Reload」ボタンで強制再読み込みできます。

## 開発実行（exe を作らずに動かす）

```powershell
cd <repo-root>
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .\stdharvest
python -m pip install -e .\stdsearch
python -m pip install -e .\standarddocapp
python -m standarddocapp
```

## exe 形式でビルドする方法 (PyInstaller)

PyInstaller で **コンソールが出ない単一の Windows GUI 実行ファイル** を作ります。
ビルド前に、まず開発実行（`python -m standarddocapp`）で GUI が起動できることを
確認してください。

### 推奨方法 ① bat スクリプト（PowerShell 実行ポリシー回避済み）

リポジトリ直下に `build_exe.bat` を置いてあります。コマンドプロンプトでも
エクスプローラーからのダブルクリックでも動きます。

```bat
REM リポジトリ直下から
build_exe.bat

REM venv を再利用してビルドだけ走らせたい場合
build_exe.bat -SkipDeps

REM dist\ / build\ をクリーンしてから走らせたい場合
build_exe.bat -Clean
```

中身は `standarddocapp\build_tools\build.bat` を呼ぶだけで、
そちらは `powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1`
を実行します。よって PowerShell の実行ポリシー
（`...スクリプトの実行が無効になっているため...` / `UnauthorizedAccess`）
の影響を受けません。

### 推奨方法 ② PowerShell から直接

事前に **一度だけ** 実行ポリシーを許可してから走らせる方法です。

```powershell
# 一度だけ（CurrentUser スコープ。管理者権限不要）
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# 以降は普通に
.\standarddocapp\build_tools\build.ps1
```

その場限りで通したい場合は次でも可:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
    .\standarddocapp\build_tools\build.ps1
```

> **なぜエラーが出るのか:** Windows の既定では、署名されていない `.ps1` の
> 直接実行は `Restricted` でブロックされます。`build.bat` 経由か、
> 上記いずれかの方法で `Bypass` / `RemoteSigned` 相当の実行コンテキストを
> 与えてください。

### 推奨方法 ③ Python から直接

```powershell
python .\standarddocapp\build_tools\build_app.py
# venv を再利用するなら:
python .\standarddocapp\build_tools\build_app.py --skip-deps
# 出力をクリーンにしてから走らせたい:
python .\standarddocapp\build_tools\build_app.py --clean
```

`build.ps1` / `build.bat` が内部で行うのと同じ手順です。

### `build.ps1` / `build_app.py` が内部で行うこと

1. `standarddocapp/.venv-build/` に隔離された venv を作成（既にあれば再利用）
2. その venv に `stdharvest` / `stdsearch` / `standarddocapp` を editable install
3. `pyinstaller>=6.0` を最新化
4. `StandardDocApp.spec` を実行して `dist/StandardDocApp.exe` を出力

### spec を使わず手動コマンドでビルドする場合

`StandardDocApp.spec` を使わず、`python -m PyInstaller` 直叩きでもビルドできます。
ポイントは次の 2 つです。

- **エントリは `__main__.py` ではなく `build_tools\standarddocapp_launcher.py`**
  （`__main__.py` を直接渡すと
  `ImportError: attempted relative import with no known parent package`
  になります）。
- `stdharvest` / `stdsearch` / `standarddocapp` のすべてを `--collect-submodules`
  しておくと動的 import を取りこぼしません。

```powershell
cd standarddocapp
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name StandardDocApp ^
  --paths src ^
  --paths ..\stdharvest\src ^
  --paths ..\stdsearch\src ^
  --collect-submodules standarddocapp ^
  --collect-submodules stdharvest ^
  --collect-submodules stdsearch ^
  --hidden-import openpyxl ^
  --hidden-import requests ^
  --hidden-import mammoth ^
  --hidden-import bs4 ^
  --hidden-import lxml ^
  --hidden-import win32com.client ^
  --add-data "src\standarddocapp\assets;standarddocapp\assets" ^
  build_tools\standarddocapp_launcher.py
```

アイコンを付ける場合は次を追加してください:

```text
--icon src\standarddocapp\assets\app.ico
```

### アイコン (`.exe` のアイコン) を変更する

`StandardDocApp.spec` は **`src\standarddocapp\assets\app.ico` が存在すれば
自動的にアイコンとして使う** ようになっています（無ければ PyInstaller
デフォルト）。 設定手順:

1. 任意の方法で `app.ico` を用意します。  
   PNG しか手元に無い場合は ImageMagick で生成可能:

   ```powershell
   magick convert app.png -define icon:auto-resize=256,128,64,48,32,16 app.ico
   ```

2. それを `standarddocapp\src\standarddocapp\assets\app.ico` に保存。
3. 再ビルド:

   ```bat
   build_exe.bat -Clean
   ```

詳細は [`src/standarddocapp/assets/README.md`](./src/standarddocapp/assets/README.md)
を参照してください。

### 出力先

```
standarddocapp\dist\StandardDocApp.exe
```

`--windowed` (spec 内では `console=False`) でビルドしているので、
**ダブルクリックしても黒いコンソール画面は出ません**。共有先の PC に
このファイルだけコピーして起動してください（事前セットアップ不要、
Microsoft Office または LibreOffice のいずれかが入っていれば PDF 化が動作します）。

### よくあるビルドエラー

| 症状 | 原因 | 対処 |
| --- | --- | --- |
| `このシステムではスクリプトの実行が無効になっているため...` / `UnauthorizedAccess` | PowerShell の実行ポリシーで `.ps1` がブロック | `build_exe.bat` を使う、もしくは `powershell -ExecutionPolicy Bypass -File ...` で起動 |
| `ImportError: attempted relative import with no known parent package` | `__main__.py` を直接 PyInstaller のエントリに渡している | エントリを `build_tools\standarddocapp_launcher.py` に変更（同梱の spec はこの形になっています） |
| `ModuleNotFoundError: No module named 'stdharvest'` | venv に `stdharvest` / `stdsearch` を editable install していない | `build_app.py` を `--skip-deps` 無しで再実行、または手動で `pip install -e ..\stdharvest -e ..\stdsearch` |
| `lxml._elementpath` 等が見つからない | PyInstaller の hidden import 取りこぼし | spec を使うか、手動コマンドの場合は `--hidden-import lxml._elementpath` を追加 |
| 起動直後に一瞬で落ちる | `console=False` で例外を見逃している | デバッグ時は spec 内 `console=True` にしてビルドし、コマンドプロンプトから起動して例外を確認 |

## ファイル構成

```
standarddocapp/
├── pyproject.toml
├── README.md
├── StandardDocApp.spec        # PyInstaller spec (icon 自動検出)
├── build_tools/
│   ├── build.bat              # PowerShell 実行ポリシー回避用 .bat ラッパ
│   ├── build.ps1              # PowerShell ラッパ
│   ├── build_app.py           # Python ビルドスクリプト
│   └── standarddocapp_launcher.py  # PyInstaller エントリ (絶対 import)
└── src/
    └── standarddocapp/
        ├── __init__.py
        ├── __main__.py        # python -m standarddocapp で GUI 起動
        ├── app.py             # Tk root, Notebook, メニューバー, ステータスバー
        ├── harvest_tab.py     # 収集タブ (Reload, 経過時間, 最新解決)
        ├── search_tab.py      # 検索タブ (Reload, 経過時間, 最新解決)
        ├── about_tab.py       # 設定・ログ・About タブ (サンプル/検証)
        ├── samples.py         # 3GPP/IEEE/Search のサンプル xlsx ジェネレータ
        ├── sample_dialog.py   # サンプル作成カスタムダイアログ (Toplevel)
        ├── html_check.py      # HTML 整合性チェッカ
        ├── widgets.py         # ScrollableFrame / StatusBar など共通ウィジェット
        ├── log_panel.py       # ScrolledText ベースのログ表示
        ├── osutil.py          # フォルダ/ファイルを開くヘルパ
        ├── paths.py           # アプリログディレクトリ
        ├── runner.py          # バックグラウンド実行 + logging キュー
        ├── sysinfo.py         # Office/LibreOffice/Python/OS 検出
        └── assets/
            ├── __init__.py
            ├── README.md      # アイコン差し替え手順
            └── app.ico        # 任意 (置けば自動適用)
```
