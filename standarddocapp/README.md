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

## exe のビルド

PyInstaller で単一の Windows GUI 実行ファイルを生成します。

```powershell
cd <repo-root>
.\standarddocapp\build_tools\build.ps1
```

または Python から直接:

```powershell
cd <repo-root>
python .\standarddocapp\build_tools\build_app.py
```

成果物は `standarddocapp\dist\StandardDocApp.exe` に出力されます。共有先の
PC にコピーして起動してください（事前のセットアップ不要、Microsoft Office
または LibreOffice のいずれかが入っていれば PDF 化が動作します）。

`build.ps1` は内部で次を行います:
1. `.venv-build/` に隔離された venv を作成
2. `stdharvest`, `stdsearch`, `standarddocapp` を editable install
3. `pyinstaller` を最新化
4. `StandardDocApp.spec` を実行して `dist/StandardDocApp.exe` を出力

## ファイル構成

```
standarddocapp/
├── pyproject.toml
├── README.md
├── StandardDocApp.spec        # PyInstaller spec
├── build_tools/
│   ├── build.ps1              # PowerShell ラッパ
│   └── build_app.py           # Python ビルドスクリプト
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
        └── sysinfo.py         # Office/LibreOffice/Python/OS 検出
```
