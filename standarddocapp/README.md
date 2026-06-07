# standarddocapp

`stdharvest`（収集）と `stdsearch`（検索）を 1 つの GUI から実行するための
Windows 用 tkinter アプリ。エンドユーザーは **`StandardDocApp.exe`** だけを
起動し、画面上のタブで「収集 / 検索」を切り替えて利用します。

- 入口は **`StandardDocApp.exe`** に統一（`stdharvest.exe` / `stdsearch.exe`
  は作成しません）
- 既存の CLI（`stdharvest run --excel ...` / `stdsearch run --excel ...`）は
  そのまま利用可能
- PDF 化は **Microsoft Office を優先**し、不可なら **LibreOffice (soffice)**
  にフォールバック。Playwright / Chromium には依存しません
- PyInstaller で **単一 exe (`StandardDocApp.exe`) として配布可能**（配布先に
  Python 不要）

---

## 概要

このアプリは、標準文書のダウンロード・PDF 化・HTML 化・検索を GUI から
実行するためのツールです。

関連ツール:

- `stdharvest`: Excel から標準文書を一括取得し、PDF / HTML を生成する
- `stdsearch`: `stdharvest` が生成した HTML を検索し、CSV / HTML / JSON に出力する

---

## ディレクトリ構成

```text
repo-root/
├─ build_exe.bat                    # ルートから呼ぶラッパー
├─ tools/
│  └─ check_icon.py                 # app.ico サイズ検証 (Pillow)
├─ stdharvest/
├─ stdsearch/
└─ standarddocapp/
   ├─ build_exe.bat                 # 実際のビルドスクリプト (0/5〜5/5)
   ├─ StandardDocApp.spec           # PyInstaller 設定
   ├─ README.md
   ├─ pyproject.toml
   ├─ src/
   │  └─ standarddocapp/
   │     ├─ __init__.py
   │     ├─ __main__.py             # 絶対 import で書かれたエントリ
   │     ├─ app.py                  # Tk root, Notebook, アイコン適用
   │     ├─ harvest_tab.py
   │     ├─ search_tab.py
   │     ├─ about_tab.py
   │     ├─ samples.py
   │     ├─ sample_dialog.py
   │     ├─ html_check.py
   │     ├─ widgets.py
   │     ├─ log_panel.py
   │     ├─ osutil.py
   │     ├─ paths.py
   │     ├─ runner.py
   │     ├─ sysinfo.py
   │     └─ assets/
   │        ├─ __init__.py
   │        ├─ README.md
   │        ├─ app.ico              # ① ファイル / ② ウィンドウ / ③ タスクバー兼用
   │        └─ app.png              # ロゴ等の参考用
   ├─ build/                        # PyInstaller 作業ディレクトリ (生成物)
   └─ dist/                         # PyInstaller 出力先 (StandardDocApp.exe)
```

---

## 1. セットアップ (Windows / cmd or PowerShell)

リポジトリ直下で次を実行します。`stdharvest` / `stdsearch` / `standarddocapp`
の **3 つを editable install** する必要があります。

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e .\stdharvest -e .\stdsearch -e .\standarddocapp
```

PowerShell の場合は `.venv\Scripts\Activate.ps1` を使用。スクリプト実行ポリシー
で弾かれる場合は次のいずれか。

```powershell
# 一度だけ (CurrentUser スコープ; 管理者権限不要)
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# その場限り
powershell -NoProfile -ExecutionPolicy Bypass
```

---

## 2. 開発環境での起動

```bat
python -m standarddocapp
```

または `pip install -e .\standarddocapp` 後にショートカットコマンドで:

```bat
StandardDocApp
```

起動するとタブ付きの単一ウィンドウが開き、「収集 / 検索 / 設定・ログ」を
切り替えて使えます。

---

## 3. exe のビルド

リポジトリルートの `build_exe.bat` を実行します。

```bat
build_exe.bat
```

または `standarddocapp` フォルダで直接:

```bat
cd standarddocapp
build_exe.bat
```

`build_exe.bat` は次の 6 ステップを実行します。

| ステップ | 内容 |
|---|---|
| `[0/5] Checking required files` | `StandardDocApp.spec` / `__main__.py` / `assets\app.ico` の存在確認。1 つでも欠けると即エラーで停止 |
| `[1/5] Installing required packages` | `pip install -e ..\stdharvest -e ..\stdsearch -e . pyinstaller pillow` |
| `[2/5] Validating app.ico` | `tools\check_icon.py` で **16/24/32/48/64/128/256** が揃っているか検査 |
| `[3/5] Removing old build / dist` | 旧 `build/` `dist/` を削除 |
| `[4/5] Running PyInstaller` | `python -m PyInstaller --noconfirm --clean StandardDocApp.spec` |
| `[5/5] Checking output exe` | `dist\StandardDocApp.exe` が生成されたかを確認 |

### ビルド成果物

```text
standarddocapp/
└─ dist/
   └─ StandardDocApp.exe        # --windowed (コンソール非表示)
```

`dist\StandardDocApp.exe` を **そのまま 1 ファイルだけ** 渡せば配布完了です。

### spec を使わず手動コマンドでビルドする場合

```bat
cd standarddocapp
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name StandardDocApp ^
  --paths src ^
  --paths ..\stdharvest\src ^
  --paths ..\stdsearch\src ^
  --collect-all lxml ^
  --collect-all openpyxl ^
  --collect-all mammoth ^
  --collect-all bs4 ^
  --collect-all win32com ^
  --collect-all PIL ^
  --collect-submodules stdharvest ^
  --collect-submodules stdsearch ^
  --collect-submodules standarddocapp ^
  --hidden-import requests ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32com.client ^
  --add-data "src\standarddocapp\assets;standarddocapp\assets" ^
  --icon src\standarddocapp\assets\app.ico ^
  src\standarddocapp\__main__.py
```

---

## 4. アイコンについて

### 4.1 アプリのアイコンには 3 種類ある

| # | 場所 | 設定経路 |
|---|---|---|
| ① | エクスプローラのファイルアイコン | `StandardDocApp.spec` の `EXE(icon=str(ICON_PATH))` で焼き込み |
| ② | アプリのタイトルバーアイコン | `app.py` の `_apply_window_icon()` が `_MEIPASS\standarddocapp\assets\app.ico` を読み、`root.iconbitmap(default=...)` |
| ③ | Windows のタスクバーアイコン | ② と `SetCurrentProcessExplicitAppUserModelID("StandardDocApp.<version>")` |

3 つすべてが **同じ `app.ico` 1 ファイル** から自動的に設定されるように
なっています。差し替え時に編集する spec / コードはありません。

### 4.2 配置場所

```text
standarddocapp/src/standarddocapp/assets/app.ico
```

`app.png` を併置しておくと、後から ImageMagick で `.ico` を再生成する際の
元データになります（こちらも spec によりバンドルされます）。

### 4.3 必須サイズ

`app.ico` は次のサイズを **すべて** 含むマルチサイズ ICO にしてください
（`tools\check_icon.py` で検査）。可能なら全 32bpp。

- 16x16
- 24x24
- 32x32
- 48x48
- 64x64
- 128x128
- 256x256

PNG しか手元に無い場合は ImageMagick で生成可能:

```bat
magick convert app.png -define icon:auto-resize=256,128,64,48,32,16 app.ico
```

### 4.4 アイコン妥当性チェック

```bat
python -m pip install pillow
python tools\check_icon.py
```

成功時の出力例:

```text
ICO path:        ...\standarddocapp\src\standarddocapp\assets\app.ico
format:          ICO
primary size:    (256, 256)
available sizes: [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
OK: app.ico contains all recommended sizes (16/24/32/48/64/128/256).
```

`build_exe.bat` の `[2/5]` で同じ検査を行っているため、サイズ不足の `.ico`
を置いたままビルドが通ることはありません。

### 4.5 焼き込み確認 (PowerShell)

ビルド後、exe にアイコンが本当に焼き込まれているかを抽出して画像保存:

```powershell
Add-Type -AssemblyName System.Drawing
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon(
    (Resolve-Path .\standarddocapp\dist\StandardDocApp.exe).Path)
$icon.ToBitmap().Save("$env:TEMP\extracted.png",
    [System.Drawing.Imaging.ImageFormat]::Png)
explorer "$env:TEMP\extracted.png"
```

ここで自分のアイコンが取り出せれば「① ファイルアイコン」は焼き込み済みです。
② ③ は exe を起動してウィンドウとタスクバーで実機確認します。

---

## 5. ビルド後の確認チェックリスト

- [ ] `dist/StandardDocApp.exe` が作成される
- [ ] exe をダブルクリックして GUI が起動する
- [ ] Python 未起動状態でも exe 単体で起動する
- [ ] コンソールウィンドウが表示されない
- [ ] exe ファイルのアイコンが `app.ico` になっている (① エクスプローラ)
- [ ] GUI ウィンドウ左上のアイコンが `app.ico` になっている (② タイトルバー)
- [ ] タスクバーのアイコンが `app.ico` になっている (③ タスクバー)
- [ ] `stdharvest` / `stdsearch` の処理呼び出しで ImportError が出ない
- [ ] `assets/app.ico が見つからない` という警告が出ない
- [ ] `dist` フォルダを別の場所にコピーしても起動する

---

## 6. Windows のアイコンキャッシュ対策

ビルドが成功していても、Windows Explorer が古いアイコンを表示し続ける
ことがあります。これは **表示キャッシュの問題** であり、exe へのアイコン
埋め込み失敗とは別問題です。まずは以下のいずれかで切り分けてください。

- **exe を別フォルダにコピーして表示を確認する**
- **「§ 4.5 焼き込み確認」で焼き込みは正常か確認する**

その上で表示が古いままなら、アイコンキャッシュをリフレッシュします。

```bat
ie4uinit.exe -show
```

それでも変わらない場合は、フルクリア（再ログイン推奨）:

```bat
taskkill /IM explorer.exe /F
DEL /A /Q "%localappdata%\IconCache.db"
DEL /A /F /Q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"
start explorer.exe
```

---

## 7. PDF 生成について

本アプリでは、PDF 生成に **Playwright / Chromium は使用しません**。
`stdharvest/pdf_converter.py` が以下の優先順位で変換を試みます。

1. **Microsoft Office (Word / PowerPoint / Excel)** が利用可能なら COM 経由で PDF 化
2. 不可な場合は **LibreOffice (`soffice`)** にフォールバック

そのため `python -m playwright install chromium` 等の追加コマンドは不要です。
実行前に Excel / Word / PowerPoint は **保存して閉じてから** 走らせてください。

### Word/PowerPoint 内画像の HTML 表示について (EMF/WMF 対策)

Word 文書（特に 3GPP）には、グラフや数式が **EMF / WMF（Windows メタファイル）**
で貼り付けられていることが多く、そのままでは HTML（ブラウザ）で画像が表示
できません。本アプリは Word を簡易 HTML 本文化する際、文書内の **EMF / WMF
などブラウザ非対応の画像を自動的に PNG へ変換**して埋め込みます（Windows 上の
Pillow による描画。150 dpi）。PNG / JPEG など元々ブラウザ対応の画像はそのまま
埋め込みます。PowerPoint はスライド全体を PNG 画像として書き出すため、元々
正しく表示されます。

> 画像変換には `pillow` を使用します（`stdharvest` の依存に追加済み。
> editable install / exe ビルドのいずれでも自動的に導入されます）。

### 配布先 PC の前提条件

| 項目 | 必要条件 |
|---|---|
| OS | Windows 10 / 11 (64-bit) |
| Python | 不要 (exe に同梱) |
| Chromium / ブラウザ | 不要 (PDF は Office / LibreOffice で生成) |
| PDF 化 | Microsoft Office または LibreOffice (`soffice`) のいずれか |
| 日本語フォント | 不要 (各 Office が自前で処理) |
| .NET / Visual C++ ランタイム | 通常の Windows なら追加不要 |

---

## 7.5 プロキシ環境での利用 (認証付きプロキシ / 407 対策)

### 基本: 認証付きプロキシでも「設定なし」で通ります（Windows 統合認証）

社内・庁内などの **認証付きプロキシ** でも、多くの場合 **ProxyUser /
ProxyPassword を設定する必要はありません**。本アプリは Windows 上では
**ブラウザと同じ仕組み（WinHTTP の `WinHttpRequest`）でダウンロード**し、
プロキシに対して **現在ログオン中の Windows アカウントで自動認証
（NTLM / Kerberos のシングルサインオン）** します。

- `ProxyURL` を **空欄** にすれば、`netsh winhttp show proxy` → IE/WinINET の
  順でシステムのプロキシを自動検出し、統合認証で接続します。
- `ProxyURL` を明示指定した場合（例 `m7adc99proxy.ring.meti.go.jp:8080`）も、
  ユーザー名・パスワード未設定なら統合認証で接続します。
- 実行時ログに `プロキシ ... へ Windows 統合認証 (現在のログオンユーザー) で
  接続します` と表示されれば、この経路が使われています。

> 以前 `netsh` や MSXML2 / WinHTTP 系のツールで「特に設定せずに」ダウンロード
> できていたのと同じ挙動です。アカウント名やパスワードを Excel に書く必要は
> ありません。

### 統合認証で通らない場合のみ: ユーザー名 / パスワードを設定

統合認証が拒否される（ログに 407 が残る）特殊なプロキシでは、Sheet2
(Settings) に **任意で** 認証情報を設定できます。

| Sheet2 ラベル | 例 | 説明 |
|---|---|---|
| `ProxyURL` | `m7adc99proxy.ring.meti.go.jp:8080` | プロキシのホスト:ポート。`http://` は省略可（自動補完）。空欄なら自動検出 |
| `ProxyUser` | `taro.tokkyo` または `DOMAIN\taro.tokkyo` | プロキシ認証のユーザー名。**任意** |
| `ProxyPassword` | `********` | プロキシ認証のパスワード。**任意** |

- **`ProxyUser` / `ProxyPassword` を設定すると統合認証は使わず**、その
  ユーザー名・パスワードで（Basic 認証として）接続します。
- `ProxyUser` / `ProxyPassword` は Sheet2 の **16・17 行目**（`CombineHtmlBatchSize`
  の下）に追加されました。古い `sample_download.xlsx` には欄が無いので、その場合は
  「収集」タブの **サンプル Excel 作成** で作り直すか、Sheet2 の A16=`ProxyUser` /
  A17=`ProxyPassword` を手で追記して B 列に値を入れてください。
- ユーザー名・パスワードに `@` `:` `\` などの記号が含まれていても、内部で
  自動的に URL エンコードして `http://user:pass@host:port` 形式に組み立てます。
  自分で `ProxyURL` に `http://user:pass@...` を直接書いても構いません。
- 認証失敗 (407) のときは無駄なリトライをせず、即座にエラーとして
  Excel の Message 欄に対処方法を表示します。
- **`ProxyURL` を空欄にすると、Windows のシステムプロキシ設定を自動検出します。**
  まず `netsh winhttp show proxy` を確認し、未設定なら IE/WinINET
  （インターネット オプション）のプロキシ設定を参照します。検出した場合は
  ログに「システムのプロキシ設定を自動検出して使用します: host:port」と表示します。
  自動検出されたプロキシが認証を要求する場合でも、`ProxyUser` / `ProxyPassword`
  を設定すればそのまま使えます。
- 環境変数 (`HTTPS_PROXY` 等) を使いたい場合も `ProxyURL` を空欄にしてください
  （自動検出で何も見つからなければ requests が環境変数を参照します）。

> パスワードは Excel に平文で保存されます。共有フォルダに置く場合は取り扱いに
> 注意してください。

---

## 8. よくある問題 (トラブルシューティング)

| 症状 | 対処 |
|---|---|
| `ImportError: attempted relative import with no known parent package` | `src\standarddocapp\__main__.py` が **絶対 import** であることを確認（`from standarddocapp.app import launch`）。spec のエントリも `__main__.py` を直接渡す形で問題ありません |
| ダウンロードが全件失敗し `ProxyError ... 407 authenticationrequired` | 認証付きプロキシです。Sheet2 の `ProxyUser` / `ProxyPassword` にプロキシ用のユーザー名・パスワードを設定（§ 7.5 参照） |
| `ProxyError ... Unable to connect to proxy` (407 以外) | `ProxyURL` のホスト名・ポートを確認。プロキシ不要なら `ProxyURL` を空欄に |
| PowerShell で `スクリプトの実行が無効` / `UnauthorizedAccess` | `build_exe.bat`（cmd プロンプト）から起動。または `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` を一度だけ実行 |
| `pip install -e .` が失敗する | `python -m pip install --upgrade pip` 後に再実行 |
| `ModuleNotFoundError: stdharvest` / `stdsearch` (exe 実行時) | 兄弟パッケージを editable install していない。`build_exe.bat` を使えば自動で行われます |
| `ModuleNotFoundError: lxml.xxx` (exe 実行時) | spec を使うか、手動コマンドの場合は `--collect-all lxml` を追加 |
| `ModuleNotFoundError: win32com.gen_py` (exe 実行時) | spec を使うか、`--collect-all win32com` を追加 |
| exe のファイルアイコンが古いまま | Windows のアイコンキャッシュ。`ie4uinit.exe -show`、または `dist\StandardDocApp.exe` を別フォルダにコピーして表示確認 |
| ウィンドウのタイトルバーだけ Tk ロゴ | `assets\app.ico` が無い、または spec の `datas` に含まれていない。`build_exe.bat` で再ビルド |
| `assets/app.ico が見つからない` ログが出る | `standarddocapp\src\standarddocapp\assets\app.ico` の配置を再確認 |
| PyInstaller ビルドが途中で失敗する | 旧成果物を削除して再ビルド: `cd standarddocapp & rmdir /s /q build & rmdir /s /q dist & python -m PyInstaller --noconfirm --clean StandardDocApp.spec` |

---

## 9. アンインストール

開発環境を消したい場合:

```bat
python -m pip uninstall standarddocapp stdsearch stdharvest
```

配布先 (exe のみ) では、`StandardDocApp.exe` を削除するだけで完了です。
`%LOCALAPPDATA%\StandardDocApp` に残ったログも併せて削除すると完全に
クリーンな状態に戻ります。

---

## 10. 重要な設計方針 (開発者向け)

> **PDF 生成のためにブラウザエンジンを後からインストールさせる方式は採用していません。**
> Office / LibreOffice の COM / CLI を使い、Playwright / Chromium / Selenium
> などの依存を再導入しないでください。

> **`__main__.py` の import は絶対 import を維持してください。**
> 相対 import に戻すと PyInstaller での onefile ビルドが
> `ImportError: attempted relative import with no known parent package`
> で失敗します。`python -m standarddocapp` でも絶対 import は問題なく動作します。

> **アイコンは `src\standarddocapp\assets\app.ico` 1 ファイルに集約。**
> ① ファイル / ② ウィンドウ / ③ タスクバーすべてがここを起点に設定されます。
> 差し替え時に spec / Python コードを編集する必要はありません。

> **`StandardDocApp.spec` は単一のソース・オブ・トゥルース。**
> ビルド条件を変える場合は spec を編集してください。手動コマンドは
> 「spec が無い場合の代替手段」として用意されているだけで、CI で配布物を
> 作る際は必ず spec 経由でビルドしてください。
