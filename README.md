# get_3gpp_and_ieee_v3_202604

3GPP / IEEE 標準文書向けのツール群リポジトリです。

- `stdharvest`: ダウンロード・ZIP 解凍・PDF 化・HTML 化
- `stdsearch`: 生成済み HTML から指定語を検索・抽出
- `standarddocapp`: 上記 2 つを **1 つの Windows GUI（`StandardDocApp.exe`）**
  に統合した tkinter 製アプリ。エンドユーザーはこの exe だけを起動し、
  タブで「収集 / 検索」を切り替えて使えます。

| パッケージ | 役割 | セットアップ・使い方 | サンプル |
|---|---|---|---|
| `stdharvest` | 収集 (CLI / ライブラリ) | [`stdharvest/README.md`](./stdharvest/README.md) | [`stdharvest/samples/sample_download.xlsx`](./stdharvest/samples/sample_download.xlsx) |
| `stdsearch` | 検索 (CLI / ライブラリ) | [`stdsearch/README.md`](./stdsearch/README.md) | [`stdsearch/samples/sample_search.xlsx`](./stdsearch/samples/sample_search.xlsx) |
| `standarddocapp` | GUI 統合アプリ | [`standarddocapp/README.md`](./standarddocapp/README.md) | （上記 2 つを GUI から呼び出し） |

---

## クイックスタート (`StandardDocApp` GUI)

リポジトリ直下で次を実行します。`stdharvest` / `stdsearch` / `standarddocapp`
の **3 つを editable install** する必要があります。

```bat
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e .\stdharvest -e .\stdsearch -e .\standarddocapp
python -m standarddocapp
```

PowerShell の場合は `.venv\Scripts\Activate.ps1` を使ってください。スクリプト
実行ポリシーで弾かれる場合は次のいずれかで通します。

```powershell
# 一度だけ (CurrentUser スコープ; 管理者権限不要)
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# その場限り
powershell -NoProfile -ExecutionPolicy Bypass
```

> GUI を使う場合でも、既存の CLI（`stdharvest run --excel ...` /
> `stdsearch run --excel ...`）はそのまま利用可能です。

---

## エンドユーザー向け 1 ファイル exe を作る

リポジトリ直下の **`build_exe.bat`** をダブルクリック、または cmd / PowerShell
から実行するだけです（中身は `standarddocapp\build_exe.bat` を呼ぶだけの薄い
ラッパで、PowerShell の実行ポリシーには引っかかりません）。

```bat
build_exe.bat
```

成果物: `standarddocapp\dist\StandardDocApp.exe` （`--windowed` ビルドのため
ダブルクリックしても黒いコンソールは出ません）。

`build_exe.bat` は内部で次を行います。

1. `pip install --upgrade pip`
2. `pip install -e .\stdharvest -e .\stdsearch -e .\standarddocapp pyinstaller`
3. 旧 `build/` `dist/` の削除
4. `python -m PyInstaller --noconfirm --clean StandardDocApp.spec`

`.exe` のアイコンを差し替えたい場合は
`standarddocapp\src\standarddocapp\assets\app.ico` に `.ico` を置いてから
再ビルドするだけです。spec / 手動コマンド・トラブルシュートを含む詳細は
[`standarddocapp/README.md`](./standarddocapp/README.md) のビルド章を参照してください。

---

## クイックスタート (`stdharvest` CLI)

リポジトリ直下を例にします。

1. **仮想環境の作成**（未作成の場合）

```bat
python -m venv .venv
```

2. **仮想環境の有効化**

```bat
REM cmd
.venv\Scripts\activate.bat

REM PowerShell
.venv\Scripts\Activate.ps1
```

Linux / macOS の場合は `source .venv/bin/activate` です。

3. **仮想環境内で** `stdharvest` をインストールして実行

```bat
cd stdharvest
python -m pip install -e .
stdharvest run --excel samples/sample_download.xlsx
```

4. **仮想環境の無効化**

```bat
deactivate
```

> Windows で Microsoft Office (Word / PowerPoint / Excel) が利用可能な場合は
> Office 経由で PDF 化し、利用不可の場合は LibreOffice をフォールバック利用します。
> 実行前に Excel / Word / PowerPoint は保存して閉じてください。

---

## クイックスタート (`stdsearch` CLI)

リポジトリ直下を例にします。

1. **仮想環境の作成**（未作成の場合）

```bat
python -m venv .venv
```

2. **仮想環境の有効化**

```bat
REM cmd
.venv\Scripts\activate.bat

REM PowerShell
.venv\Scripts\Activate.ps1
```

3. **仮想環境内で** `stdsearch` をインストールして実行

```bat
cd stdsearch
python -m pip install -e .
stdsearch run --excel samples/sample_search.xlsx
```

4. **仮想環境の無効化**

```bat
deactivate
```
