# get_3gpp_and_ieee_v3_202604

3GPP / IEEE 標準文書向けのツール群リポジトリです。

- `stdharvest`: ダウンロード・ZIP 解凍・PDF 化・HTML 化
- `stdsearch`: 生成済み HTML から指定語を検索・抽出

- `stdharvest`:
  - セットアップ・使い方: [`stdharvest/README.md`](./stdharvest/README.md)
  - 入力サンプル: [`stdharvest/samples/sample_download.xlsx`](./stdharvest/samples/sample_download.xlsx)
  - ソース: [`stdharvest/src/stdharvest/`](./stdharvest/src/stdharvest/)
- `stdsearch`:
  - セットアップ・使い方: [`stdsearch/README.md`](./stdsearch/README.md)
  - 入力サンプル: [`stdsearch/samples/sample_search.xlsx`](./stdsearch/samples/sample_search.xlsx)
  - ソース: [`stdsearch/src/stdsearch/`](./stdsearch/src/stdsearch/)

クイックスタート (`stdharvest`):

リポジトリ直下（この `README.md` があるフォルダ）を例にします。

1. **仮想環境の作成**（未作成の場合）

```powershell
python -m venv .venv
```

2. **仮想環境の有効化（入室）**

```powershell
# PowerShell
.\.venv\Scripts\Activate.ps1
```

コマンドプロンプトの場合は `.\.venv\Scripts\activate.bat`、Linux / macOS の場合は `source .venv/bin/activate` です。

3. **仮想環境内で** `stdharvest` をインストールして実行

```powershell
cd stdharvest
python -m pip install -e .
stdharvest run --excel samples/sample_download.xlsx
```

4. **仮想環境の無効化（退室）**

```powershell
deactivate
```

> Windows で Microsoft Office (Word / PowerPoint / Excel) が利用可能な場合は
> Office 経由で PDF 化し、利用不可の場合は LibreOffice をフォールバック利用します。
> 実行前に Excel / Word / PowerPoint は保存して閉じてください。

クイックスタート (`stdsearch`):

リポジトリ直下（この `README.md` があるフォルダ）を例にします。

1. **仮想環境の作成**（未作成の場合）

```powershell
python -m venv .venv
```

2. **仮想環境の有効化（入室）**

```powershell
# PowerShell
.\.venv\Scripts\Activate.ps1
```

コマンドプロンプトの場合は `.\.venv\Scripts\activate.bat`、Linux / macOS の場合は `source .venv/bin/activate` です。

3. **仮想環境内で** `stdsearch` をインストールして実行

```powershell
cd stdsearch
python -m pip install -e .
stdsearch run --excel samples/sample_search.xlsx
```

4. **仮想環境の無効化（退室）**

```powershell
deactivate
```
