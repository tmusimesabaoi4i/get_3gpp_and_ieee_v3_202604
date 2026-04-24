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

```powershell
cd stdharvest
python -m pip install -e .
stdharvest run --excel samples/sample_download.xlsx
```

> Windows で Microsoft Office (Word / PowerPoint / Excel) が利用可能な場合は
> Office 経由で PDF 化し、利用不可の場合は LibreOffice をフォールバック利用します。
> 実行前に Excel / Word / PowerPoint は保存して閉じてください。

クイックスタート (`stdsearch`):

```powershell
cd stdsearch
python -m pip install -e .
stdsearch run --excel samples/sample_search.xlsx
```
