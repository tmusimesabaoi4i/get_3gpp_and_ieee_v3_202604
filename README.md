# get_3gpp_and_ieee_v3_202604

3GPP / IEEE 標準文書のダウンロード・PDF/HTML 化ツール `stdharvest` v1.0 の
リポジトリです。実体は [`stdharvest/`](./stdharvest/) 配下にあります。

- セットアップ・使い方: [`stdharvest/README.md`](./stdharvest/README.md)
- 入力サンプル: [`stdharvest/samples/sample_download.xlsx`](./stdharvest/samples/sample_download.xlsx)
- ソース: [`stdharvest/src/stdharvest/`](./stdharvest/src/stdharvest/)

クイックスタート:

```powershell
cd stdharvest
python -m pip install -e .
stdharvest run --excel samples/sample_download.xlsx
```

> LibreOffice がインストールされている環境で PDF 化が有効になります。
> また、実行前に Excel / Word / PowerPoint は保存して閉じてください。
