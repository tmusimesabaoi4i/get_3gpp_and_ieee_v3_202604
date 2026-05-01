# standarddocapp/assets

PyInstaller の spec (`StandardDocApp.spec`) は、このフォルダにある
`app.ico` を見つけたら自動的に `StandardDocApp.exe` のアイコンに使います。

## アイコンを差し替える手順

1. 任意の方法で `.ico` を用意します（推奨: 256x256 までを含むマルチサイズ）。
   - PNG しか手元に無い場合は ImageMagick で:
     ```powershell
     magick convert app.png -define icon:auto-resize=256,128,64,48,32,16 app.ico
     ```
   - オンラインの ico コンバータでも可。
2. 出力した `.ico` を **このフォルダに `app.ico`** という名前で保存。
3. 再ビルド:
   ```powershell
   .\standarddocapp\build_tools\build.bat -Clean
   ```

`.ico` を置かない場合は、PyInstaller デフォルトのアイコンになります
（`.spec` 側で自動判定しているため、何もしなくてビルドは通ります）。

## 任意のリソース

`app.ico` 以外のファイル（例: ロゴ PNG）も spec から
`standarddocapp/assets` にバンドルされるので、
`importlib.resources.files("standarddocapp.assets") / "logo.png"`
のように実行時に取り出せます。
