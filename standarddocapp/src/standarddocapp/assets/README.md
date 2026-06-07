# standarddocapp/assets

`app.ico` を 1 つここに置くだけで、`StandardDocApp.exe` の **3 つのアイコン**
（① エクスプローラのファイルアイコン / ② ウィンドウのタイトルバー /
③ Windows タスクバー）がすべて差し替わるようになっています。

| アイコンの種類 | 設定経路 |
|---|---|
| ① ファイルアイコン | `StandardDocApp.spec` が `app.ico` を `EXE(icon=...)` に渡して焼き込み |
| ② ウィンドウアイコン | `app.py` の `_apply_window_icon()` が `importlib.resources` 経由で `app.ico` を `root.iconbitmap(default=...)` に渡す |
| ③ タスクバーアイコン | ② と `SetCurrentProcessExplicitAppUserModelID("StandardDocApp.<version>")` で独立アプリ扱いに |

`spec` は `Path("...assets/app.ico").exists()` で自動判定するので、`.ico`
が無ければビルドはエラーにならず、PyInstaller デフォルトのアイコンになります。

## アイコンを差し替える手順

1. 任意の方法で `.ico` を用意します（推奨: **16/24/32/48/64/128/256** を含む
   マルチサイズ、全 32bpp）。PNG しか手元に無い場合は ImageMagick で生成可能:

   ```powershell
   magick convert app.png -define icon:auto-resize=256,128,64,48,32,16 app.ico
   ```

2. 出力した `.ico` を **このフォルダに `app.ico`** という名前で保存。
3. 再ビルド:

   ```bat
   build_exe.bat
   ```

`.ico` を置かない場合は、PyInstaller デフォルトの青い Python アイコンに
なります。

## ビルド後にアイコンが変わらないとき

ほぼすべて Windows Explorer のアイコンキャッシュが原因です。順に試してください。

```bat
REM 1. 軽いリフレッシュ
ie4uinit.exe -show

REM 2. キャッシュ全消し (再ログイン推奨)
taskkill /IM explorer.exe /F
del /A /Q "%LOCALAPPDATA%\IconCache.db"
del /A /F /Q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache_*.db"
start explorer.exe
```

exe そのものに埋め込まれているアイコンは PowerShell で確認できます。

```powershell
Add-Type -AssemblyName System.Drawing
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon(
    (Resolve-Path .\standarddocapp\dist\StandardDocApp.exe).Path)
$icon.ToBitmap().Save("$env:TEMP\extracted.png",
    [System.Drawing.Imaging.ImageFormat]::Png)
explorer "$env:TEMP\extracted.png"
```

ここで自分のアイコンが見えれば「① ファイルアイコン」は焼き込み済みです。
② ③ は実際にアプリを起動してウィンドウ・タスクバーで確認します。

## 任意のリソース

`app.ico` 以外のファイル（例: ロゴ PNG）も spec から
`standarddocapp/assets` にバンドルされるので、
`importlib.resources.files("standarddocapp.assets") / "logo.png"`
のように実行時に取り出せます。
