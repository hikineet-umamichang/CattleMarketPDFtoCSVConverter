
# CattleMarketPDFtoCSVConverter

畜産生産実績一覧表（JA北海道情報センター標準帳票 ED1083T06）を見やすくするPythonスクリプトです。
distディレクトリのmain.exeをダウンロードするだけで使えます。

## Dependency

依存関係はrequirements.txtに記載しています。

## Usage

実行するとフォルダ選択画面が表示されるので、解析したいPDFが格納されているフォルダを選択してください。同じフォルダに必要情報を抜き出したCSVを格納します。

## Build

自分でビルドするときは下記のコマンドを実行してください

```
pyinstaller --onefile --noconsole --icon=cow.ico main.py
```

