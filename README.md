# 室長パトロール結果データの変換ツール
室長パトロール結果（PowerApp）をPowerAutomateで出力したファイルを提出用ファイルに成形する

## 環境構築
(1) Conda環境設定

`conda create -n geturei python`

`conda activate geturei`

(2) モジュールインストール

`pip install pandas`

`pip install openpyxl`

`pip install pywin32`

`pip install pypdf`

`pip install pyinstaller`

## exeファイル作成

「Exe作成」フォルダに`download.py`を保管し、以下のコマンドを実行

`pyinstaller download.py --onefile --noconsole --name ファイル変換 --icon image.ico`

(3) 動作

PowerAutomateを介してSharepointからコピーした`室長パトロール結果.csv`を読込み、`雛型フォルダ`内の`雛型.xlsx`に入力、`環境・バイオ研究室.xlsx`として同じフォルダ中に保存

※ 入力データ名は必ず`室長パトロール結果.csv`とする