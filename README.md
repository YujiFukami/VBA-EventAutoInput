# VBA-EventAutoInput
# イベント機能活用による自動バックアップ用VBA

- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

その他、実行環境など報告していただくと感謝感激雨霰。

# 使い方

## 「実行サンプル 関連語句自動入力.xlsm」の使い方

「B列」セルの値を入力すると、入力された値の関連語句を「C列」セルに自動出力する。

「C列」セルの値を入力（または変更）すると、「B列」セルの値の関連語句として登録される。
![実行サンプル中身](https://user-images.githubusercontent.com/73621859/130750404-cfae07c8-f0f8-4ece-82ae-9f32cbe9a107.jpg)

使用デモ
https://user-images.githubusercontent.com/73621859/130791021-31265529-5f2c-49e9-9d4a-35bdff7bd6ff.mp4

## 設定

実行サンプル「実行サンプル 関連語句自動入力.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModEventAutoInput.bas
-  ModFile

### 設定2（参照ライブラリ）

-  特になし

### 設定3 (イベントプロシージャ設定)

　シートのセル値変更時実行時イベントプロシージャ「Worksheet_Change」プロシージャの中で、
「セルの値変更時に登録単語出力と単語登録」を実行させる。
　引数に「Target」を渡す。
![イベントプロシージャ設定](https://user-images.githubusercontent.com/73621859/130750342-8a148c22-baed-4989-8635-3b2e189c0a80.jpg)

 モジュール「ModAutoInput」の冒頭にての設定
-  入力値と関連語句を登録するテキストファイルの名前を定数「TextFileName」に設定する
-  「入力セル範囲取得」に入力セルの範囲を設定する
-  「出力セル範囲取得」に関連語句を出力するセルの範囲を設定する。
![ModAutoInputの設定](https://user-images.githubusercontent.com/73621859/130750229-12265e8b-1af3-4766-bc6b-10fe763b79fd.jpg)
