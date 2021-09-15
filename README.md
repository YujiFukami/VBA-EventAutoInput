# VBA-EventAutoInput
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

実行環境など報告していただくと感謝感激雨霰。

# 説明
あらかじめ入力値とその関連語句を登録しておき、セルに入力値が入力されると関連語句が右のセルに自動入力される。
入力値と関連語句の登録はセルに入力しながら行える。


## 活用例
セルの手入力作業の高速化


# 使い方

「B列」セルの値を入力すると、入力された値の関連語句を「C列」セルに自動出力する。

「C列」セルの値を入力（または変更）すると、「B列」セルの値の関連語句として登録される。
![実行サンプル中身](https://user-images.githubusercontent.com/73621859/130750404-cfae07c8-f0f8-4ece-82ae-9f32cbe9a107.jpg)

使用デモ

[![関連語句自動入力　説明](http://img.youtube.com/vi/A5ttsYXCxqw/0.jpg)](http://www.youtube.com/watch?v=A5ttsYXCxqw)


## 設定
実行サンプル「Sample_EventAutoInput.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModEventAutoInput.bas

### 設定2（参照ライブラリ）
なし

### 設定3 (イベントプロシージャ設定)

　シートのセル値変更時実行時イベントプロシージャ「Worksheet_Change」プロシージャの中で、
「セルの値変更時に登録単語出力と単語登録」を実行させる。
　引数に「Target」を渡す。
 
![イベントプロシージャ設定](https://user-images.githubusercontent.com/73621859/130750342-8a148c22-baed-4989-8635-3b2e189c0a80.jpg)

 モジュール「ModAutoInput」の冒頭にての設定
-  入力値と関連語句を登録するテキストファイルの名前を定数「TextFileName」に設定する
-  「入力セル範囲取得」に入力セルの範囲を設定する
-  「出力セル範囲取得」に関連語句を出力するセルの範囲を設定する。

![ModAutoInputの設定](https://user-images.githubusercontent.com/73621859/130877369-38cd43ae-cf43-4195-adf3-9d5773a7943f.jpg)

