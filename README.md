# Aggregate-report

## 取り組みの目的
### 課題
- 現状集計業務はSSをExcelシートにDL→Excel上で加工→グラフ作成という作業方法のため工数がかかる
- 動作が不安定でたまにエラーが発生する
- グラフの色分けが見づらい
- 人に教えにくいなどが課題としてある。

### ゴール
Google Colaboratoryを使用して課題を解決して、業務改善を図る

## 業務フロー
| No  | As-is                                                                                                     | To-be                                            | 
| ---- | -------------------------------------------------------------------------------------------------------- | ------------------------------------------------ | 
| 1   | 週の終わりに翌週のTODOシートを追加する                                                                    | 週の終わりに翌週のTODOシートを追加する           | 
| 2   | 1日の中で自分の作業の実績を記入する                                                                       | 1日の中で自分の作業の実績を記入する              | 
| 3   | 週の始まりに先週のTODOシートをExcelにDL                                                                   | 週の始まりに先週分のTODOシートをExcelにDL        | 
| 4   | TODOシートの内容をINDIREC関数で転記                                                                       | 廃止                                             | 
| 5   | 4のシートをデータソースとして、別シートでピボットグラフが作成される                                       | データを指定のフォルダに格納し、スクリプトを実行 | 
| 5.1 | あらかじめデータソースの箱となるテーブルを指定しておく                                                    | 廃止                                             | 
| 5.2 | グラフに反映必須な項目<br>・標準作業時間の4週分の合計<br>・サービス<br>・中分類<br>・小分類<br>・実働時間 | 廃止                                             | 
| 5.3 | 「アドイン」の「レビュー指摘分析ツール」を使用して、5で作成したピボットグラフに各作業分類別に色を付ける   | 廃止                                             | 

## 機能要件

指定のグラフにデータを格納したら、標準作業時間の合計、サービス、中分類、小分類、実働時間を読み込む

1で読み込んだ情報を基に標準作業時間のみ折れ線グラフ、サービス、中分類、小分類、実働時間は積み上げ縦棒の複合グラフを作業者ごとに作成する

分類ごとに色を指定し、色付けする

## 詳細設計
1. agg.pyでフォルダ内の全てのエクセルをマージする
→opnepyxl
2. xx.pyでマージしたエクセルに対して、ピボット表、ピボットグラフを作成する
→pandas,matplotlib,seaborn


## 使用するランタイム/ライブラリ
- python3.9
- openxl：エクセルを操作するライブラリ
- pandas：表形式のデータを扱うためのライブラリ、マージしたエクセルを集計するために使用する
- matplotlib：グラフ化用のライブラリ
- japanize_matplotlib：matplotlibを日本語化するライブラリ
- seaborn：グラフを高機能にしたり、色を綺麗なものにするライブラリ

### 使用方法
1. ローカルへライブラリをインストール
pip install インストールしたいライブラリ
2. ソースコード内でモジュールをインポート
import openpyxl as pxみたいな

ライブラリーやモジュールについてはこちら
https://ai-inter1.com/python-module_package_library/

## その他
- ls：今いるディレクトリの配下にあるディレクトリやファイルを表示
- mkdir：ディレクトリを作成する
- cd：ディレクトリを移動する

## 参考にしたURL
Pythonで手軽にExcel操作【OpenPyXL】
https://zenn.dev/cppm/articles/3dd1bbc8b8720a

Pythonで複数のExcel・CSVを結合（マージ）する方法
https://www.ex-it-blog.com/Python-Excel-csv-merge

linuxコマンド頻出一覧
https://tech-blog.rakus.co.jp/entry/20210604/linux

gitの初期設定(windows)
https://prog-8.com/docs/git-env-win

vscodeでgit/githubを使う
https://xn--cmiya-system-works-5e4q.com/blog/detail/vscode-github/

Matplotlib & Seabornの使い方
https://kino-code.com/matplotlib_seaborn-04/
