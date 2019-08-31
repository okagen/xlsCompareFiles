# xlsCompareFiles
Compare CSV files or Excel files, then find differences. / CSVファイル、またはExcelファイルを突合して異なる個所を見つける。

## Sample direcotry structure.

~~~
root directory
│  xlsCompareFiles.xlsm
│
└─Data
    ├─Comp_A
    │      H19.csv
    │      H20.csv
    │      H21+noise.csv
    │
    └─Comp_B
            H19.csv
            H21.csv
~~~

## Usage.
  - 比較するファイルが保存されているディレクトリを入力する。
  - 比較するファイルの名前と、比較範囲を入力する。
  - ボタンを押すと、設定に従ってファイル同士が比較される。異なる部分がある場合、矢印ぬセルが赤く塗りつぶされる。
  - 比較した結果の詳細は、各シートを参照する。

