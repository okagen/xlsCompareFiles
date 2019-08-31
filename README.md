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
  - Enter the effective digit when comparing numbers. / 数字を比較する場合の、有効桁数を入力する。
  - Enter the path of directory where the files to be compared are stored. / 比較するファイルが保存されているディレクトリを入力する。
    - Source file directory path. / ソースファイルディレクトリ（比較元）
    - Destination file directory path. / 突合ファイルディレクトリ（比較先）
  - Enter the name of the file to be compared, the sheet to be compared and the comparison range. Also, enter the name of the sheet that outputs the comparison results. / 比較するファイルの名前と、比較するシートおよび比較範囲を入力する。また、比較結果を出力するシートの名前を入力する。
  - The files are compared according to the settings when the button is clicked. If there are different parts, the cell with arrow mark in the column H is painted red. / ボタンを押すと、設定に従ってファイル同士が比較される。異なる部分がある場合、H列の矢印のセルが赤く塗りつぶされる。
  - 比較した結果の詳細は、各シートを参照する。 / Refer to each sheet for details of the comparison results.

