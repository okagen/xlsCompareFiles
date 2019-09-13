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
  <img src="https://github.com/okagen/xlsCompareFiles/blob/master/Data/01-afterSetting.png" width="600">
  
  - The files are compared according to the settings when the button is clicked. If there are different parts, the cell with arrow mark in the column H is painted red. / ボタンを押すと、設定に従ってファイル同士が比較される。異なる部分がある場合、H列の矢印のセルが赤く塗りつぶされる。
  
  <img src="https://github.com/okagen/xlsCompareFiles/blob/master/Data/02-compare.png" width="600">
  
  - Refer to each sheet for details of the comparison results. / 比較した結果の詳細は、各シートを参照する。
  
    - In the case of there is the same value in the same cell. / 同じセルに同じ値がある場合。  
    <img src="https://github.com/okagen/xlsCompareFiles/blob/master/Data/03_result-1.png" width="600">  
  
    - The values in the header are ignored. / 異なる値が設定されていても、ヘッダであれば無視される。  
    <img src="https://github.com/okagen/xlsCompareFiles/blob/master/Data/04-result-2.png" width="600">  
  
    - If there are different values in the same cell position, the cell color turns red. / 同じセルの位置に異なる値があった場合、セルの色が赤くなる。
    <img src="https://github.com/okagen/xlsCompareFiles/blob/master/Data/05-result-3.png" width="600">
    
## Terminology

  - 突合（TOTSUGO）
      - 突き合わせて調べること。 / To check the ｍatching two or more materials together.
      
