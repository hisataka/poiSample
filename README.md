poiSample
======================
Excelファイルを取り込むために、POIを利用するサンプルです。  

取り込めるファイル形式
----------------------
・97-2003のxls形式  
・2007のxlsx形式  

2010に関してはサイト上に記載がないため、対象外  
（2007と同じxlsx形式なので、読み込むことはできる。）  
http://poi.apache.org/spreadsheet/index.html  

サンプルの実行方法
------------------
・sample.xlsを読み込む  
　(1)HSSFReadSample.classとsample.xlsを適当なフォルダにコピーする。  
　(2)次のコマンドを実行する。  
　　$>java -classpath .;C:\poi-3.7\*;C:\poi-3.7\lib\*;C:\poi-3.7\ooxml-lib\* HSSFReadSample sample.xls  

・sample.xlsxを読み込む  
　(1)XSSFReadSample.classとsample.xlsxを適当なフォルダにコピーする。  
　(2)次のコマンドを実行する。  
　　$>java -classpath .;C:\poi-3.7\*;C:\poi-3.7\lib\*;C:\poi-3.7\ooxml-lib\* XSSFReadSample sample.xlsx  

poi-bin-3.7-20101029.zipをCドライブ直下に解凍しておく必要がある。  
JREがインストールされており、javaコマンドが実行できるようパスが通っている必要がある。  

関連情報
--------
http://www.javadrive.jp/poi/  
http://www.visards.co.jp/java/poi/poi5.html  
http://www.fk.urban.ne.jp/home/kishida/kouza/poi/poi.html  