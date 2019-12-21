# coding: utf-8
#
# Jython+POIによる”こんにちは。世界”
#
import sys
from java.io import *
from org.apache.poi.hssf.usermodel import *
from org.apache.poi.xssf.usermodel import *
from org.apache.poi.ss.usermodel import *
from org.apache.poi.hssf.util import *

#コマンドライン引数のチェック
if (len(sys.argv) != 2):
    print "Parameter Error!!"
    quit()

mode = sys.argv[1]
if (mode != "2003" and mode != "2007"):
    print "Mode is 2003 or 2007"
    quit()

#ワークブックの生成
if (mode == "2003") :
    workBook = HSSFWorkbook()
else:
    workBook = XSSFWorkbook()

#シートの生成
sheet = workBook.createSheet("HelloWorld")

#rowの生成
row = sheet.createRow(0)

#cellの生成
cell = row.createCell(0)

#cellスタイルの生成
st = workBook.createCellStyle()

#フォントの生成
fnt = workBook.createFont()
fnt.setFontName(u"ＭＳ 明朝")
fnt.setFontHeightInPoints(48)
fnt.setColor(HSSFColor.AQUA.index)

#cellスタイルにフォント設定
st.setFont(fnt)

#cellにスタイル設定
cell.setCellStyle(st)

#cellに値を設定
cell.setCellValue(u"Hello World On Jython♪")

#ワークブック書き出し
if (mode=="2003"):
    fout = FileOutputStream("./HelloWorld-Jython.xls") 
else:
    fout = FileOutputStream("./HelloWorld-Jython.xlsx") 

workBook.write(fout)
fout.close

print "done!"
