/**
 *  Groovy+POIによる”こんにちは。世界”
 */

import org.apache.poi.util.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.hssf.util.*;
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.*;

// コマンドライン引数のチェック
if (args.length!=1) {
    println("パラメーターエラーです。")
    return
}
def mode=args[0]
if (!mode.equals("2003") && !mode.equals("2007")) {
    println("パラメーターは、2003 か 2007を指定してください。")
    return
}
// ワークブックの生成
def workBook;
if(mode.equals("2003")) {
    workBook = new HSSFWorkbook()
}
else {
    workBook = new XSSFWorkbook()
}


// シートの生成
def sheet = workBook.createSheet("HelloWorld")

// rowの生成
def row = sheet.createRow(0)

// cellの生成
def cell = row.createCell(0)

// cellスタイルの生成
def st = workBook.createCellStyle()

// フォントの生成
def fnt = workBook.createFont()
fnt.setFontName("ＭＳ 明朝")
fnt.setFontHeightInPoints((short)48)
fnt.setColor((short)HSSFColor.AQUA.index)

// cellスタイルにフォント設定
st.setFont(fnt)

// cellにスタイル設定
cell.setCellStyle(st)

// Cellに値を設定
cell.setCellValue("Hello World On Groovy♪")

// ワークブック書き出し
def fout
if (mode.equals("2003")) {
 fout = new java.io.FileOutputStream("./HelloWorld-Groovy.xls")
}
else {
 fout = new java.io.FileOutputStream("./HelloWorld-Groovy.xlsx")
}
workBook.write(fout)
fout.close()

println("done!")
