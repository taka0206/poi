=begin
  JRuby+POIによる”こんにちは。世界”
=end

require 'poi-3.7-20101029.jar'
require 'poi-ooxml-3.7-20101029.jar'
require 'xmlbeans-2.3.0.jar'
require 'poi-ooxml-schemas-3.7-20101029.jar'
require 'dom4j-1.6.1.jar'

include Java

module HssfUserModel
  include_package 'org.apache.poi.hssf.usermodel'
end

module HssfUtil
  include_package 'org.apache.poi.hssf.util'
end

module XssfUserModel
  include_package 'org.apache.poi.xssf.usermodel'
end

#コマンドライン引数のチェック
if ARGV.length != 1 then
  puts 'Parameter Error!!'
  exit
end
mode = ARGV[0]
if mode != '2003' && mode != '2007' then
  puts 'Mode is 2003 or 2007'
  exit
end

#ワークブックの生成
if mode == '2003' then
  workBook = HssfUserModel::HSSFWorkbook.new
else
  workBook = XssfUserModel::XSSFWorkbook.new
end

#シートの生成
sheet = workBook.createSheet('HelloWorld')

#rowの生成
row = sheet.createRow(0)

#cellの生成
cell = row.createCell(0)

#cellスタイルの生成
st = workBook.createCellStyle()

#フォントの生成
fnt = workBook.createFont()
fnt.setFontName("ＭＳ 明朝")
fnt.setFontHeightInPoints(48)
fnt.setColor(HssfUtil::HSSFColor::AQUA::index)

#cellスタイルにフォント設定
st.setFont(fnt)

#cellにスタイル設定
cell.setCellStyle(st)

#cellに値を設定
cell.setCellValue('Hello World On JRuby♪')

#ワークブック書き出し
if mode == '2003' then
  fout = java.io.FileOutputStream.new(
        './HelloWorld-JRuby.xls') 
else
  fout = java.io.FileOutputStream.new(
        './HelloWorld-JRuby.xlsx') 
end

workBook.write(fout)
fout.close

puts 'done!'
