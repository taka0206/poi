#!/usr/bin/perl
# 
# Perl+POIによる”こんにちは。世界”
#
use strict;
use Java;

my $jv = new Java;

#コマンド引数チェック
if( @ARGV != 1){
	die "使用方法 $0 mode \n" 
}
my $mode = @ARGV[0];

unless ($mode eq "2003" || $mode eq "2007") {
	die "モードは2003か2007を指定してください\n"
} 
# ワークブックの生成
my $workBook;

if ($mode eq "2003") {
  $workBook = $jv->create_object("org.apache.poi.hssf.usermodel.HSSFWorkbook");
}
else {
  $workBook = $jv->create_object("org.apache.poi.xssf.usermodel.XSSFWorkbook");
}
# シートの生成
my $sheet = $workBook->createSheet("HelloWorld");

# Rowの生成
my $row = $sheet->createRow(0);

# cellの生成
my $cell = $row->createCell(0);

#cellスタイルの生成
my $st = $workBook->createCellStyle();

#フォントの生成
my $fnt = $workBook->createFont();
$fnt->setFontName('ＭＳ 明朝');

my $pnt = $jv->create_object('java.lang.Integer',"48");
$fnt->setFontHeightInPoints($pnt->shortValue());
my $cl = $jv->create_object('java.lang.Integer',"49");
$fnt->setColor($cl->shortValue());

#cellスタイルにフォント設定
$st->setFont($fnt);

#cellにスタイル設定
$cell->setCellStyle($st);

#cellに値を設定
$cell->setCellValue("Hello World On Perl♪");

# ワークブック書き出し
my $out;
if ($mode eq "2003") {
  $out = $jv->create_object("java.io.FileOutputStream", "./HelloWorld-Perl.xls");
}
else {
  $out = $jv->create_object("java.io.FileOutputStream', './HelloWorld-Perl.xlsx");
}
$workBook->write($out);
$out->close();

print "done!";
