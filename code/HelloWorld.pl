#!/usr/bin/perl
# 
# Perl+POI�ɂ��h����ɂ��́B���E�h
#
use strict;
use Java;

my $jv = new Java;

#�R�}���h�����`�F�b�N
if( @ARGV != 1){
	die "�g�p���@ $0 mode \n" 
}
my $mode = @ARGV[0];

unless ($mode eq "2003" || $mode eq "2007") {
	die "���[�h��2003��2007���w�肵�Ă�������\n"
} 
# ���[�N�u�b�N�̐���
my $workBook;

if ($mode eq "2003") {
  $workBook = $jv->create_object("org.apache.poi.hssf.usermodel.HSSFWorkbook");
}
else {
  $workBook = $jv->create_object("org.apache.poi.xssf.usermodel.XSSFWorkbook");
}
# �V�[�g�̐���
my $sheet = $workBook->createSheet("HelloWorld");

# Row�̐���
my $row = $sheet->createRow(0);

# cell�̐���
my $cell = $row->createCell(0);

#cell�X�^�C���̐���
my $st = $workBook->createCellStyle();

#�t�H���g�̐���
my $fnt = $workBook->createFont();
$fnt->setFontName('�l�r ����');

my $pnt = $jv->create_object('java.lang.Integer',"48");
$fnt->setFontHeightInPoints($pnt->shortValue());
my $cl = $jv->create_object('java.lang.Integer',"49");
$fnt->setColor($cl->shortValue());

#cell�X�^�C���Ƀt�H���g�ݒ�
$st->setFont($fnt);

#cell�ɃX�^�C���ݒ�
$cell->setCellStyle($st);

#cell�ɒl��ݒ�
$cell->setCellValue("Hello World On Perl��");

# ���[�N�u�b�N�����o��
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
