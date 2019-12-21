=begin
  JRuby+POI�ɂ��h����ɂ��́B���E�h
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

#�R�}���h���C�������̃`�F�b�N
if ARGV.length != 1 then
  puts 'Parameter Error!!'
  exit
end
mode = ARGV[0]
if mode != '2003' && mode != '2007' then
  puts 'Mode is 2003 or 2007'
  exit
end

#���[�N�u�b�N�̐���
if mode == '2003' then
  workBook = HssfUserModel::HSSFWorkbook.new
else
  workBook = XssfUserModel::XSSFWorkbook.new
end

#�V�[�g�̐���
sheet = workBook.createSheet('HelloWorld')

#row�̐���
row = sheet.createRow(0)

#cell�̐���
cell = row.createCell(0)

#cell�X�^�C���̐���
st = workBook.createCellStyle()

#�t�H���g�̐���
fnt = workBook.createFont()
fnt.setFontName("�l�r ����")
fnt.setFontHeightInPoints(48)
fnt.setColor(HssfUtil::HSSFColor::AQUA::index)

#cell�X�^�C���Ƀt�H���g�ݒ�
st.setFont(fnt)

#cell�ɃX�^�C���ݒ�
cell.setCellStyle(st)

#cell�ɒl��ݒ�
cell.setCellValue('Hello World On JRuby��')

#���[�N�u�b�N�����o��
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
