/**
 *  Groovy+POI�ɂ��h����ɂ��́B���E�h
 */

import org.apache.poi.util.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.hssf.util.*;
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.*;

// �R�}���h���C�������̃`�F�b�N
if (args.length!=1) {
    println("�p�����[�^�[�G���[�ł��B")
    return
}
def mode=args[0]
if (!mode.equals("2003") && !mode.equals("2007")) {
    println("�p�����[�^�[�́A2003 �� 2007���w�肵�Ă��������B")
    return
}
// ���[�N�u�b�N�̐���
def workBook;
if(mode.equals("2003")) {
    workBook = new HSSFWorkbook()
}
else {
    workBook = new XSSFWorkbook()
}


// �V�[�g�̐���
def sheet = workBook.createSheet("HelloWorld")

// row�̐���
def row = sheet.createRow(0)

// cell�̐���
def cell = row.createCell(0)

// cell�X�^�C���̐���
def st = workBook.createCellStyle()

// �t�H���g�̐���
def fnt = workBook.createFont()
fnt.setFontName("�l�r ����")
fnt.setFontHeightInPoints((short)48)
fnt.setColor((short)HSSFColor.AQUA.index)

// cell�X�^�C���Ƀt�H���g�ݒ�
st.setFont(fnt)

// cell�ɃX�^�C���ݒ�
cell.setCellStyle(st)

// Cell�ɒl��ݒ�
cell.setCellValue("Hello World On Groovy��")

// ���[�N�u�b�N�����o��
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
