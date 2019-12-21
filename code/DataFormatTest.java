import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �\�������̃e�X�g
 */
public class DataFormatTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    if (mode.equals("2003")) {
      // hssf(Excel2003�h�L�������g)�̏���
      // �V�[�g�̐��� 
      Sheet sheet = workBook.createSheet("�\���`���ꗗ");
      // DataFormat�C���X�^���X�̎Q�Ǝ擾
      HSSFDataFormat df = 
        (HSSFDataFormat)workBook.createDataFormat();
      // 1.�܂��r���h�C���\���`�����ꗗ���Ă݂�B
      System.out.println("�r���h�C���\���`���� = " + 
        HSSFDataFormat.getNumberOfBuiltinBuiltinFormats());
      // ���炩����Row��21�s����
      for (int i=0; i<21; i++) {
        sheet.createRow(i);
      }
      int rNum = 1;
      int cNum = 0;
      Row titleRow = sheet.createRow(0);
      titleRow.createCell(cNum).setCellValue("No");
      titleRow.createCell(cNum + 1).setCellValue(
        "�\���`��");
      for (int i=0; 
          i<HSSFDataFormat.getNumberOfBuiltinBuiltinFormats();
          i++) {
        sheet.getRow(rNum).createCell(cNum).setCellValue(i);
        sheet.getRow(rNum).createCell(cNum + 1).setCellValue(
          HSSFDataFormat.getBuiltinFormat((short)i));
        rNum++;
        if (rNum > 20) {
          rNum = 1;
          cNum += 2;
          titleRow.createCell(cNum).setCellValue("No");
          titleRow.createCell(cNum + 1).setCellValue(
            "�\���`��");
        }
      }
      // �Ō�ɗ񕝂������ݒ�ɂ���B
      for(int i=0; i<=cNum+1; i++) {
        sheet.autoSizeColumn(i);
      }
      // �\���`���ݒ�p�ɐV�����V�[�g�����
      Sheet sheet2 = workBook.createSheet("�\���`���ݒ�");
      // ���[�U�[�\���`�����쐬����B
      short nFormatDate = df.getFormat("yyyy�Nmm��dd��");
      System.out.println("�\���`���ԍ� = " + nFormatDate);
      Row fmtRow = sheet2.createRow(0);
      fmtRow.createCell(0).setCellValue(new Date());
      fmtRow.createCell(1).setCellValue(42.195);
      CellStyle styleNew = workBook.createCellStyle();
      CellStyle styleBuildin = workBook.createCellStyle();
      // 1��ڂ̓��[�U�[�\���`��
      styleNew.setDataFormat(nFormatDate);
      fmtRow.getCell(0).setCellStyle(styleNew);
      // 2��ڂ̓r���h�C���\���`��
      styleBuildin.setDataFormat(
        HSSFDataFormat.getBuiltinFormat("#,##0.00"));
      fmtRow.getCell(1).setCellStyle(styleBuildin);
      // �񕝂������ݒ�ɂ���B
      sheet2.autoSizeColumn(0);
      sheet2.autoSizeColumn(1);
    }
    else {
      // xssf(Excel2007�h�L�������g)�̏���
      // �\���`���ݒ�p�V�[�g�����B
      Sheet sheetX = workBook.createSheet("�\���`���ݒ�");
      // DataFormat�C���X�^���X�̎Q�Ǝ擾
      XSSFDataFormat df = 
        (XSSFDataFormat)workBook.createDataFormat();
      // ���[�U�[�\���`�����쐬����B
      short nFormatDate = df.getFormat("yyyy�Nmm��dd��");
      System.out.println("�\���`���ԍ� = " + nFormatDate);
      Row fmtRow = sheetX.createRow(0);
      fmtRow.createCell(0).setCellValue(new Date());
      fmtRow.createCell(1).setCellValue(42.195);
      CellStyle styleNew = workBook.createCellStyle();
      CellStyle styleBuildin = workBook.createCellStyle();
      // 1��ڂ̓��[�U�[�\���`��
      styleNew.setDataFormat(nFormatDate);
      fmtRow.getCell(0).setCellStyle(styleNew);
      // 2��ڂ̓r���h�C���\���`��
      styleBuildin.setDataFormat((short)4);
      fmtRow.getCell(1).setCellStyle(styleBuildin);
      // �񕝂������ݒ�ɂ���B
      sheetX.autoSizeColumn(0);
      sheetX.autoSizeColumn(1);
    }
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
      }
    }
    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new DataFormatTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
