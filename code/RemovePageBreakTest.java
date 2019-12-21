import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �������y�[�W�����e�X�g
 */ 
public class RemovePageBreakTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet();

    // Row��30�sCell��15�񐶐�
    for (int i=0; i<30; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<15; j++) {
        row.createCell(j).setCellValue(i + "-" + j);
      }
    }
    // ���y�[�W�ʒu��10�s�A20�s��5��A10��ɐݒ�
    sheet.setRowBreak(9);
    sheet.setRowBreak(19);
    sheet.setColumnBreak(4);
    sheet.setColumnBreak(9);
    // ���y�[�W�ʒu(�s)�����ׂĉ���
    for(int breakLine : sheet.getRowBreaks()) {
      sheet.removeRowBreak(breakLine);
    }
    // ���y�[�W�ʒu(��)�����ׂč폜
    for(int breakCol : sheet.getColumnBreaks()) {
      sheet.removeColumnBreak(breakCol);
    }
    // �S�s�������Ă݂� - IllegalArgumentException������
    /*
    for (int i=0; i<=sheet.getLastRowNum(); i++) {
      sheet.removeRowBreak(i);
    }
    */
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./" + this.getClass().getName() + "_Book1.xls" : 
                      "./" + this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
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
    new RemovePageBreakTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
