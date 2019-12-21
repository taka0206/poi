import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �I�[�g�t�B���^�[�ݒ�e�X�g
 */ 
public class SetAutoFilterTest {

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
    // Row��10�sCell��10�񐶐�
    for (int i=0; i<10; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<10; j++) {
        row.createCell(j).setCellValue(i + "-" + j);
      }
    }
    // �I�[�g�t�B���^�[�ݒ� 4�s����6�s�A4�񂩂�7��
    // �������A�I���s�̐ݒ�͖��������B
    AutoFilter afil = sheet.setAutoFilter(
            new CellRangeAddress(3, 5, 3, 6));

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
    new SetAutoFilterTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}