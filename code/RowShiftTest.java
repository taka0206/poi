import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row�ړ��̃e�X�g
 */
public class RowShiftTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    Sheet sheet = workBook.createSheet();
    // Row��10�s�����B
    for (int i=0; i<10; i++) {
      Row row = sheet.createRow(i);
      // 1�Ԗڂ�Cell��Row�ԍ���ݒ�
      row.createCell(0).setCellValue(i);
    }
    // �^�񒆂������2�s�폜
    sheet.removeRow(sheet.getRow(4));
    sheet.removeRow(sheet.getRow(5));
    // Row�̈ړ�
    sheet.shiftRows(6, 9, -2);
    // Row�ԍ���\��
    for (int i=0; i<sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        System.out.println(
          (int)(row.getCell(0).getNumericCellValue()));
      }
      else {
        System.out.println("null");
      }
    }
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
    new RowShiftTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
