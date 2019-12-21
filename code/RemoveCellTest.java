import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Cell�폜�̃e�X�g
 */
public class RemoveCellTest {

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
    // Row�𐶐�
    Row row = sheet.createRow(0);
    // Cell��10����
    for (int i=0; i<10; i++) {
      row.createCell(i);
    }
    // �^�񒆂������Cell���폜
    row.removeCell(row.getCell(4));
    // Cell�̏�Ԃ��o��
    for (int i=0; i<10; i++) {
      Cell cell = row.getCell(i);
      if (cell != null) {
        System.out.print("��");
      }
      else {
        System.out.print("�~");
      }
    }
    System.out.println("");
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
    new RemoveCellTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
