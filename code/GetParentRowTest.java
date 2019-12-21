import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �e��Row�擾�e�X�g
 */
public class GetParentRowTest {

  /**
   * Cell�ŗL�̏���
   *@param cell �Z���̎Q��
   */
  public void cellProc(Cell cell) {
    Row row = cell.getRow();
    System.out.println(
      "�e��Row�̗L��Cell��" + row.getFirstCellNum() +
      "����" + row.getLastCellNum() + "�܂łŁA\n" + 
      "���́A" + cell.getColumnIndex() + "�Ԗڂł��B");
  }
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�𐶐�
    Sheet sheet = workBook.createSheet();
    // Row�𐶐�
    Row row = sheet.createRow(0);
    // Cell��10�������A�l��ݒ�
    for (int i=3; i<13; i++) {
      row.createCell(i).setCellValue("�Z��" + i);
    }
    // Cell�ŗL�����Ăяo��
    cellProc(row.getCell(8));
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
    new GetParentRowTest().Run(args[0]);
    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}