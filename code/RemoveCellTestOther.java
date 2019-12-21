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
    // Row��2�s�ACell��10�������B
    for (int i=0; i<2; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<10; j++) {
        row.createCell(j).setCellValue(i + "-" + j);
      }
    }
    Row r0 = sheet.getRow(0);
    Row r1 = sheet.getRow(1);
    // Cell�̎������o��
    System.out.println("Sheet0��Cell�� = " + r0.getPhysicalNumberOfCells());
    System.out.println("Sheet1��Cell�� = " + r1.getPhysicalNumberOfCells());
  
    Cell cell = r0.getCell(3); // 1�s�ڂ�Row����3�Ԗڂ�Cell���擾�B
    r1.removeCell(cell); // 2�s�ڂɑ΂���Cell���폜
    // �ēxCell�̎������o��
    System.out.println("Sheet0��Cell�� = " + r0.getPhysicalNumberOfCells());
    System.out.println("Sheet1��Cell�� = " + r1.getPhysicalNumberOfCells());
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
    new RemoveCellTest().Run(args[0]);
  }
}
