import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Cell�����q�̃e�X�g�B
 */
public class CellIteratorTest {
  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N��ǂݍ���
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/Iterator.xls" : "./input/Iterator.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    // Row�̎擾
    Row row = sheet.getRow(0);
    // �L����Cell�̂݃C�e���[�^�[�ŏ��� 
    for(Cell cell : row) {
      System.out.println(
        "Cell[" + cell.getColumnIndex() + 
        "] = " + cell.getNumericCellValue());
    }
    Iterator<Cell> it = row.iterator();
    while(it.hasNext()) {
      Cell cell = it.next();
      System.out.println(
        "Cell[" + cell.getColumnIndex() +
        "] = " + cell.getNumericCellValue());
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
    new CellIteratorTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
