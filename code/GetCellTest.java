import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * Cell�擾�e�X�g
 */
class GetCellTest {
  /** �����̎��s
   * @param ���[�h
   * @value1 �����l1
   * @value2 �����l2
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/calctest.xls" : "./input/calctest.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    Sheet sheet = workBook.getSheetAt(0);
    Row row = sheet.getRow(0);
    Cell cell = row.getCell(0,Row.RETURN_BLANK_AS_NULL);
    if (cell == null) {
      System.out.println("Cell�ɒl���ݒ肳��Ă��܂���B");
    }
    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 3) {
      System.out.println("�G���[�F���[�h���w�肳��Ă��܂���B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }

    new GetCellTest().Run(args[0]);
  }
}
