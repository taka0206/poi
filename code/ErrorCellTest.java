import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * CELL_TYPE_ERROR�̌��o
 */
class ErrorCellTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/ErrorBook.xls" : "./input/ErrorBook.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      //return;
    }
    String typeString[] = new String[] {"CELL_TYPE_NUMERIC", 
                                    "CELL_TYPE_STRING",
                                    "CELL_TYPE_FOMULA",
                                    "CELL_TYPE_BLANK",
                                    "CELL_TYPE_BOOLEAN",
                                    "CELL_TYPE_ERR"};
    // 1�Ԗڂ̃V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    // 1�s�ڂ̑I��
    Row row = sheet.getRow(0);
    // CellType�擾
    for (int i=0; i<3; i++) {
      System.out.println(i + "�Ԗڂ�CellType��" + typeString[row.getCell(i).getCellType()] + "�ł��B");
    }
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new ErrorCellTest().Run(args[0]);
  }
}
