import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * �s�������擾�e�X�g
 */
class GetMergedRegionTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ���[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? 
            "./input/MergedRegionTest_Book1.xls" : "./input/MergedRegionTest_Book1.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
    // Cell�����̈搔�擾
    System.out.println("�����̈��" + 
          sheet.getNumMergedRegions() + "�ӏ��ł��B");
    // Cell�������擾�ƕ\��
    for (int i=0; i<sheet.getNumMergedRegions(); i++) {
      CellRangeAddress range = sheet.getMergedRegion(i);
      System.out.println("�����̈�[" + i + "] �� " +
                     range.formatAsString());
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
    // �����̎��s
    new GetMergedRegionTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
