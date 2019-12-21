import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * �p�X���[�h�t���u�b�N�̓ǂݍ���
 */
class BreakPassword {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    // �܂��A�p�X���[�h��ݒ肷��B
    Biff8EncryptionKey.setCurrentUserPassword("POI");
    FileInputStream fis = null;
    // ���Ƃ͕��ʂɃ��[�N�u�b�N��ǂݍ���
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/secret.xls" : "./input/secret.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    System.out.println("�閧�̃V�[�g[0]�ARow[0]�ACell[0]�ɂ́A\n�w" + 
              workBook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue() +
              "�x\n�Ə����Ă���܂����B");
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
    new BreakPassword().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
    
  }
}