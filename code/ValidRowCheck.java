import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * �L���ȍs�𔻒肷��B
 */
class ValidRowCheck {
 
  /** �����̎��s
   * @param mode ���[�h
   */
  public void Run(String mode) {

    FileInputStream fis = null;
    Workbook workBook = null;
    // ���[�N�u�b�N�̓ǂݍ���
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/validrow.xls" : "./input/validrow.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
      return;
    }
    // �V�[�g���擾
    Sheet sheet = workBook.getSheetAt(0);

    System.out.println(
      "�L���s��" + (sheet.getFirstRowNum() + 1) + 
      "�s����" +
      (sheet.getLastRowNum() + 1) + "�s�܂łł��B");
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
    new ValidRowCheck().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
