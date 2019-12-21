import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * �V�[�g�����e�X�g
 */
public class SheetIteratorTest {

  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {

    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g��3������ 
    for (int i=0; i<3; i++) {
      workBook.createSheet();
    }
    if (mode.equals("2007")) {
      // Sheet���C�e���[�^�[�ŏ��� 
      for(XSSFSheet sheet : (XSSFWorkbook)workBook) {
        // �V�[�g����\��
        System.out.println(sheet.getSheetName());
      }
    }
    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if (!args[0].equals("2007")) {
      System.out.println("�G���[�F���[�h��2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new SheetIteratorTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
